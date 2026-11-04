"""กรอก Workday Enter Time จาก output/<ชื่อ>/preview_<เดือน>.xlsx (ผู้ใช้ตรวจ/แก้ไฟล์นี้ได้ก่อนรัน)

**ไม่กด Submit** กรอกเสร็จแล้วผู้ใช้ตรวจและกด Submit เอง

    python fill_workday.py --month 2026-10 --dry-run   เปิดฟอร์มและเลือกค่าให้ดู แล้วกด Cancel ไม่บันทึก
    python fill_workday.py --month 2026-10             กรอกจริง (กด OK ทีละ entry)
    python fill_workday.py --week 2026-10-26

ถ้ามี Edge ที่เปิด debug port 9222 อยู่ (จากรอบก่อน) จะต่อเข้าไปใช้ session เดิม
ไม่งั้นจะเปิด Edge แยก process ด้วย profile ใน wd-profile/<ชื่อ> ให้ผู้ใช้ login เอง สคริปต์จบแล้ว Edge ยังเปิดอยู่
ต้องตั้ง Workday เป็นภาษา English (ข้อความบนหน้าใช้จับ dialog และปุ่ม)
URL, ชื่อช่อง และการจับคู่ Role กับ Time Type มาจากหัวข้อ workday ใน team config
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
import os
import re
import shutil
import subprocess
import sys
import time
from dataclasses import dataclass
from pathlib import Path
from urllib.parse import urlparse

import openpyxl
from playwright.sync_api import Error as PlaywrightError
from playwright.sync_api import sync_playwright

from parse_schedule import (INHOUSE, OT, TIME_RE, _merge, day_segments, load_config, output_dir, parse_time_ranges,
                            person_of, require_team)

DIALOG = '[data-automation-id="popUpDialog"]'
MONTHS = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October",
          "November", "December"]

# ค่าตั้งต้นของหัวข้อ workday (ค่าของ tenant อยู่ใน team config)
WORKDAY_DEFAULTS = {
    "url": "",                     # เช่น https://<host>/<tenant>
    "enter_time_path": "",         # path ของหน้า Enter Time ต่อจาก url
    "browser_path": "",            # ว่าง = หา msedge.exe เอง
    "profile_dir": "wd-profile",   # แยกโฟลเดอร์ย่อยต่อคน
    "date_format": "%d/%m/%Y",     # วันที่บนหัว dialog ตามที่ตั้งใน Workday (US = %m/%d/%Y)
    "fields": {                    # ชื่อช่องใน dialog Enter Time
        "time_type": "Time Type",
        "calculation": "",         # ช่องที่ใช้ระบุ OT (ว่าง = ไม่มีช่องนี้ จึงกรอก OT ไม่ได้)
        "comment": "Comment",
        "hours": "Hours",
    },
    "project_task_by_role": {},
    "default_project_task": "",
    "inhouse_code_by_role": {},
    "overtime_calculation": "",
    "skip_leave_days": True,
}

# วันที่บนหัว dialog Enter Time ตาม Locale ในบัญชี Workday (ลอง workday.date_format ก่อน)
# d/m กับ m/d สลับกันได้เฉพาะวันที่ 1-12 แต่วันที่ที่สลับแล้วห่างกันเป็นเดือน คลิกถึงจากปฏิทินเดือนเดียวไม่ได้
DIALOG_DATE_FORMATS = ["%d/%m/%Y", "%m/%d/%Y", "%Y-%m-%d", "%Y/%m/%d", "%d.%m.%Y", "%d-%m-%Y",
                       "%d %B %Y", "%B %d, %Y", "%d %b %Y", "%b %d, %Y", "%d-%b-%Y"]
# ท้าย aria-label ของ calendarevent: "..., 15 October 2026" หรือแบบ US "..., October 15, 2026"
EVENT_DATE_RES = [re.compile(r"(?P<d>\d{1,2}) (?P<m>[A-Za-z]+) (?P<y>\d{4})$"),
                  re.compile(r"(?P<m>[A-Za-z]+) (?P<d>\d{1,2}), (?P<y>\d{4})$")]

# หาจุดว่างในช่องวันที่ที่ไม่ทับ entry/ปุ่ม เพื่อคลิกเปิด dialog Enter Time
FREE_SPOT_JS = """
cell => {
  const r = cell.getBoundingClientRect();
  for (const fy of [0.55, 0.65, 0.75, 0.45, 0.85, 0.35])
    for (const fx of [0.5, 0.3, 0.7]) {
      const x = r.left + r.width * fx, y = r.top + r.height * fy;
      const el = document.elementFromPoint(x, y);
      if (!el || !cell.contains(el)) continue;
      if (el.closest('button, a, [role=button], [data-automation-id="calendarevent"]')) continue;
      return {x: x - r.left, y: y - r.top};
    }
  return null;
}
"""


class FillError(Exception):
    pass


def dialog_dates(text: str, preferred: str = "") -> set[dt.date]:
    """วันที่ทุกแบบที่ข้อความบนหัว dialog อ่านได้"""
    out = set()
    for fmt in [preferred, *DIALOG_DATE_FORMATS] if preferred else DIALOG_DATE_FORMATS:
        try:
            out.add(dt.datetime.strptime(" ".join(text.split()), fmt).date())
        except ValueError:
            pass
    return out


def event_date(label: str) -> dt.date | None:
    """วันที่ท้าย aria-label ของ calendarevent (None ถ้าไม่มี)"""
    text = " ".join(label.split())
    for pattern in EVENT_DATE_RES:
        m = pattern.search(text)
        if m and m["m"].capitalize() in MONTHS:
            try:
                return dt.date(int(m["y"]), MONTHS.index(m["m"].capitalize()) + 1, int(m["d"]))
            except ValueError:
                return None
    return None


@dataclass
class Row:
    date: dt.date
    day_type: str
    contract: str
    role: str
    start: str
    end: str
    hours: float
    category: str
    job: str

    @property
    def key(self) -> str:
        return f"{self.date}|{self.contract}|{self.role}|{self.start}-{self.end}|{self.category}"


# ---------------------------------------------------------------- preview

def read_preview(path: Path) -> tuple[list[Row], list[tuple[dt.date, str, str]]]:
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Preview"]
    header = [str(c.value).strip() if c.value else "" for c in ws[1]]
    rows = []
    for values in ws.iter_rows(min_row=2, values_only=True):
        r = dict(zip(header, values))
        if not r.get("Date"):
            continue
        date = r["Date"].date() if isinstance(r["Date"], dt.datetime) else r["Date"]
        rows.append(Row(date, str(r["DayType"]), str(r["Contract"]).strip(), str(r.get("Role") or "").strip(),
                        str(r["Start"]), str(r["End"]), float(r["Hours"]), str(r["Category"]),
                        str(r.get("Job") or "").strip()))
    warns = []
    if "Warnings" in wb.sheetnames:
        for date, code, msg in wb["Warnings"].iter_rows(min_row=2, values_only=True):
            if date:
                warns.append((date.date() if isinstance(date, dt.datetime) else date, str(code), str(msg)))
    return rows, warns


def time_type_target(row: Row, wd: dict) -> tuple[list[str], callable, str]:
    """คืน (คำค้นที่จะลอง, ฟังก์ชันตรวจตัวเลือก, คำอธิบาย) ของ Time Type"""
    if row.contract == INHOUSE:
        code = wd["inhouse_code_by_role"].get(row.role)
        if not code:
            raise FillError(f"ไม่มี Time Entry Code สำหรับ Inhouse role {row.role!r} "
                            "(เพิ่ม workday.inhouse_code_by_role ใน team config)")
        return [code], (lambda text: text == code), code
    task = wd["project_task_by_role"].get(row.role) or wd["default_project_task"]
    if not task:
        raise FillError(f"ไม่มี task สำหรับ role {row.role!r} (เพิ่ม workday.project_task_by_role ใน team config)")
    return ([f"{row.contract} {task}", row.contract],
            (lambda text: text.startswith(row.contract + " ") and text.endswith("> " + task)),
            f"{row.contract} > {task}")


def _minutes(hhmm: str) -> int:
    h, m = hhmm.split(":")
    return int(h) * 60 + int(m)


def _strip_times(line: str) -> str:
    """ตัดช่วงเวลาออกจากบรรทัด เช่น 'Pilot change (15:00-20:00)' -> 'Pilot change'"""
    line = TIME_RE.sub("", line)
    line = re.sub(r"\(\s*\)|\[\s*\]", "", line)
    return re.sub(r"\s{2,}", " ", line).rstrip(" ,;:")


def job_for(row: Row) -> str:
    """บรรทัดของ Job ที่ตรงกับช่วงเวลาของ entry นี้ (ตัดเวลาออก)

    บรรทัดที่ระบุเวลา ใช้เมื่อเวลานั้นทับกับ entry; บรรทัดที่ไม่ระบุเวลา ใช้กับ Regular เท่านั้น
    ถ้าไม่มีบรรทัดไหนเข้าเงื่อนไข (เช่น OT ที่ Job ไม่ได้ระบุเวลา) ใช้ทุกบรรทัด
    """
    start, end = _minutes(row.start), _minutes(row.end)
    lines = [line for line in row.job.splitlines() if line.strip()]

    def keep(line: str) -> bool:
        ranges = parse_time_ranges(line)
        if not ranges:
            return row.category != OT
        return any(s < end and start < e for rs, re_ in ranges for s, e in day_segments(rs, re_))

    return "\n".join(_strip_times(line) for line in ([x for x in lines if keep(x)] or lines))


def comment_for(row: Row) -> str:
    """OT ขึ้นต้นด้วยช่วงเวลา (Workday ไม่มีช่องเวลาเริ่ม-จบ) ส่วนชั่วโมงปกติใส่แค่รายละเอียดงาน"""
    job = job_for(row)
    if row.category != OT:
        return job
    prefix = f"OT {row.start}-{row.end}"
    return f"{prefix} | {job}" if job else prefix


def event_matches(label: str, selected: str, row: Row) -> bool:
    """aria-label ของ calendarevent เช่น 'Not Submitted | <project> | 8 Hours - Billable | <phase> > <task> | <วันที่>'"""
    label, selected = " ".join(label.split()), " ".join(selected.split())
    if f"| {row.hours:g} Hour" not in label:
        return False
    if row.contract == INHOUSE:
        return f"| {selected} |" in label
    project, _, task_path = selected[len(row.contract) + 1:].partition(" > ")
    return f"| {project} |" in label and task_path in label


def load_log(path: Path) -> dict[str, str]:
    """entry ที่เคยกรอกสำเร็จ: key -> Time Type ที่เลือก (log รุ่นเก่าเป็น list ของ key จึงไม่มี Time Type)"""
    if not path.exists():
        return {}
    data = json.loads(path.read_text(encoding="utf-8"))
    return {key: "" for key in data} if isinstance(data, list) else dict(data)


# ---------------------------------------------------------------- Workday UI

def workday_config(cfg: dict) -> dict:
    return _merge(WORKDAY_DEFAULTS, cfg.get("workday") or {})


def check_workday(wd: dict, rows: list[Row]) -> None:
    """ค่าที่ต้องมีใน team config ก่อนเปิด browser"""
    missing = [k for k in ("url", "enter_time_path") if not wd.get(k)]
    if any(r.category == OT for r in rows):
        missing += [k for k, v in (("fields.calculation", wd["fields"]["calculation"]),
                                   ("overtime_calculation", wd["overtime_calculation"])) if not v]
    if missing:
        raise SystemExit(f"team config ไม่มีค่า workday: {', '.join(missing)}")


def workday_host(wd: dict) -> str:
    return urlparse(wd["url"]).netloc


class Workday:
    def __init__(self, page, dry_run: bool, date_format: str = "%d/%m/%Y", fields: dict | None = None):
        self.page = page
        self.dry_run = dry_run
        self.date_format = date_format
        self.fields = fields or WORKDAY_DEFAULTS["fields"]
        page.set_default_timeout(15000)

    def dialog(self):
        return self.page.locator(DIALOG).filter(has=self.page.locator('h2:text-is("Enter Time")'))

    def goto_month(self, date: dt.date) -> None:
        title = self.page.locator('[data-automation-id="dateRangeTitle"]')
        target = f"{MONTHS[date.month - 1]} {date.year}"
        for _ in range(24):
            current = title.inner_text().strip()
            if current == target:
                return
            if "–" in current:  # มุมมอง Week/Day: สลับเป็น Month (Workday จำมุมมองล่าสุดไว้)
                self.set_month_view(current)
                continue
            name, _, year = current.partition(" ")
            if name not in MONTHS or not year.isdigit():
                raise FillError(f"อ่านชื่อเดือนบนปฏิทินไม่ได้ ({current}) ต้องตั้ง Workday เป็นภาษา English")
            shown = dt.date(int(year), MONTHS.index(name) + 1, 1)
            button = "prevMonthButton" if date < shown else "nextMonthButton"
            self.page.locator(f'[data-automation-id="{button}"]').click()
            self.page.wait_for_function(
                "([sel, old]) => document.querySelector(sel)?.innerText.trim() !== old",
                arg=['[data-automation-id="dateRangeTitle"]', current])
        raise FillError(f"ไปเดือน {target} ไม่ได้")

    def set_month_view(self, current: str) -> None:
        self.page.locator('[data-automation-id="calendarViewSelection"]').click()
        option = self.page.locator(
            '[data-automation-id="dropdown-option"]:text-is("Month"), [role="option"]:text-is("Month"),'
            ' [role="menuitem"]:text-is("Month"), [role="menuitemradio"]:text-is("Month")').first
        try:
            option.click(timeout=5000)
            self.page.wait_for_function(
                "([sel, old]) => { const t = document.querySelector(sel)?.innerText.trim();"
                " return t && t !== old && !t.includes('–'); }",
                arg=['[data-automation-id="dateRangeTitle"]', current], timeout=15000)
        except PlaywrightError as exc:
            raise FillError(f"สลับปฏิทินเป็นมุมมอง Month ไม่ได้ ({current}) ให้เปลี่ยนเองก่อน") from exc

    def day_events(self, date: dt.date) -> list[str]:
        month = f"{MONTHS[date.month - 1]} {date.year}"
        events = self.page.locator('[data-automation-id="calendarevent"]')
        # หลังเปลี่ยนเดือน entry ยังโหลดไม่ครบ: รอจนมี event ของเดือนนี้ (อย่างน้อย Time Period End)
        # และรายการนิ่ง 1.5 วินาที ไม่งั้นจะเห็นเป็นว่างแล้วกรอกซ้ำ
        labels, stable_since, deadline = None, time.time(), time.time() + 15
        while True:
            current = events.evaluate_all("els => els.map(e => e.getAttribute('aria-label') || '')")
            if current != labels:
                labels, stable_since = current, time.time()
            loaded = any((d.year, d.month) == (date.year, date.month) for d in map(event_date, labels) if d)
            if loaded and time.time() - stable_since >= 1.5:
                break
            if time.time() > deadline:
                raise FillError(f"รายการ entry ของ {month} ไม่โหลด (กันกรอกซ้ำจึงไม่ทำต่อ)")
            time.sleep(0.3)
        return [lb for lb in labels if "|" in lb and event_date(lb) == date]

    def open_dialog(self, date: dt.date) -> None:
        self.goto_month(date)
        cell = self.page.locator(f'[data-automation-id="calendarDateCell-{date.month - 1}-{date.day}"]')
        # หลังกด OK ฉากบังของ dialog ยังค้างอยู่ครู่หนึ่ง และปฏิทินอาจ render ใหม่ จึงลองซ้ำ
        deadline = time.time() + 15
        spot = None
        while not spot:
            if time.time() > deadline:
                raise FillError("ไม่มีพื้นที่ว่างในช่องวันที่ให้คลิก")
            try:
                cell.scroll_into_view_if_needed()
                spot = cell.evaluate(FREE_SPOT_JS)
            except PlaywrightError:
                pass
            if not spot:
                time.sleep(0.5)
        cell.click(position=spot)
        dialog = self.dialog()
        dialog.wait_for(state="visible")
        shown = dialog.locator('[data-automation-id="textView"]').first.inner_text().strip()
        if date not in dialog_dates(shown, self.date_format):
            self.cancel()
            raise FillError(f"dialog เปิดผิดวัน (หัว dialog แสดง {shown!r} แต่ต้องเป็น {date:%Y-%m-%d}) "
                            "ถ้าเป็นวันที่เดียวกันแต่คนละรูปแบบ ให้ใส่ workday.date_format ใน config")

    def select_prompt(self, label: str, queries: list[str], matches) -> str:
        dialog = self.dialog()
        box = dialog.get_by_role("textbox", name=label, exact=True)
        selected = dialog.get_by_role("listbox", name=f"items selected for {label}").get_by_role("option")
        for query in queries:
            box.click()
            box.fill(query)
            box.press("Enter")
            deadline = time.time() + 10
            while time.time() < deadline:
                if selected.count():
                    text = selected.first.inner_text().strip()
                    if matches(text):
                        return text
                    raise FillError(f"{label}: Workday เลือก {text!r} ไม่ตรงกับที่ต้องการ")
                options = self.page.locator('[data-automation-id="promptOption"]')
                hit = next((options.nth(i) for i in range(options.count())
                            if matches(options.nth(i).inner_text().strip())), None)
                if hit is not None:
                    hit.click()
                    time.sleep(1)
                    continue
                if self.page.get_by_text("No matches found").count():
                    break
                time.sleep(0.3)
        if label == self.fields["time_type"]:
            raise FillError(f"{label}: หา {queries[0]!r} ใน Workday ไม่เจอ "
                            "(Contract อาจยังไม่เปิดให้ลงเวลา หรือเลขสัญญา/Role ใน Excel ผิด)")
        raise FillError(f"{label}: หาไม่เจอด้วยคำค้น {queries}")

    def fill_entry(self, row: Row, wd: dict) -> str:
        f = self.fields
        queries, matches, _ = time_type_target(row, wd)
        selected = self.select_prompt(f["time_type"], queries, matches)
        dialog = self.dialog()
        # เลือก Time Type แล้วฟอร์มโหลดช่อง Details ใหม่ ต้องรอให้นิ่งก่อน
        dialog.get_by_role("textbox", name=f["calculation"] or f["comment"], exact=True).wait_for()
        time.sleep(1)
        if row.category == OT:
            calc = wd["overtime_calculation"]
            self.select_prompt(f["calculation"], [calc], lambda text: text == calc)
            time.sleep(1)
        comment = dialog.get_by_role("textbox", name=f["comment"], exact=True)
        comment.fill(comment_for(row))
        # Hours กรอกท้ายสุด เพราะการเลือก prompt ทำให้ฟอร์มรีเซ็ตค่า Hours
        hours = dialog.get_by_role("textbox", name=f["hours"], exact=True)
        for _ in range(3):
            hours.click()
            hours.press("Control+A")
            hours.press_sequentially(f"{row.hours:g}")
            hours.press("Tab")
            time.sleep(1)
            try:
                if abs(float(hours.input_value().replace(",", "")) - row.hours) < 0.001:
                    break
            except ValueError:
                pass
        else:
            raise FillError(f"กรอก Hours ไม่ติด (ค่าในช่อง {hours.input_value()!r})")
        if comment.input_value() != comment_for(row):
            raise FillError("Comment ถูกล้างหลังกรอก Hours")
        return selected

    def save(self) -> None:
        dialog = self.dialog()
        ok = dialog.get_by_role("button", name="OK", exact=True)
        if ok.inner_text().strip() != "OK":
            raise FillError("ไม่พบปุ่ม OK ของ dialog Enter Time")
        ok.click()
        deadline = time.time() + 20
        while time.time() < deadline:
            if not dialog.count():
                return
            errors = dialog.locator('[role=listitem]:has-text("Error")')
            if errors.count():
                raise FillError("Workday แจ้ง: " + " / ".join(errors.all_inner_texts()))
            time.sleep(0.5)
        raise FillError("dialog ไม่ปิดหลังกด OK")

    def cancel(self) -> None:
        """กด Cancel (และ Discard ถ้าถาม) จนกว่า dialog จะปิด"""
        dialog = self.dialog()
        discard = self.page.locator('[data-automation-id="wd-CommandButton_uic_genericYesButton"]')
        for _ in range(4):
            if not dialog.count():
                return
            if discard.count() and discard.first.is_visible():
                discard.first.click()
            else:
                dialog.get_by_role("button", name="Cancel", exact=True).click()
            deadline = time.time() + 5
            while time.time() < deadline and dialog.count() and not discard.count():
                time.sleep(0.3)
            time.sleep(0.5)
        if dialog.count():
            raise FillError("ปิด dialog ไม่ได้")


# ---------------------------------------------------------------- main

EDGE_PATHS = [r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
              r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"]


def profile_path(cfg: dict) -> Path:
    """session ที่ login Workday ไว้ ห้ามแชร์ให้คนอื่น"""
    return cfg["base_dir"] / workday_config(cfg)["profile_dir"] / person_of(cfg)


def find_edge(wd: dict) -> str:
    local = Path(os.environ.get("LOCALAPPDATA", "")) / "Microsoft/Edge/Application/msedge.exe"
    candidates = [wd.get("browser_path") or "", *EDGE_PATHS, str(local)]
    exe = next((p for p in candidates if p and Path(p).exists()), None) or shutil.which("msedge")
    if not exe:
        raise SystemExit("ไม่พบ Microsoft Edge: ติดตั้ง Edge หรือใส่ workday.browser_path ใน config ของผู้ใช้")
    return exe


def launch_edge(cfg: dict, port: int) -> None:
    """เปิด Edge เป็น process แยก ให้ browser เปิดค้างไว้ตรวจ/Submit ต่อได้หลังสคริปต์จบ"""
    wd = workday_config(cfg)
    profile = profile_path(cfg)
    profile.mkdir(parents=True, exist_ok=True)
    flags = subprocess.DETACHED_PROCESS | subprocess.CREATE_NEW_PROCESS_GROUP if os.name == "nt" else 0
    subprocess.Popen([find_edge(wd), f"--remote-debugging-port={port}", f"--user-data-dir={profile}",
                      "--no-first-run", "--start-maximized", wd["url"] + wd["enter_time_path"]],
                     creationflags=flags, close_fds=True,
                     stdin=subprocess.DEVNULL, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)


def connect(pw, cfg: dict, port: int):
    """ต่อ Edge ผ่าน CDP ถ้ายังไม่มีก็เปิดให้ (สคริปต์จบแล้ว browser ไม่ปิดตาม)"""
    url = f"http://127.0.0.1:{port}"
    try:
        browser = pw.chromium.connect_over_cdp(url, timeout=3000)
    except PlaywrightError:
        launch_edge(cfg, port)
        deadline = time.time() + 30
        while True:
            try:
                browser = pw.chromium.connect_over_cdp(url, timeout=3000)
                break
            except PlaywrightError:
                if time.time() > deadline:
                    raise SystemExit(f"ต่อ Edge ที่พอร์ต {port} ไม่ได้: ถ้ามีหน้าต่าง Edge ที่ใช้ profile "
                                     f"{profile_path(cfg)} เปิดอยู่ ให้ปิดก่อนแล้วรันใหม่ หรือใช้ --port อื่น")
                time.sleep(1)
    pages = [p for ctx in browser.contexts for p in ctx.pages]
    workday = [p for p in pages if workday_host(workday_config(cfg)) in p.url]
    return (workday or pages)[-1] if pages else browser.contexts[0].new_page()


def wait_enter_time(page, cfg: dict, minutes: float = 15) -> None:
    wd = workday_config(cfg)
    url = wd["url"] + wd["enter_time_path"]
    header = page.locator('[data-automation-id="pageHeaderTitleText"]:text-is("Enter Time")')
    if not header.count():
        try:
            if workday_host(wd) not in page.url:  # Edge ที่สคริปต์เปิดโหลด URL นี้อยู่แล้ว
                page.goto(url)
        except PlaywrightError:
            pass  # ถูก redirect ไปหน้า login (SSO) ระหว่างโหลด เป็นเรื่องปกติ
        print("รอ login และหน้า Enter Time ... (ถ้าไม่ขึ้นเอง ให้เปิดเมนู Time > Enter Time ใน Workday)",
              flush=True)
        deadline = time.time() + minutes * 60
        redirected = False
        while not header.count():
            if time.time() > deadline:
                raise SystemExit("หมดเวลารอ login / หน้า Enter Time (ถ้า Workday เป็นภาษาไทย ให้เปลี่ยนเป็น English)")
            try:
                # หลัง login ครั้งแรกพาไปหน้า Enter Time ครั้งเดียว ที่เหลือให้ผู้ใช้เปิดเองได้
                if not redirected and page.locator('[data-automation-id="globalSearchInput"]').count():
                    redirected = True
                    if wd["enter_time_path"] not in page.url:
                        page.goto(url)
            except PlaywrightError:
                pass
            time.sleep(2)
    title = page.locator('[data-automation-id="dateRangeTitle"]')
    title.wait_for()
    if not any(m[:3] in title.inner_text() for m in MONTHS):  # Month = "October 2026", Week = "26 Oct – 1 Nov 2026"
        raise SystemExit(f"ปฏิทินแสดง {title.inner_text()!r}: ต้องตั้ง Workday เป็นภาษา English")


def confirm(question: str, assume_yes: bool) -> bool:
    if assume_yes:
        return True
    if not sys.stdin.isatty():
        return False
    return input(f"{question} [y/N]: ").strip().lower() == "y"


def main(argv=None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")  # ข้อความ SystemExit
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    g = p.add_mutually_exclusive_group(required=True)
    g.add_argument("--week", help="ใช้ output/preview_<week>.xlsx")
    g.add_argument("--month", help="ใช้ output/preview_<month>.xlsx")
    g.add_argument("--preview", type=Path)
    p.add_argument("--dry-run", action="store_true", help="เปิดฟอร์มและเลือกค่า แต่กด Cancel ไม่บันทึก")
    p.add_argument("--accept-under-8h", action="store_true", help="ยืนยันกรอกต่อแม้มีวัน UNDER_8H")
    p.add_argument("--only-date", type=dt.date.fromisoformat, action="append", help="กรอกเฉพาะวันที่นี้ (ใส่ซ้ำได้)")
    p.add_argument("--skip-date", type=dt.date.fromisoformat, action="append", default=[],
                   help="ไม่กรอกวันที่นี้ (ใส่ซ้ำได้)")
    p.add_argument("--person", help="ชื่อหัวคอลัมน์ของคน (ทับค่า person_column ใน config)")
    p.add_argument("--port", type=int, default=9222)
    p.add_argument("--config", type=Path, help="config ของผู้ใช้ (ค่าตั้งต้นอยู่ใน data dir)")
    args = p.parse_args(argv)

    cfg = load_config(args.config)
    require_team(cfg)
    if args.person:
        cfg["person_column"] = args.person
    wd = workday_config(cfg)
    out_dir = output_dir(cfg)
    label = args.week or args.month or args.preview.stem.removeprefix("preview_")
    preview = args.preview or out_dir / f"preview_{label}.xlsx"
    if not preview.exists():
        raise SystemExit(f"ไม่พบ {preview}: รัน preview ของเดือนนี้ก่อน")
    rows, warns = read_preview(preview)
    if args.only_date:
        rows = [r for r in rows if r.date in args.only_date]
    rows = [r for r in rows if r.date not in args.skip_date]
    print(f"{person_of(cfg)}: อ่าน {preview}: {len(rows)} entry{' (dry-run)' if args.dry_run else ''}")
    check_workday(wd, rows)

    dates = {r.date for r in rows}
    under = [w for w in warns if w[1] == "UNDER_8H" and w[0] in dates]
    if under:
        print("\nวันทำงานที่ชั่วโมงปกติไม่ครบ:")
        for date, _, msg in under:
            print(f"  {date} {msg}")
        if not confirm("กรอกต่อหรือไม่", args.accept_under_8h or args.dry_run):
            print("หยุด: รันใหม่พร้อม --accept-under-8h ถ้าต้องการกรอกต่อ")
            return 2

    log_path = out_dir / f"filled_{label}.json"
    # Workday คือข้อมูลจริง (ผู้ใช้อาจลบ/กรอกเอง) log ใช้แค่ข้ามการเปิด dialog ของ entry ที่ยังอยู่ใน Workday
    filled = load_log(log_path)
    results: list[tuple[Row, str, str]] = []

    with sync_playwright() as pw:
        page = connect(pw, cfg, args.port)
        page.bring_to_front()  # prompt ของ Workday ปิดเองถ้าหน้าต่างไม่มี focus
        wait_enter_time(page, cfg)
        ui = Workday(page, args.dry_run, wd["date_format"], wd["fields"])
        print("ระหว่างกรอก อย่าคลิกในหน้าต่าง browser\n", flush=True)

        events_by_date: dict[dt.date, list[str]] = {}
        for row in rows:
            tag = f"{row.date} {row.contract:<9} {row.start}-{row.end} {row.category:<7}"
            if row.day_type == "Leave" and wd.get("skip_leave_days", True):
                results.append((row, "skipped", "วันลา: ไม่กรอกอัตโนมัติ (เบิก OT วันลาไม่ได้) ตรวจและตัดสินใจเอง"))
                print(f"SKIP  {tag} วันลา")
                continue
            try:
                ui.goto_month(row.date)
                if row.date not in events_by_date:
                    events_by_date[row.date] = ui.day_events(row.date)
                existing = events_by_date[row.date]
                known = filled.get(row.key)
                dup = next((lb for lb in existing if event_matches(lb, known, row)), None) if known else None
                if dup:  # เคยกรอกและยังอยู่ใน Workday ไม่ต้องเปิด dialog
                    existing.remove(dup)  # ใช้จับคู่ได้ครั้งเดียว
                    results.append((row, "skipped", f"มีใน Workday แล้ว: {dup}"))
                    print(f"SKIP  {tag} มีใน Workday แล้ว")
                    continue
                ui.open_dialog(row.date)
                selected = ui.fill_entry(row, wd)
                dup = next((lb for lb in existing if event_matches(lb, selected, row)), None)
                if dup:
                    existing.remove(dup)
                    ui.cancel()
                    if row.key in filled and not filled[row.key] and not args.dry_run:
                        filled[row.key] = selected  # log รุ่นเก่า: เก็บ Time Type ไว้ รอบหน้าไม่ต้องเปิด dialog
                        log_path.write_text(json.dumps(filled, ensure_ascii=False, indent=1), encoding="utf-8")
                    results.append((row, "skipped", f"มีใน Workday แล้ว: {dup}"))
                    print(f"SKIP  {tag} มีใน Workday แล้ว")
                    continue
                again = " (เคยกรอกแล้วแต่ไม่พบใน Workday จึงกรอกใหม่)" if row.key in filled else ""
                if args.dry_run:
                    page.screenshot(path=str(out_dir / f"fill_dryrun_{row.date}_{row.start.replace(':', '')}.png"))
                    ui.cancel()
                    results.append((row, "dry-run", selected + again))
                    print(f"DRY   {tag} {row.hours:g} ชม. -> {selected}{again}")
                    continue
                ui.save()
                filled[row.key] = selected
                log_path.write_text(json.dumps(filled, ensure_ascii=False, indent=1), encoding="utf-8")
                results.append((row, "filled", selected + again))
                print(f"OK    {tag} {row.hours:g} ชม. -> {selected}{again}")
            except (FillError, PlaywrightError) as exc:
                msg = str(exc).splitlines()[0]
                try:
                    page.screenshot(path=str(out_dir / f"fill_error_{row.date}_{row.start.replace(':', '')}.png"))
                    ui.cancel()
                except PlaywrightError:
                    pass
                results.append((row, "error", msg))
                print(f"ERROR {tag} {msg}")

        report = out_dir / f"fill_{label}.csv"
        with open(report, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f)
            writer.writerow(["Date", "Contract", "Role", "Start", "End", "Hours", "Category", "Status", "Detail"])
            for row, status, detail in results:
                writer.writerow([row.date, row.contract, row.role, row.start, row.end, row.hours, row.category,
                                 status, detail])
        count = lambda s: sum(1 for _, st, _ in results if st == s)  # noqa: E731
        print(f"\nสรุป: กรอกสำเร็จ {count('filled')}, dry-run {count('dry-run')}, "
              f"ข้าม {count('skipped')}, error {count('error')}  ->  {report}")
        leave = [row for row, status, detail in results if status == "skipped" and detail.startswith("วันลา")]
        if leave:
            print("\n!! วันลาที่มีงาน ไม่ได้กรอกให้ (เบิก OT วันลาไม่ได้) ตัดสินใจและกรอกเอง:")
            for row in leave:
                print(f"   {row.date} {row.contract} {row.start}-{row.end} {row.category} {row.hours:g} ชม.")
        errors = [(row, detail) for row, status, detail in results if status == "error"]
        if errors:
            print(f"\n!! ไม่ได้กรอก {len(errors)} entry (กรอกเองใน Workday หรือแก้ Excel แล้วรันใหม่):")
            for row, detail in errors:
                print(f"   {row.date} {row.contract} {row.role or '-'} {row.start}-{row.end} {row.category} "
                      f"{row.hours:g} ชม.: {detail}")
        extra = [(date, label) for date, labels in sorted(events_by_date.items()) for label in labels]
        if extra:
            print("\n!! ใน Workday มี entry ที่ไม่ตรงกับ preview (อาจกรอกเองหรือชั่วโมงไม่ตรง) ตรวจว่าซ้ำหรือไม่:")
            for date, label in extra:
                print(f"   {date} {' '.join(label.split())}")
        print("** ยังไม่ได้กด Submit: ตรวจใน Workday แล้วกด Submit เอง (หน้าต่าง Edge ยังเปิดอยู่) **")
    return 1 if count("error") else 0


if __name__ == "__main__":
    sys.exit(main())
