"""อ่านตารางงานของทีมจาก Excel แล้วสร้าง preview entry แยก Regular / OT พร้อมคำเตือน

ไม่แตะ Workday และไม่แก้ไฟล์ Excel ต้นฉบับ (copy ไปอ่านใน temp)
เวลาภายในเก็บเป็นนาทีตั้งแต่เที่ยงคืน 0-1440

    python parse_schedule.py --month 2026-10
    python parse_schedule.py --week 2026-10-26
    python parse_schedule.py --month 2026-10 --person <ชื่อ>      ทับชื่อใน config ของผู้ใช้

ค่าตั้งต้น <- team config (ค่าขององค์กร อยู่นอก repo) <- config ของผู้ใช้ (ชื่อคอลัมน์ ฯลฯ)
config ของผู้ใช้อยู่ใน data dir (ดู home_dir) สร้างด้วย setup_user.py
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import os
import re
import shutil
import sys
import tempfile
from dataclasses import dataclass
from pathlib import Path

import openpyxl
import yaml
from openpyxl.styles import Font

HOLIDAY, LEAVE, WORKDAY = "Holiday", "Leave", "Workday"
REGULAR, OT = "Regular", "OT"
INHOUSE = "Inhouse"
DAY = 24 * 60

# ค่าตั้งต้นที่ไม่ผูกกับองค์กรใด ค่าขององค์กร (ชื่อไฟล์, ชื่อชีต, Workday) อยู่ใน team config
DEFAULT_CONFIG = {
    "excel_path": "",                         # ว่าง = หาไฟล์ excel_name เอง (โฟลเดอร์ของ team config แล้ว OneDrive)
    "excel_name": "",
    "sheet": "{year}",                        # {year} = ปีของช่วงที่อ่าน
    "contract_sheet": "Contract",
    "output_dir": "output",
    "person_column": "",
    "regular_window": "08:00-18:00",
    "min_regular_hours": 8,
    "pad_regular_from_hours": 6,
    "lunch_deduction": False,
    "lunch_window": "12:00-13:00",
    "colors": {
        "holiday": {"theme": [8], "tint_min": 0.5, "tint_max": 0.7, "rgb": []},
        "leave": {"rgb": ["FFFFC000"]},
    },
    "ignore_free_text": ["Leave", "A leave"],
}

TIME_RE = re.compile(r"(\d{1,2})\s*[:.]\s*(\d{2})\s*-\s*(\d{1,2})\s*[:.]\s*(\d{2})")
BLOCK_SPLIT_RE = re.compile(r"(?im)^(?=[ \t]*con(?:tact|tract)\s*:)")
FIELD_RE = re.compile(r"^\s*(contact|contract|site|role|time|job)\s*:\s*(.*)$", re.I)
FIELD_ATTRS = {"contact": "contract_raw", "contract": "contract_raw", "site": "site",
               "role": "role", "time": "time_text", "job": "job"}
CONTRACT_ID_RE = re.compile(r"A\d{2}-\d{4}|A\d{7}", re.I)


# ---------------------------------------------------------------- config

def _merge(base: dict, override: dict) -> dict:
    out = dict(base)
    for key, value in (override or {}).items():
        out[key] = _merge(out[key], value) if isinstance(value, dict) and isinstance(out.get(key), dict) else value
    return out


USER_CONFIG = "config.yaml"                 # config ของผู้ใช้ใน home_dir()
TEAM_CONFIG_NAME = "timesheet.team.yaml"    # ชื่อไฟล์ team config ที่ setup หาใน OneDrive


def home_dir() -> Path:
    """data dir ของผู้ใช้: config, output, session ที่ login Workday (อยู่นอกโฟลเดอร์ plugin ไม่หายตอนอัปเดต)"""
    if os.environ.get("TIMESHEET_HOME"):
        return Path(os.environ["TIMESHEET_HOME"])
    if os.environ.get("LOCALAPPDATA"):
        return Path(os.environ["LOCALAPPDATA"]) / "workday-timesheet"
    return Path.home() / ".workday-timesheet"


def read_yaml(path: Path) -> dict:
    with open(path, encoding="utf-8") as f:
        return yaml.safe_load(f) or {}


def load_config(path: Path | None = None) -> dict:
    """ค่าตั้งต้น <- team config (path อยู่ในคีย์ team_config) <- config ของผู้ใช้"""
    path = Path(path) if path else home_dir() / USER_CONFIG
    user = read_yaml(path) if path.exists() else {}
    cfg = dict(DEFAULT_CONFIG)
    team = user.get("team_config")
    if team:
        team = Path(os.path.expandvars(team))
        if not team.exists():
            raise SystemExit(f"ไม่พบ team config: {team} (รัน setup ใหม่ หรือแก้ team_config ใน {path})")
        cfg = _merge(cfg, read_yaml(team))
    cfg = _merge(cfg, user)
    cfg["base_dir"] = path.resolve().parent
    return cfg


def require_team(cfg: dict) -> None:
    if not cfg.get("team_config"):
        raise SystemExit("ยังไม่ได้ตั้งค่า: รัน setup ก่อน (ต้องมี team config ของทีม)")


def onedrive_roots() -> list[Path]:
    """โฟลเดอร์ OneDrive ของเครื่องนี้ (ของบริษัทก่อน)"""
    roots = [os.environ.get(k) for k in ("OneDriveCommercial", "OneDrive")]
    home = Path.home()
    if home.is_dir():
        roots += [str(p) for p in sorted(home.iterdir()) if p.is_dir() and p.name.startswith("OneDrive")]
    seen, out = set(), []
    for r in roots:
        if r and Path(r).is_dir() and Path(r).resolve() not in seen:
            seen.add(Path(r).resolve())
            out.append(Path(r))
    return out


def search_excel(name: str, roots: list[Path], max_depth: int = 5) -> list[Path]:
    """หาไฟล์ชื่อ name ใต้ roots (ไฟล์ .url ของ shortcut ไฟล์เดี่ยวชื่อไม่ตรงจึงไม่นับ) เรียงใหม่สุดก่อน"""
    found = set()
    for root in roots:
        base = len(root.parts)
        for dirpath, dirnames, filenames in os.walk(root):
            if len(Path(dirpath).parts) - base >= max_depth:
                dirnames.clear()
            dirnames[:] = [d for d in dirnames if not d.startswith(".")]
            if name in filenames:
                found.add(Path(dirpath) / name)
    return sorted(found, key=lambda p: p.stat().st_mtime, reverse=True)


def excel_candidates(cfg: dict) -> list[Path]:
    """ไฟล์ชื่อ excel_name ในโฟลเดอร์เดียวกับ team config ก่อน แล้วค่อยหาใน OneDrive"""
    name = cfg.get("excel_name")
    if not name:
        raise SystemExit("team config ไม่มี excel_name (ชื่อไฟล์ตารางงาน)")
    out = []
    if cfg.get("team_config"):
        beside = Path(os.path.expandvars(cfg["team_config"])).parent / name
        if beside.exists():
            out.append(beside)
    return out + [p for p in search_excel(name, onedrive_roots()) if p not in out]


def find_excel(cfg: dict) -> Path:
    if cfg.get("excel_path"):
        path = cfg["base_dir"] / Path(os.path.expandvars(cfg["excel_path"]))
        if not path.exists():
            raise SystemExit(f"ไม่พบไฟล์ Excel: {path} (แก้ excel_path ใน config ของผู้ใช้ หรือรัน setup ใหม่)")
        return path
    found = excel_candidates(cfg)
    if not found:
        raise SystemExit(f"หาไฟล์ {cfg['excel_name']!r} ไม่เจอ: ใน Teams/SharePoint ให้กด "
                         "'Add shortcut to My files' ที่โฟลเดอร์ที่มีไฟล์ (ไม่ใช่ที่ตัวไฟล์) แล้วรอ OneDrive sync "
                         "หรือรัน setup พร้อม --excel <path>")
    return found[0]


def person_of(cfg: dict) -> str:
    person = str(cfg.get("person_column") or "").strip()
    if not person:
        raise SystemExit("ยังไม่ได้ตั้งชื่อ: รัน setup --person <ชื่อ> หรือใช้ --person")
    return person


def output_dir(cfg: dict) -> Path:
    """แยกโฟลเดอร์ output ต่อคน (เครื่องเดียวใช้หลายคนได้ไม่ทับกัน)"""
    return cfg["base_dir"] / cfg["output_dir"] / person_of(cfg)


# ---------------------------------------------------------------- time helpers

def fmt(minutes: int) -> str:
    return f"{minutes // 60:02d}:{minutes % 60:02d}"


def parse_time_ranges(text: str | None) -> list[tuple[int, int]]:
    """หาทุกช่วง HH:MM-HH:MM ในข้อความ; end 00:00/24:00 = 1440, end < start = ข้ามเที่ยงคืน"""
    ranges = []
    for m in TIME_RE.finditer(text or ""):
        sh, sm, eh, em = (int(g) for g in m.groups())
        if sh > 24 or eh > 24 or sm > 59 or em > 59:
            continue
        start = (sh * 60 + sm) % DAY
        end = eh * 60 + em
        if end == 0 or end >= DAY:
            end = DAY
        if start != end:
            ranges.append((start, end))
    return ranges


def day_segments(start: int, end: int) -> list[tuple[int, int]]:
    """ช่วงข้ามเที่ยงคืนแตกเป็น 2 ช่วงในกรอบวันเดียว (ใช้เทียบการครอบคลุม)"""
    return [(start, end)] if end > start else [(start, DAY), (0, end)]


def subtract(seg: tuple[int, int], others: list[tuple[int, int]]) -> list[tuple[int, int]]:
    parts = [seg]
    for os_, oe in others:
        parts = [p for s, e in parts
                 for p in ((s, min(e, os_)), (max(s, oe), e)) if p[0] < p[1]]
    return parts


def union_minutes(segs: list[tuple[int, int]]) -> int:
    total, cur_s, cur_e = 0, None, None
    for s, e in sorted(segs):
        if cur_e is None or s > cur_e:
            if cur_e is not None:
                total += cur_e - cur_s
            cur_s, cur_e = s, e
        else:
            cur_e = max(cur_e, e)
    if cur_e is not None:
        total += cur_e - cur_s
    return total


def parse_window(text: str) -> tuple[int, int]:
    ranges = parse_time_ranges(text)
    if not ranges or ranges[0][1] <= ranges[0][0]:
        raise ValueError(f"ช่วงเวลาไม่ถูกต้อง: {text!r}")
    return ranges[0]


def split_regular_ot(start: int, end: int, day_type: str, cfg: dict) -> list[tuple[int, int, str]]:
    """แบ่งช่วง (start < end ภายในวันเดียว) เป็น Regular / OT ตามประเภทวัน"""
    if day_type == HOLIDAY:
        return [(start, end, OT)]
    win_s, win_e = parse_window(cfg["regular_window"])
    lunch = [parse_window(cfg["lunch_window"])] if cfg.get("lunch_deduction") else []
    out = []
    for s, e, cat in ((start, min(end, win_s), OT),
                      (max(start, win_s), min(end, win_e), REGULAR),
                      (max(start, win_e), end, OT)):
        if s >= e:
            continue
        if cat == REGULAR:
            out.extend((ps, pe, cat) for ps, pe in subtract((s, e), lunch))
        else:
            out.append((s, e, cat))
    return out


# ---------------------------------------------------------------- cell parsing

@dataclass
class Block:
    contract_raw: str = ""
    site: str = ""
    role: str = ""
    time_text: str = ""
    job: str = ""
    has_time_line: bool = False

    @property
    def contract_id(self) -> str:
        raw = self.contract_raw.strip()
        if raw == "-" or re.search(INHOUSE, raw, re.I):
            return INHOUSE
        m = CONTRACT_ID_RE.search(raw)
        return m.group(0).upper() if m else raw

    @property
    def ranges(self) -> list[tuple[int, int]]:
        return parse_time_ranges(self.time_text)

    @property
    def job_ranges(self) -> list[tuple[int, int]]:
        return parse_time_ranges(self.job)


def parse_block(text: str) -> tuple[Block, str]:
    """อ่านบล็อก Contact/Site/Role/Time/Job; คืน (block, ข้อความที่อยู่หลังบล็อก)

    Job อ่านต่อหลายบรรทัดจนเจอบรรทัดว่าง ข้อความหลังบรรทัดว่างนั้นถือว่าอยู่นอกบล็อก
    """
    block = Block()
    state, last, gap = "head", None, False
    tail = []
    for line in text.splitlines():
        if state == "tail":
            tail.append(line)
            continue
        m = FIELD_RE.match(line)
        if m and state == "head":
            last, gap = FIELD_ATTRS[m.group(1).lower()], False
            setattr(block, last, m.group(2).strip())
            if last == "time_text":
                block.has_time_line = True
            elif last == "job":
                state = "job"
        elif not line.strip():
            if state == "job" and block.job:
                state = "tail"
            elif state == "head" and last:
                gap = True
        elif state == "head" and (gap or last is None):
            state = "tail"
            tail.append(line)
        else:
            # บรรทัดต่อของฟิลด์ล่าสุด เช่น Job หลายบรรทัด หรือเลขสัญญาหลายตัว
            setattr(block, last, (getattr(block, last) + "\n" + line.strip()).strip())
    return block, "\n".join(tail).strip()


def split_cell(text) -> tuple[list[Block], list[str]]:
    """แยกช่องเป็นบล็อกงาน และข้อความอิสระที่อยู่นอกบล็อก"""
    if text is None or not str(text).strip():
        return [], []
    parts = BLOCK_SPLIT_RE.split(str(text))
    free = [parts[0].strip()] if parts[0].strip() else []
    blocks = []
    for part in parts[1:]:
        if not part.strip():
            continue
        block, tail = parse_block(part)
        blocks.append(block)
        if tail:
            free.append(tail)
    return blocks, free


# ---------------------------------------------------------------- colors

def fill_kind(cell, cfg: dict) -> str | None:
    """คืน HOLIDAY / LEAVE / None จากสี fill ของช่อง"""
    fill = getattr(cell, "fill", None)
    if fill is None or fill.fill_type != "solid":
        return None
    color = fill.fgColor
    holiday, leave = cfg["colors"]["holiday"], cfg["colors"]["leave"]
    if color.type == "theme":
        tint = color.tint or 0.0
        if color.theme in holiday["theme"] and holiday["tint_min"] <= tint <= holiday["tint_max"]:
            return HOLIDAY
    elif color.type == "rgb" and isinstance(color.rgb, str):
        rgb = color.rgb.upper()
        if rgb in [c.upper() for c in holiday.get("rgb", [])]:
            return HOLIDAY
        if rgb in [c.upper() for c in leave.get("rgb", [])]:
            return LEAVE
    return None


def classify_day(date_cell, person_cell, cfg: dict) -> str:
    """ฟ้าในช่องวันที่หรือช่องของคน = วันหยุด, ส้มในช่องของคน = วันลา"""
    if HOLIDAY in (fill_kind(date_cell, cfg), fill_kind(person_cell, cfg)):
        return HOLIDAY
    if fill_kind(person_cell, cfg) == LEAVE:
        return LEAVE
    return WORKDAY


# ---------------------------------------------------------------- entries

@dataclass
class Day:
    date: dt.date
    day_type: str
    text: str | None


@dataclass
class ContractInfo:
    contract_no: str = ""
    project: str = ""
    customer: str = ""


@dataclass
class Entry:
    date: dt.date
    day_type: str
    contract: str
    site: str
    role: str
    start: int
    end: int
    category: str
    job: str
    note: str = ""

    @property
    def hours(self) -> float:
        return (self.end - self.start) / 60


@dataclass
class Warn:
    date: dt.date
    code: str
    message: str


def _oneline(text: str, limit: int = 80) -> str:
    text = " / ".join(line.strip() for line in text.splitlines() if line.strip())
    return text if len(text) <= limit else text[: limit - 3] + "..."


def _merge_pieces(pieces: list[tuple[int, int, Block]]) -> list[tuple[int, int, Block]]:
    """รวมช่วงที่ทับ/ต่อกันของ contract เดียวกัน (ใช้ข้อมูลบล็อกแรก)"""
    merged: list[list] = []
    for s, e, b in sorted(pieces, key=lambda p: (p[2].contract_id, p[0])):
        if merged and merged[-1][2].contract_id == b.contract_id and s <= merged[-1][1]:
            merged[-1][1] = max(merged[-1][1], e)
        else:
            merged.append([s, e, b])
    return sorted((tuple(p) for p in merged), key=lambda p: p[0])


def pad_regular(entries: list[Entry], need: int, cfg: dict) -> tuple[list[Entry], list[tuple[int, int]]]:
    """ขยาย entry Regular ให้ได้เพิ่มอีก need นาที ภายใน regular_window

    ลำดับ: เติมช่องว่างระหว่าง entry (เช่นพักเที่ยง) -> ต่อท้าย entry สุดท้าย -> ต่อหน้า entry แรก
    """
    win_s, win_e = parse_window(cfg["regular_window"])
    lunch = parse_window(cfg["lunch_window"]) if cfg.get("lunch_deduction") else None
    regs = sorted((e for e in entries if e.category == REGULAR), key=lambda e: e.start)
    added: list[tuple[int, int]] = []
    if not regs:
        return entries, added

    def extend(entry: Entry, start: int, end: int) -> None:
        nonlocal need
        n = min(end - start, need)
        if n <= 0:
            return
        need -= n
        if start == entry.end:
            seg = (entry.end, entry.end + n)
            entry.end += n
        else:
            seg = (entry.start - n, entry.start)
            entry.start -= n
        added.append(seg)
        entry.note = "; ".join(filter(None, [entry.note, f"เติม {fmt(seg[0])}-{fmt(seg[1])}"]))

    last = regs[0]
    for nxt in regs[1:]:
        # ถ้าหักพักเที่ยง ไม่เติมช่องว่างที่คร่อมช่วงพักเที่ยง
        if nxt.start > last.end and not (lunch and last.end < lunch[1] and lunch[0] < nxt.start):
            extend(last, last.end, nxt.start)
        if nxt.end > last.end:
            last = nxt
    extend(last, last.end, win_e)
    extend(regs[0], win_s, regs[0].start)

    # รวม entry งานเดียวกันที่ต่อกันหลังเติม
    out: list[Entry] = []
    for e in sorted(entries, key=lambda e: e.start):
        p = out[-1] if out else None
        if (p and p.category == e.category == REGULAR and p.end == e.start
                and (p.contract, p.site, p.role, p.job) == (e.contract, e.site, e.role, e.job)):
            p.end = e.end
            p.note = "; ".join(filter(None, [p.note, e.note]))
        else:
            out.append(e)
    return out, sorted(added)


def mapping_warnings(entries: list[Entry], cfg: dict) -> list[Warn]:
    """entry ที่น่าจะกรอก Workday ไม่ได้: เลขสัญญาผิดรูปแบบ หรือ Role ไม่อยู่ในการจับคู่ของ team config"""
    wd = cfg.get("workday") or {}
    warns, seen = [], set()
    for e in entries:
        if e.day_type == LEAVE:
            continue  # วันลาไม่กรอกอยู่แล้ว
        role = repr(e.role) if e.role else "(ไม่มี Role)"
        if e.contract == INHOUSE:
            codes = wd.get("inhouse_code_by_role")
            problem = (codes is not None and e.role not in codes
                       and ("UNKNOWN_ROLE", f"Inhouse Role {role} ไม่มีใน inhouse_code_by_role ของ team config "
                                            "จะกรอกไม่ได้ (แก้ Role ใน Excel หรือให้คนดูแลทีมเพิ่ม)"))
        elif not CONTRACT_ID_RE.fullmatch(e.contract):
            problem = ("CONTRACT_FORMAT", f"เลขสัญญา {e.contract!r} ไม่ใช่รูปแบบปกติ (เช่น A1234567) "
                                          "อาจหา Time Type ใน Workday ไม่เจอ")
        else:
            tasks = wd.get("project_task_by_role")
            problem = (tasks is not None and e.role not in tasks and not wd.get("default_project_task")
                       and ("UNKNOWN_ROLE", f"{e.contract} Role {role} ไม่มีใน project_task_by_role ของ team config"))
        if problem and (e.date, *problem) not in seen:
            seen.add((e.date, *problem))
            warns.append(Warn(e.date, *problem))
    return warns


def build_entries(days: list[Day], contracts: dict[str, ContractInfo], cfg: dict,
                  first: dt.date, last: dt.date) -> tuple[list[Entry], list[Warn]]:
    """แปลงข้อมูลรายวันเป็น entry เฉพาะวันที่ first..last

    days ควรมีวันก่อน first 1 วันด้วย เพื่อรับช่วงที่ข้ามเที่ยงคืนเข้ามาวันแรก
    """
    ignore = {t.strip().casefold() for t in cfg.get("ignore_free_text", [])}
    in_range = lambda d: first <= d <= last  # noqa: E731
    days = sorted(days, key=lambda d: d.date)
    pieces: dict[dt.date, list[tuple[int, int, Block]]] = {d.date: [] for d in days}
    spills = []
    warns: list[Warn] = []

    for day in days:
        blocks, free = split_cell(day.text)
        if day.day_type != HOLIDAY:
            for text in free:
                if text.casefold() not in ignore:
                    warns.append(Warn(day.date, "FREE_TEXT", f"ข้อความนอกบล็อก: {_oneline(text)}"))
        for b in blocks:
            ranges = b.ranges
            if not ranges:
                detail = f"อ่านเวลาไม่ได้: {b.time_text!r}" if b.has_time_line else "ไม่มีบรรทัด Time"
                warns.append(Warn(day.date, "NO_TIME", f"{b.contract_id or '?'} {b.role}: {detail}"))
                continue
            if b.contract_id != INHOUSE and b.contract_id not in contracts:
                warns.append(Warn(day.date, "UNKNOWN_CONTRACT",
                                  f"ไม่พบ {b.contract_raw!r} ในชีต Contract"))
            for s, e in ranges:
                if e > s:
                    pieces[day.date].append((s, e, b))
                else:
                    pieces[day.date].append((s, DAY, b))
                    spills.append((day.date, s, e, b))
            covered = [seg for s, e in ranges for seg in day_segments(s, e)]
            uncovered = [(s, e) for s, e in b.job_ranges
                         if not all(any(cs <= js and je <= ce for cs, ce in covered)
                                    for js, je in day_segments(s, e))]
            if uncovered:
                text = ", ".join(f"{fmt(s)}-{fmt(e)}" for s, e in uncovered)
                warns.append(Warn(day.date, "TIME_IN_JOB_TEXT",
                                  f"{b.contract_id}: พบเวลา {text} ในบรรทัด Job แต่ไม่อยู่ในบรรทัด Time (อาจเป็น OT ที่ไม่ได้นับ)"))

    # ส่วนหลังเที่ยงคืนย้ายไปวันถัดไป ตัดส่วนที่วันถัดไปเขียนไว้แล้วออกกันนับซ้ำ
    for src, s, e, b in spills:
        target = src + dt.timedelta(days=1)
        existing = [(ps, pe) for ps, pe, _ in pieces.get(target, [])]
        remaining = subtract((0, e), existing)
        msg = (f"{b.contract_id} {fmt(s)}-{fmt(e)} แยกเป็น {fmt(s)}-24:00 ({src:%d/%m})"
               f" + 00:00-{fmt(e)} ({target:%d/%m})")
        if remaining != [(0, e)]:
            msg += "; วันถัดไปมีช่วงเวลานี้อยู่แล้ว ตัดส่วนที่ซ้ำออก"
        if target in pieces and in_range(target):
            pieces[target].extend((rs, re_, b) for rs, re_ in remaining)
        elif remaining:
            msg += "; วันถัดไปอยู่นอกช่วงที่เลือก ยังไม่ได้สร้าง entry"
        warns.append(Warn(src if in_range(src) else target, "CROSS_MIDNIGHT", msg))

    entries: list[Entry] = []
    min_regular = cfg["min_regular_hours"] * 60
    for day in days:
        if not in_range(day.date):
            continue
        merged = _merge_pieces(pieces[day.date])
        for (s1, e1, b1), (s2, e2, b2) in zip(merged, merged[1:]):
            if s2 < e1:
                warns.append(Warn(day.date, "OVERLAP",
                                  f"{b1.contract_id} {fmt(s1)}-{fmt(e1)} ทับกับ {b2.contract_id} {fmt(s2)}-{fmt(e2)}"))
        day_entries = [Entry(day.date, day.day_type, b.contract_id, b.site, b.role, ss, ee, cat, b.job)
                       for s, e, b in merged
                       for ss, ee, cat in split_regular_ot(s, e, day.day_type, cfg)]
        regular = union_minutes([(x.start, x.end) for x in day_entries if x.category == REGULAR])
        pad_from = cfg.get("pad_regular_from_hours")
        if day.day_type == WORKDAY and pad_from is not None and pad_from * 60 <= regular < min_regular:
            day_entries, added = pad_regular(day_entries, min_regular - regular, cfg)
            new_regular = union_minutes([(x.start, x.end) for x in day_entries if x.category == REGULAR])
            text = ", ".join(f"{fmt(s)}-{fmt(e)}" for s, e in added)
            warns.append(Warn(day.date, "PADDED_8H",
                              f"ชั่วโมงปกติ {regular / 60:g} ชม. เติม {text} เป็น {new_regular / 60:g} ชม."))
            regular = new_regular
        entries.extend(day_entries)
        if day.day_type == WORKDAY and regular < min_regular:
            warns.append(Warn(day.date, "UNDER_8H",
                              f"ชั่วโมงปกติ {regular / 60:g} ชม. น้อยกว่า {cfg['min_regular_hours']} ชม."))
        if day.day_type == LEAVE:
            work = ", ".join(f"{x.contract} {fmt(x.start)}-{fmt(x.end)} {x.category}" for x in day_entries)
            msg = "วันลา ต้องขอผ่าน Time Off แยก ไม่ได้กรอกในหน้านี้"
            if work:
                msg = f"วันลาแต่มีงาน {work}: จะไม่กรอกอัตโนมัติ (เบิก OT วันลาไม่ได้) ตัดสินใจและกรอกเอง"
            warns.append(Warn(day.date, "LEAVE_DAY", msg))

    warns.extend(mapping_warnings(entries, cfg))
    warns.sort(key=lambda w: w.date)
    return entries, warns


# ---------------------------------------------------------------- Excel I/O

def _header_index(ws, names: list[str]) -> dict[str, int]:
    """หาเลขคอลัมน์ (1-based) จากชื่อหัวตารางในแถว 1"""
    headers = {str(c.value).strip().casefold(): c.column for c in ws[1] if c.value is not None}
    missing = [n for n in names if n.casefold() not in headers]
    if missing:
        raise KeyError(f"ไม่พบคอลัมน์ {missing} ในชีต {ws.title!r}")
    return {n: headers[n.casefold()] for n in names}


def person_names(ws) -> list[str]:
    """ชื่อคนในแถวหัวตาราง (ทุกคอลัมน์หลังคอลัมน์วันที่)"""
    return [str(c.value).strip() for c in ws[1][1:] if c.value is not None and str(c.value).strip()]


def read_workbook(path: Path, cfg: dict, first: dt.date, last: dt.date) -> tuple[list[Day], dict[str, ContractInfo]]:
    """อ่านจากสำเนาของไฟล์ (ต้นฉบับมักเปิดค้างใน Excel และห้ามแก้)"""
    with tempfile.TemporaryDirectory() as tmp:
        copy = Path(tmp) / "schedule.xlsx"
        shutil.copy2(path, copy)
        wb = openpyxl.load_workbook(copy, data_only=True)

    person = person_of(cfg)
    days = []
    for year in sorted({first.year, last.year}):
        name = cfg["sheet"].format(year=year)
        if name not in wb.sheetnames:
            if year == last.year:
                raise SystemExit(f"ไม่พบชีต {name!r} ในไฟล์ Excel (มี {wb.sheetnames})")
            continue  # วันก่อนต้นปีที่อ่านเผื่องานข้ามเที่ยงคืน
        ws = wb[name]
        try:
            person_col = _header_index(ws, [person])[person]
        except KeyError:
            raise SystemExit(f"ไม่พบคอลัมน์ {person!r} ในชีต {name!r} ชื่อที่มี: {', '.join(person_names(ws))}")
        for row in range(2, ws.max_row + 1):
            date_cell = ws.cell(row=row, column=1)
            value = date_cell.value
            if isinstance(value, dt.datetime):
                value = value.date()
            if not isinstance(value, dt.date) or not first <= value <= last:
                continue
            person_cell = ws.cell(row=row, column=person_col)
            days.append(Day(value, classify_day(date_cell, person_cell, cfg), person_cell.value))
    days.sort(key=lambda d: d.date)

    cs = wb[cfg["contract_sheet"]]
    cols = _header_index(cs, ["Prime Contract No.", "Contract No.", "Prime Project Name", "Customer"])
    contracts = {}
    for row in range(2, cs.max_row + 1):
        prime = cs.cell(row=row, column=cols["Prime Contract No."]).value
        if not prime:
            continue
        get = lambda name: str(cs.cell(row=row, column=cols[name]).value or "").strip()  # noqa: E731
        # เลขซ้ำใช้แถวล่างสุด (แถวที่เพิ่มทีหลัง)
        contracts[str(prime).strip().upper()] = ContractInfo(get("Contract No."), get("Prime Project Name"), get("Customer"))
    return days, contracts


PREVIEW_HEADERS = ["Date", "DayType", "Contract", "Contract No", "Customer", "Project Name",
                   "Site", "Role", "Start", "End", "Hours", "Category", "Note", "Job"]
WARN_HEADERS = ["Date", "Code", "Message"]


def write_outputs(entries: list[Entry], warns: list[Warn], contracts: dict[str, ContractInfo],
                  out_dir: Path, label: str) -> tuple[Path, Path]:
    out_dir.mkdir(parents=True, exist_ok=True)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Preview"
    ws.append(PREVIEW_HEADERS)
    for e in entries:
        info = contracts.get(e.contract, ContractInfo(project=INHOUSE if e.contract == INHOUSE else ""))
        ws.append([e.date, e.day_type, e.contract, info.contract_no, info.customer, info.project,
                   e.site, e.role, fmt(e.start), fmt(e.end), round(e.hours, 2), e.category, e.note, e.job])
    ws_w = wb.create_sheet("Warnings")
    ws_w.append(WARN_HEADERS)
    for w in warns:
        ws_w.append([w.date, w.code, w.message])

    widths = {"Date": 12, "Project Name": 40, "Note": 24, "Job": 60, "Message": 100}
    for sheet in (ws, ws_w):
        sheet.freeze_panes = "A2"
        for cell in sheet[1]:
            cell.font = Font(bold=True)
            sheet.column_dimensions[cell.column_letter].width = widths.get(cell.value, 14)
        for (cell,) in sheet.iter_rows(min_row=2, max_col=1):
            cell.number_format = "yyyy-mm-dd ddd"

    preview = out_dir / f"preview_{label}.xlsx"
    wb.save(preview)
    warn_csv = out_dir / "warnings.csv"
    with open(warn_csv, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow(WARN_HEADERS)
        writer.writerows([w.date.isoformat(), w.code, w.message] for w in warns)
    return preview, warn_csv


# ---------------------------------------------------------------- CLI

def print_summary(entries: list[Entry], warns: list[Warn], days: list[Day], first: dt.date, last: dt.date) -> None:
    print(f"{'Date':<15}{'Type':<9}{'Regular':>8}{'OT':>6}  Entries")
    for day in days:
        if not first <= day.date <= last:
            continue
        mine = [e for e in entries if e.date == day.date]
        reg = union_minutes([(e.start, e.end) for e in mine if e.category == REGULAR]) / 60
        ot = union_minutes([(e.start, e.end) for e in mine if e.category == OT]) / 60
        detail = ", ".join(f"{e.contract} {fmt(e.start)}-{fmt(e.end)} {e.category}" for e in mine)
        print(f"{day.date:%Y-%m-%d %a} {day.day_type:<9}{reg:>8g}{ot:>6g}  {detail}")
    print(f"\nคำเตือน {len(warns)} รายการ")
    for w in warns:
        print(f"  {w.date:%Y-%m-%d} {w.code:<17} {w.message}")
    if any(w.code == "UNDER_8H" for w in warns):
        print("\n** มีวันทำงานที่ชั่วโมงปกติไม่ครบ ตรวจก่อนนำไปกรอก Workday **")


def parse_args(argv=None):
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    g = p.add_mutually_exclusive_group(required=True)
    g.add_argument("--week", type=dt.date.fromisoformat, help="วันแรกของสัปดาห์ (YYYY-MM-DD) อ่าน 7 วัน")
    g.add_argument("--month", help="ทั้งเดือน (YYYY-MM)")
    p.add_argument("--person", help="ชื่อหัวคอลัมน์ของคน (ทับค่า person_column ใน config)")
    p.add_argument("--config", type=Path, help="config ของผู้ใช้ (ค่าตั้งต้นอยู่ใน data dir)")
    p.add_argument("--excel", type=Path, help="ทับค่า excel_path ใน config")
    return p.parse_args(argv)


def main(argv=None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")  # ข้อความ SystemExit
    args = parse_args(argv)
    cfg = load_config(args.config)
    require_team(cfg)
    if args.person:
        cfg["person_column"] = args.person
    person = person_of(cfg)
    if args.week:
        first, last, label = args.week, args.week + dt.timedelta(days=6), args.week.isoformat()
    else:
        first = dt.date.fromisoformat(f"{args.month}-01")
        last = (first.replace(day=28) + dt.timedelta(days=4)).replace(day=1) - dt.timedelta(days=1)
        label = args.month

    excel = Path(args.excel or find_excel(cfg))
    print(f"อ่าน {excel} (แก้ไขล่าสุด {dt.datetime.fromtimestamp(excel.stat().st_mtime):%Y-%m-%d %H:%M})")
    days, contracts = read_workbook(excel, cfg, first - dt.timedelta(days=1), last)
    entries, warns = build_entries(days, contracts, cfg, first, last)
    preview, warn_csv = write_outputs(entries, warns, contracts, output_dir(cfg), label)

    print(f"{person}: {first} ถึง {last}\n")
    print_summary(entries, warns, days, first, last)
    print(f"\nบันทึก {preview}\nบันทึก {warn_csv}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
