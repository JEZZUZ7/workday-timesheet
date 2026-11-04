"""ตั้งค่าผู้ใช้: เขียน config.yaml ใน data dir

    python setup_user.py                                  หา team config + ไฟล์ Excel แล้วให้เลือกชื่อด้วยเลข
    python setup_user.py --person 3                       เลือกด้วยเลขลำดับในรายชื่อ หรือใส่ชื่อก็ได้
    python setup_user.py --team-config <path> --excel <path> --person <ชื่อ>
    python setup_user.py --show                           แสดงค่าปัจจุบัน

รันจาก terminal จะถามให้พิมพ์เลข ถ้าไม่มี terminal (Claude รันให้) จะแสดงรายชื่อแล้วจบด้วย exit code 4

exit code: 0 = บันทึกแล้ว, 3 = ไม่พบ team config, 4 = ต้องเลือกชื่อ (ใส่ --person), 5 = ไม่พบไฟล์ Excel
"""
from __future__ import annotations

import argparse
import datetime as dt
import os
import shutil
import sys
import tempfile
from pathlib import Path

import openpyxl
import yaml

from parse_schedule import (TEAM_CONFIG_NAME, USER_CONFIG, excel_candidates, home_dir, load_config,
                            onedrive_roots, person_names, read_yaml, search_excel)

NO_TEAM, NEED_PERSON, NO_EXCEL = 3, 4, 5


def find_team_config(arg: Path | None, current: str | None) -> Path | None:
    if arg:
        path = Path(os.path.expandvars(str(arg))).expanduser()
        if not path.exists():
            raise SystemExit(f"ไม่พบ team config: {path}")
        return path.resolve()
    if current and Path(os.path.expandvars(current)).exists():
        return Path(os.path.expandvars(current))
    found = search_excel(TEAM_CONFIG_NAME, onedrive_roots())
    return found[0] if found else None


def this_year() -> int:
    return dt.date.today().year


def names_in(excel: Path, cfg: dict) -> tuple[str, list[str]]:
    """รายชื่อในชีตของปีนี้ (ต้นปีที่ยังไม่มีชีตใหม่ ใช้ชีตปีก่อน)"""
    with tempfile.TemporaryDirectory() as tmp:
        copy = Path(tmp) / "schedule.xlsx"
        shutil.copy2(excel, copy)
        wb = openpyxl.load_workbook(copy, data_only=True)
    year = this_year()
    sheets = [cfg["sheet"].format(year=y) for y in (year, year - 1)]
    sheet = next((s for s in sheets if s in wb.sheetnames), None)
    if sheet is None:
        raise SystemExit(f"ไม่พบชีต {sheets[0]!r} ในไฟล์ Excel (มี {wb.sheetnames}) ตรวจค่า sheet ใน team config")
    return sheet, person_names(wb[sheet])


def interactive() -> bool:
    """รันจาก terminal ที่มีคนพิมพ์ตอบได้ (Claude รันแบบไม่มี terminal จึงใช้ --person แทน)"""
    return sys.stdin.isatty()


def resolve(answer: str, names: list[str]) -> str | None:
    """เลขลำดับในรายชื่อ (1 = คนแรก) หรือชื่อ (ไม่สนตัวพิมพ์เล็ก-ใหญ่)"""
    answer = answer.strip().lstrip("﻿").strip()
    if answer.isdigit():
        return names[int(answer) - 1] if 1 <= int(answer) <= len(names) else None
    return next((n for n in names if n.casefold() == answer.casefold()), None)


def show_names(sheet: str, names: list[str], current: str | None = None) -> None:
    print(f"ชื่อในชีต {sheet!r}:")
    for i, name in enumerate(names, 1):
        print(f"  {i:>2}. {name}{'   <- ปัจจุบัน' if name == current else ''}")


def ask_person(names: list[str], current: str | None) -> str | None:
    """ถามจนกว่าจะได้เลขหรือชื่อที่ถูก กด Enter = ใช้ชื่อปัจจุบัน"""
    default = str(names.index(current) + 1) if current in names else ""
    while True:
        try:
            answer = input(f"เลือกชื่อของคุณ (เลข 1-{len(names)}){f' [{default}]' if default else ''}: ")
        except EOFError:
            return None
        person = resolve(answer or default, names)
        if person:
            return person
        print(f"  ไม่มี {answer.strip()!r} ให้ใส่เลข 1-{len(names)}")


def save(path: Path, user: dict) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text("# config ของผู้ใช้ (สร้างโดย setup_user.py) ทับค่าใน team config\n"
                    + yaml.safe_dump(user, allow_unicode=True, sort_keys=False), encoding="utf-8")


def main(argv=None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")  # ข้อความ SystemExit
    p = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    p.add_argument("--team-config", type=Path, help=f"ไฟล์ team config (ปกติชื่อ {TEAM_CONFIG_NAME})")
    p.add_argument("--excel", type=Path, help="ไฟล์ตารางงาน ถ้าหาเองไม่เจอ")
    p.add_argument("--person", help="เลขลำดับในรายชื่อ หรือชื่อหัวคอลัมน์ของผู้ใช้ในชีต")
    p.add_argument("--show", action="store_true", help="แสดงค่าปัจจุบันแล้วจบ")
    args = p.parse_args(argv)

    path = home_dir() / USER_CONFIG
    user = read_yaml(path) if path.exists() else {}
    if args.show:
        print(f"data dir: {home_dir()}")
        print(f"config:   {path}" + ("" if path.exists() else " (ยังไม่มี: รัน setup)"))
        for key in ("team_config", "person_column", "excel_path"):
            print(f"  {key}: {user.get(key, '')}")
        return 0

    team = find_team_config(args.team_config, user.get("team_config"))
    if team is None:
        print(f"ไม่พบ team config ({TEAM_CONFIG_NAME}) ใน OneDrive: ขอไฟล์จากทีม แล้วรันใหม่พร้อม --team-config <path>")
        return NO_TEAM
    user["team_config"] = str(team)
    if args.excel:
        user["excel_path"] = str(Path(args.excel).resolve())
    save(path, user)  # เก็บ team config ไว้ก่อน รอบถัดไปไม่ต้องหาใหม่
    cfg = load_config(path)

    if args.excel:
        excel = Path(user["excel_path"])
        if not excel.exists():
            raise SystemExit(f"ไม่พบ {excel}")
    else:
        found = excel_candidates(cfg)
        if not found:
            print(f"หาไฟล์ {cfg['excel_name']!r} ไม่เจอ: ใน Teams/SharePoint ให้กด 'Add shortcut to My files' "
                  "ที่โฟลเดอร์ที่มีไฟล์ (ไม่ใช่ที่ตัวไฟล์) แล้วรอ OneDrive sync หรือรันใหม่พร้อม --excel <path>")
            return NO_EXCEL
        excel = found[0]
        if len(found) > 1:
            print("เจอหลายไฟล์ ใช้ไฟล์ที่แก้ไขล่าสุด (ระบุเองได้ด้วย --excel):")
            for f in found:
                print(f"  {f}")
        user.pop("excel_path", None)  # หาเองได้ ไม่ต้องฝัง path

    sheet, names = names_in(excel, cfg)
    current = user.get("person_column")
    reused = False
    if args.person:
        person = resolve(args.person, names)
        if not person:
            print(f"ไม่พบ {args.person!r} ในชีต {sheet!r}")
            show_names(sheet, names, current)
            print("รันใหม่พร้อม --person <เลขหรือชื่อ>")
            return NEED_PERSON
    elif interactive():
        show_names(sheet, names, current)
        person = ask_person(names, current)
        if not person:
            return NEED_PERSON
    elif current in names:
        person, reused = current, True
    else:
        show_names(sheet, names, current)
        print("รันใหม่พร้อม --person <เลขหรือชื่อ>")
        return NEED_PERSON

    user["person_column"] = person
    save(path, user)
    print(f"บันทึก {path}\n  team config: {team}\n  Excel: {excel}\n  person_column: {person}")
    if reused:
        print("(ใช้ชื่อเดิมใน config: เปลี่ยนด้วย setup --person <เลขหรือชื่อ>)")
    return 0


if __name__ == "__main__":
    sys.exit(main())
