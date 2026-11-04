"""ตัวเรียกหลักของ skill (ใช้แค่ standard library รันด้วย Python ของเครื่องได้เลย)

สร้าง venv ใน data dir และติดตั้งแพ็กเกจให้เองครั้งแรก (และเมื่อ requirements.txt เปลี่ยน)
แล้วส่งต่อให้สคริปต์ใน venv

    python timesheet.py setup [--team-config P] [--excel P] [--person N] [--show]
    python timesheet.py preview 2026-10 [--person N]
    python timesheet.py preview --week 2026-10-26
    python timesheet.py dryrun 2026-10 [--only-date D] [--skip-date D] [--accept-under-8h]
    python timesheet.py fill 2026-10 --confirmed [ตัวเลือกเดียวกับ dryrun]

fill ต้องใส่ --confirmed: ใส่ได้เฉพาะเมื่อผู้ใช้ยืนยันให้กรอกจริงแล้วเท่านั้น
ไม่มีคำสั่งใดกด Submit ใน Workday ผู้ใช้กด Submit เอง
"""
from __future__ import annotations

import hashlib
import os
import subprocess
import sys
import venv
from pathlib import Path

HERE = Path(__file__).resolve().parent
REQUIREMENTS = HERE / "requirements.txt"
COMMANDS = {"setup": "setup_user.py", "preview": "parse_schedule.py", "dryrun": "fill_workday.py",
            "fill": "fill_workday.py"}


def home_dir() -> Path:
    # ต้องตรงกับ parse_schedule.home_dir (ไฟล์นี้ import แพ็กเกจใน venv ไม่ได้)
    if os.environ.get("TIMESHEET_HOME"):
        return Path(os.environ["TIMESHEET_HOME"])
    if os.environ.get("LOCALAPPDATA"):
        return Path(os.environ["LOCALAPPDATA"]) / "workday-timesheet"
    return Path.home() / ".workday-timesheet"


def venv_python(env: Path) -> Path:
    return env / ("Scripts/python.exe" if os.name == "nt" else "bin/python")


def ensure_venv() -> Path:
    env = home_dir() / ".venv"
    py = venv_python(env)
    stamp = env / "requirements.sha256"
    digest = hashlib.sha256(REQUIREMENTS.read_bytes()).hexdigest()
    if py.exists() and stamp.exists() and stamp.read_text().strip() == digest:
        return py
    if not py.exists():
        print(f"สร้าง venv ที่ {env} ...", flush=True)
        venv.EnvBuilder(with_pip=True).create(env)
    print("ติดตั้งแพ็กเกจ (ครั้งแรกใช้เวลาสักครู่) ...", flush=True)
    subprocess.run([str(py), "-m", "pip", "install", "--disable-pip-version-check", "-q", "-r", str(REQUIREMENTS)],
                   check=True)
    stamp.write_text(digest)
    return py


def translate(argv: list[str]) -> tuple[str, list[str]]:
    """แปลงคำสั่งของ skill เป็น (สคริปต์, argument)"""
    if not argv or argv[0] not in COMMANDS:
        raise SystemExit(__doc__)
    command, rest = argv[0], list(argv[1:])
    if command != "setup" and rest and not rest[0].startswith("-"):
        rest = ["--month", rest[0], *rest[1:]]  # preview 2026-10 -> --month 2026-10
    if command == "dryrun":
        rest.append("--dry-run")
    if command == "fill":
        if "--confirmed" not in rest:
            raise SystemExit("fill ต้องใส่ --confirmed (ให้ผู้ใช้ยืนยันกรอกจริงก่อน)")
        rest.remove("--confirmed")
        if "--dry-run" in rest:
            raise SystemExit("fill ใส่ --dry-run ไม่ได้ ใช้ dryrun แทน")
    return COMMANDS[command], rest


def main(argv=None) -> int:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
    if sys.version_info < (3, 10):
        raise SystemExit(f"ต้องใช้ Python 3.10 ขึ้นไป (เครื่องนี้ {sys.version.split()[0]})")
    script, rest = translate(sys.argv[1:] if argv is None else argv)
    py = ensure_venv()
    env = {**os.environ, "PYTHONIOENCODING": "utf-8", "PYTHONUTF8": "1"}
    return subprocess.run([str(py), str(HERE / script), *rest], env=env).returncode


if __name__ == "__main__":
    sys.exit(main())
