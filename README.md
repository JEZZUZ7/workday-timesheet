# workday-timesheet (Claude Code plugin)

Plugin นี้ให้ Claude Code อ่านคอลัมน์ของคุณในไฟล์ตารางงานของทีม (Excel ที่ sync ผ่าน OneDrive)
แล้วกรอก Workday Enter Time ให้ ระบบแยก Regular และ OT เอง เติม Regular ที่ได้ 6-7 ชม. ให้ครบ 8 และข้ามวันลา
**ระบบไม่กด Submit** คุณต้องตรวจแล้วกด Submit เอง

## สิ่งที่ต้องมี
- Windows ที่มี **Microsoft Edge** และ **Python 3.10 ขึ้นไป**
  - ดาวน์โหลด Python ที่ https://www.python.org/downloads/ ตอนติดตั้งให้ติ๊ก *Add python.exe to PATH*
- **Claude Code**
- Workday ที่ตั้งภาษาเป็น **English**
- ไฟล์ตารางงานของทีมใน OneDrive
  - ใน Teams หรือ SharePoint ให้กด **Add shortcut to My files** ที่**โฟลเดอร์**ที่มีไฟล์ตารางงาน แล้วรอ OneDrive sync
  - อย่ากดที่ตัวไฟล์ เพราะจะได้แค่ไฟล์ `.url` ซึ่งอ่านข้อมูลไม่ได้
- **team config** (`timesheet.team.yaml`) ซึ่งคนดูแลทีมเตรียมไว้ ดูหัวข้อ "สำหรับคนดูแลทีม"

## ติดตั้ง
ใน Claude Code:
```
/plugin marketplace add <github-org>/<repo>
/plugin install workday-timesheet@team-timesheet
```
- ถ้า repo เป็น private ต้อง login GitHub ในเครื่องก่อน ด้วย `gh auth login` หรือ git credential
- อัปเดตเป็นเวอร์ชันใหม่ด้วย `/plugin marketplace update team-timesheet`

## ใช้งาน
พิมพ์คุยกับ Claude ได้ตามปกติ เช่น:
- "ตั้งค่า timesheet ให้หน่อย" (ทำครั้งแรกครั้งเดียว)
- "preview timesheet เดือน 2026-11"
- "กรอก timesheet เดือน 2026-11"

### Claude จะทำตามลำดับนี้
1. **ตั้งค่า** (เฉพาะครั้งแรก): สร้าง venv, หา team config และไฟล์ตารางงาน แล้วให้คุณเลือกชื่อของตัวเอง
2. **preview**: สรุปชั่วโมงและอธิบายคำเตือน
   - เช่น วันที่ชั่วโมงไม่ครบ, วันลาที่มีงาน หรือเวลาที่เขียนไว้ใน Job แต่ไม่อยู่ในบรรทัด Time
3. **dry-run**: เปิด Workday ใน Edge แล้วลองกรอกทีละ entry จากนั้นกด Cancel โดยไม่บันทึก
   - ถ้ายังไม่ได้ login ให้ **login เองในหน้าต่าง Edge**
   - **ห้ามคลิกในหน้าต่าง Edge** ระหว่างที่สคริปต์ทำงาน
4. **ถามยืนยัน** ก่อนกรอกจริงทุกครั้ง แล้วกรอกจริงด้วยการกด OK ทีละ entry
5. **คุณตรวจใน Workday แล้วกด Submit เอง**

ถ้าเจอข้อมูลผิด ให้แก้ในไฟล์ตารางงานบน Teams แล้วสั่ง preview ใหม่
สคริปต์จะไม่แก้ไฟล์ Excel ต้นฉบับ

## ข้อมูลของคุณเก็บที่ไหน
ข้อมูลทั้งหมดอยู่ในโฟลเดอร์ `%LOCALAPPDATA%\workday-timesheet\` ซึ่งอยู่นอกโฟลเดอร์ plugin จึงไม่หายเมื่ออัปเดต plugin

| ไฟล์ | |
|---|---|
| `config.yaml` | ชื่อคอลัมน์ของคุณ และ path ของ team config |
| `output\<ชื่อ>\` | preview, คำเตือน, รายงาน `fill_<เดือน>.csv`, log `filled_<เดือน>.json` และภาพหน้าจอ |
| `wd-profile\<ชื่อ>\` | session ที่ login Workday ไว้ **ห้ามแชร์หรือ copy ให้คนอื่น** |
| `.venv\` | แพ็กเกจ Python ที่สคริปต์ติดตั้งให้ |

- สคริปต์อ่านวันที่ได้ทุกรูปแบบตาม Locale ในบัญชี Workday เช่น `15/10/2026`, `10/15/2026`, `October 15, 2026` ไม่ต้องตั้งค่าเพิ่ม
- ถ้าต้องการย้ายโฟลเดอร์ข้อมูล ให้ตั้ง environment variable `TIMESHEET_HOME`

## สำหรับคนดูแลทีม: team config
ค่าที่ผูกกับองค์กรหรือ tenant **ไม่อยู่ใน repo นี้** ได้แก่:
- ชื่อไฟล์และชื่อชีตของตารางงาน
- URL ของ Workday และ path ของหน้า Enter Time
- ชื่อช่อง OT
- การจับคู่ Role กับ task และ Time Entry Code

ขั้นตอนตั้งค่า:
1. copy [team-config.example.yaml](plugins/workday-timesheet/skills/timesheet/team-config.example.yaml) แล้วตั้งชื่อเป็น `timesheet.team.yaml`
2. ใส่ค่าจริงขององค์กร
   - `enter_time_path` ได้จากการเปิดหน้า Enter Time ใน Workday แล้ว copy ส่วนของ URL ที่ต่อจาก `url`
3. วางไฟล์ไว้ใน**โฟลเดอร์เดียวกับไฟล์ตารางงาน**บน Teams
   - ทุกคนที่ทำ shortcut ของโฟลเดอร์นั้นไว้แล้วจะได้ไฟล์นี้ผ่าน OneDrive และ setup จะหาเจอเอง
   - ถ้าวางไว้ที่อื่น ผู้ใช้ต้องบอก Claude ว่าไฟล์อยู่ที่ไหน
4. ถ้าแก้ไฟล์นี้ภายหลัง ทุกคนจะได้ค่าใหม่ในการรันครั้งถัดไป ไม่ต้องอัปเดต plugin

## พัฒนา plugin
ก่อน push ให้ตรวจ manifest ด้วย:
```
claude plugin validate . --strict
```
test ไม่อยู่ใน repo เพราะใช้ข้อมูลงานจริงในการทดสอบ ผู้ดูแลเก็บไว้ในเครื่อง ถ้าต้องการ ขอได้จากผู้ดูแล

**กฎของ repo: ห้ามใส่ข้อมูลส่วนตัวหรือข้อมูลองค์กรในโค้ด, เอกสาร และข้อความ commit**
- ข้อมูลที่ห้ามใส่: ชื่อพนักงาน, อีเมล, เลขสัญญา, ชื่อลูกค้าหรือโครงการ, รายละเอียดงาน, วันลาหรือตารางงานของใคร, URL ของ tenant, path ในเครื่อง และไฟล์ Excel จริง
- ค่าที่ผูกกับองค์กรให้อ่านจาก team config
- ตัวอย่างในโค้ดหรือเอกสารต้องเป็นค่าสมมติ ห้ามดัดแปลงจากข้อมูลจริง
- `.gitignore` กัน `tests/`, `*.xlsx` และ `timesheet.team.yaml` ไว้แล้ว แต่ให้ตรวจ `git diff --cached` ทุกครั้งก่อน commit

ขั้นตอนออกเวอร์ชันใหม่:
1. เพิ่ม `version` ทั้งใน `plugins/workday-timesheet/.claude-plugin/plugin.json` และใน `.claude-plugin/marketplace.json`
2. push ขึ้น repo
3. ผู้ใช้รัน `/plugin marketplace update team-timesheet`

### โครงสร้าง
```
.claude-plugin/marketplace.json              marketplace ของทีม
plugins/workday-timesheet/
  .claude-plugin/plugin.json
  skills/timesheet/
    SKILL.md                                 ขั้นตอนและกฎที่ Claude ต้องทำตาม
    team-config.example.yaml
    scripts/
      timesheet.py                           ตัวเรียกหลัก (สร้าง venv แล้วส่งต่อ)
      setup_user.py                          ตั้งค่าผู้ใช้ (ไม่ถามแบบ interactive)
      parse_schedule.py                      อ่าน Excel -> preview + คำเตือน
      fill_workday.py                        กรอก Workday ผ่าน Edge (CDP)
      requirements.txt
```
