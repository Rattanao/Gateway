# Check EDI Gateway

โปรแกรม + Claude Code skill สำหรับ**ตรวจเอกสารขาเข้าเรือ** ก่อนทำใบขนสินค้าขาเข้า
เทียบ **MANIFEST** (`.xls`/`.xlsx` จากสายเรือ) กับ **ENTER** (`.pdf` เอกสารยืนยัน/แก้ไขจากลูกค้า
มักมีข้อความสีแดงที่ลูกค้าพิมพ์ทับ) แล้วออกรายงาน `EDI.xlsx` ไฮไลต์จุดที่ไม่ตรง

*English: reconciles an ocean carrier MANIFEST (.xls) against a customer ENTER sheet (.pdf) and
produces a color-coded `EDI.xlsx` before filing a Thai customs import declaration.*

## รายงานตรวจอะไรบ้าง (ต่อ 1 B/L ย่อย)
CONSIGNEE · เลขตู้ · STATUS (CY / LCL / LCL-CFS) · TOTAL PACKAGE + ชนิดบรรจุภัณฑ์ ·
GROSS WEIGHT · MEASUREMENT · MARKS · DESCRIPTION · REEFER TEMP · DG (CLASS/UN) ·
TRANSIT/TRANSHIPMENT · SHED NO. · ประเทศปลายทาง · CARGO MOVEMENT · TAX ID ↔ NOTIFY
+ แถว TOTAL เทียบยอดรวมที่เอกสารแต่ละฉบับระบุไว้เอง
+ แถววิกฤตสำหรับ B/L ที่มีใน ENTER แต่**หายไปจาก MANIFEST** (เสี่ยงตกจากใบขน)

| สี | ความหมาย |
|---|---|
| พื้นขาว | ตรวจแล้วผ่าน |
| แดง + ⚠ | ไม่ตรง / จุดเร่งด่วน (DG, REEFER, TRANSIT ขึ้นแดงเสมอถ้าพบ) |
| ส้ม | ต้องยืนยันด้วยคน |
| เหลือง | B/L นี้ไม่มี ENTER ให้เทียบ |
| เขียว | ผ่าน (ช่องสรุป) |

## ขั้นตอนติดตั้ง (ทำครั้งเดียว)

**1. เตรียมเครื่องมือ**
- ติดตั้ง [Claude Code](https://claude.com/claude-code)
- ติดตั้ง [Python 3.10+](https://www.python.org/downloads/) (ตอนติดตั้งติ๊ก *Add python.exe to PATH*)
- ติดตั้ง [Git](https://git-scm.com/downloads) (หรือข้ามไปใช้ดาวน์โหลด ZIP ในข้อ 2)

**2. ดึงโปรแกรมมาไว้ในโฟลเดอร์ skills ของ Claude Code**

Windows (PowerShell):
```powershell
git clone https://github.com/Rattanao/Gateway.git "$env:USERPROFILE\.claude\skills\manifest-enter-edi-check"
```
macOS / Linux:
```bash
git clone https://github.com/Rattanao/Gateway.git ~/.claude/skills/manifest-enter-edi-check
```
> ไม่มี Git: กด **Code → Download ZIP** แตกไฟล์แล้วเปลี่ยนชื่อโฟลเดอร์เป็น `manifest-enter-edi-check`
> วางไว้ใน `%USERPROFILE%\.claude\skills\` (ชื่อโฟลเดอร์ต้องตรงตามนี้)

**3. ติดตั้งไลบรารี Python**
```bash
pip install openpyxl "xlrd>=2.0" pymupdf
```

**4. เปิด Claude Code ใหม่** — skill จะถูกโหลดอัตโนมัติ ไม่ต้องตั้งค่าอื่น

## วิธีใช้งาน

### แบบที่ 1: ผ่าน Claude Code (แนะนำ)
1. สร้างโฟลเดอร์งานใหม่ แล้ววางไฟล์ 2 ไฟล์ไว้ใน `input/`
   ```
   งานนี้/
   └── input/
       ├── MANIFEST.xls      (ชื่อมีคำว่า MANIFEST)
       └── ENTER.pdf         (ชื่อมีคำว่า ENTER)
   ```
2. เปิด Claude Code ในโฟลเดอร์ `งานนี้/` แล้วพิมพ์อย่างใดอย่างหนึ่ง
   - `/manifest-enter-edi-check`
   - `เทียบ MANIFEST กับ ENTER` หรือ `ตรวจสอบเอกสารขาเข้าเรือ` หรือ `ทำ EDI`
3. Claude จะรันโปรแกรม สรุปจุดที่ต้องตรวจ แล้วส่ง `EDI.xlsx` ให้

### แบบที่ 2: รันตรง ๆ (ไม่ใช้ Claude)
```bash
python scripts/build_edi.py --manifest input/MANIFEST.xls --enter input/ENTER.pdf --outdir output
```
ไม่ระบุพารามิเตอร์ก็ได้ ถ้าไฟล์อยู่ในโฟลเดอร์ปัจจุบันหรือโฟลเดอร์ `input/` (โปรแกรมหาให้เอง)

## ข้อควรระวัง
- **อ่านรายงานทุกครั้ง อย่าเชื่อสรุปอย่างเดียว** — ถ้ารายงานขึ้นแดงเกือบทุกช่องทุก B/L แปลว่าโปรแกรมอ่าน
  ENTER ฉบับนั้นไม่ได้ (รูปแบบเอกสารใหม่) ไม่ใช่ข้อมูลผิดจริง ให้แจ้งผู้ดูแลหรือตรวจด้วยตา
- ข้อความที่ลูกค้าพิมพ์ทับใน ENTER (STATUS, DG, REEFER, TRANSIT) ถือเป็นคำสั่งจริงของลูกค้า
  ห้ามละเลย — โปรแกรมดึงให้อัตโนมัติ แต่ควรเปิด PDF ยืนยันอีกครั้งในจุดที่ขึ้นแดง
- โปรแกรมบันทึก `EDI.xlsx` ทับไฟล์เดิมโดยไม่ถาม ถ้าอยากเก็บฉบับเก่าให้ copy ออกก่อน หรือใช้ `--outdir`
  ไปโฟลเดอร์อื่น (ถ้าไฟล์เดิมเปิดค้างใน Excel จะบันทึกไม่ได้ ต้องปิดก่อน)
- **ห้าม commit ไฟล์งานจริง (MANIFEST / ENTER / EDI.xlsx) ขึ้น repo นี้** เพราะเป็น repo สาธารณะ
  `.gitignore` กันไว้ให้แล้ว แต่อย่า `git add -f`

## อัปเดตเป็นเวอร์ชันล่าสุด
```bash
cd ~/.claude/skills/manifest-enter-edi-check      # Windows: cd "$env:USERPROFILE\.claude\skills\manifest-enter-edi-check"
git pull
```

## โครงสร้าง repo
```
├── SKILL.md                    เงื่อนไขการทำงานของ skill + กติกาการเทียบ/รูปแบบรายงานทั้งหมด
├── scripts/build_edi.py        ตัวโปรแกรมหลัก
├── references/format-notes.md  ตัวอย่างโครงสร้างไฟล์ MANIFEST/ENTER ที่โปรแกรมรองรับ (ข้อมูลตัวอย่างปิดชื่อแล้ว)
└── requirements.txt            ไลบรารี Python ที่ต้องใช้
```

## ปรับให้เข้ากับสายเรือ/ลูกค้ารายอื่น
รหัสนำหน้า B/L (เช่น `HASL` ของ Heung-A) ตรวจจับอัตโนมัติจากไฟล์ MANIFEST ถ้า layout ของ MANIFEST/ENTER
ต่างไป ดูหัวข้อ "If the layout doesn't match" ใน `SKILL.md` และตัวอย่างใน `references/format-notes.md`
แล้วปรับ index คอลัมน์ใน `parse_manifest()` / `parse_enter*()` ของ `scripts/build_edi.py`

เครื่องมือภายในสำหรับงานตรวจเอกสารขาเข้า — ยังไม่ได้กำหนดสัญญาอนุญาต (license)
