# LG26 slip check / ตรวจสลิปค่าค่าย LG26

Pre-screens payment slips and tells you which ones to check by hand in your bank app.
Works **offline** by default. Optionally it also checks each slip with **Slip2Go**
(paid service) to confirm the money really went to the faculty account, and how much.

คัดกรองสลิปโอนเงินค่าค่าย บอกว่าสลิปไหนต้องเช็คเองในแอปธนาคาร ปกติทำงาน**ออฟไลน์**
และเลือกเปิดให้เช็คกับ **Slip2Go** ได้ (เสียเงิน) เพื่อยืนยันว่าเงินเข้าบัญชีคณะจริงและยอดถูกต้อง

---

## 1. One-time setup / ติดตั้งครั้งแรก (ทำครั้งเดียว)

Requires Python 3.10+ / ต้องมี Python 3.10 ขึ้นไป

**Windows (PowerShell)**
```powershell
cd slip_check
python -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

**macOS (Terminal)**
```bash
cd slip_check
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Create your own settings file from the template (it is git-ignored, so your account
details never get committed):
สร้างไฟล์ตั้งค่าของตัวเองจากไฟล์ตัวอย่าง (ไฟล์นี้ไม่ถูก commit ขึ้น git):

- Windows: `copy config.example.toml config.toml`
- macOS: `cp config.example.toml config.toml`

Open `config.toml` and set `fee` (camp fee in THB, e.g. `fee = 1500`).
เปิด `config.toml` แล้วใส่ `fee` (ค่าค่ายต่อคน หน่วยบาท)

### 1b. Slip2Go setup (optional) / ตั้งค่า Slip2Go (ถ้าใช้)

1. Buy a Slip2Go package. In the Slip2Go dashboard → **API Connect**, copy:
   - the **Secret Key** → paste it as the only line in a new file `slip_check/slip2go_secret.txt`
   - the **API URL** → `api_base_url` in `config.toml`
   - (recommended) set the **IP whitelist** to your own IP

   ซื้อแพ็กเกจ Slip2Go แล้วไปที่ API Connect คัดลอก Secret Key ใส่ไฟล์ `slip2go_secret.txt`
   และ API URL ใส่ `api_base_url`
2. Fill the faculty account in `config.toml` under `[slip2go.receiver]`, exactly as your bank app shows it,
   **without** title (no นาย/นาง/Mr). Fill **both** Thai and English names if you know them.
   The bank may return either one, often cut short. That is handled.
   ใส่บัญชีคณะใน `[slip2go.receiver]` ให้ตรงกับในแอปธนาคาร ไม่ต้องมีคำนำหน้า ใส่ทั้งชื่อไทยและอังกฤษ
   ```toml
   [slip2go.receiver]
   bank_code = "006"                       # 006 KTB, 004 KBANK, 014 SCB, 002 BBL, 025 BAY ...
   account_number = "1234567890"           # digits only
   account_name_th = "คณะวิศวกรรมศาสตร์ มหาวิทยาลัย..."
   account_name_en = "Faculty of Engineering ..."
   ```
3. **First test with 2 slips**: set `enabled = true` and `max_calls_per_run = 2`, then run (section 2).
   Open the **Slips** sheet and check that `receiver_check` starts with ✔ and the amounts are right.
   Then set `max_calls_per_run = 0` (no cap).
   **ทดสอบ 2 สลิปก่อน**: ตั้ง `enabled = true`, `max_calls_per_run = 2` รันแล้วดูชีต Slips
   ว่า receiver_check เป็น ✔ และยอดเงินถูก แล้วค่อยตั้ง `max_calls_per_run = 0`

**Never share `slip2go_secret.txt`.** / **ห้ามแชร์ไฟล์ `slip2go_secret.txt`**

### 1c. Automatic download from the backoffice (optional) / ดาวน์โหลด zip อัตโนมัติ (ถ้าใช้)

1. In `config.toml` under `[backoffice]`, set `auto_download = true` and paste the backoffice
   "export payment zip" URL into `export_url`. To find it: in Chrome open DevTools (F12) → **Network**,
   click the export button in the backoffice, click the request, copy its **Request URL**.
   ใน `config.toml` ส่วน `[backoffice]` ตั้ง `auto_download = true` และใส่ URL ปุ่ม export zip
   (หาได้จาก F12 → Network → กดปุ่ม export → คัดลอก Request URL)
2. The first run asks for your login cookie. Log in to the backoffice in Chrome → F12 →
   **Application** → **Cookies** → copy the value of **`bo_session`** and paste it (input is hidden).
   It is saved in `backoffice_session.txt`. A login lasts only a few hours; when it runs out the
   tool simply asks again.
   ครั้งแรกโปรแกรมจะขอ cookie: ล็อกอิน backoffice → F12 → Application → Cookies → คัดลอกค่า
   `bo_session` มาวาง (ไม่แสดงบนจอ) cookie หมดอายุภายในไม่กี่ชั่วโมง หมดแล้วโปรแกรมจะถามใหม่

**Never share `backoffice_session.txt`: it logs in to the backoffice as you.**
**ห้ามแชร์ `backoffice_session.txt` เพราะใช้ล็อกอิน backoffice แทนเราได้**

---

## 2. Every time new slips come in / ทุกครั้งที่โหลดข้อมูลใหม่

1. **With auto download:** skip this step, the run downloads the zip itself (paste a new
   `bo_session` cookie if it asks). **Without:** download the zip from the website into `inbox/`,
   keep the name `payment-slips-YYYY-MM-DD.zip`.
   **ถ้าเปิดดาวน์โหลดอัตโนมัติ** ข้ามขั้นนี้ได้ (ถ้าถาม cookie ให้วางใหม่)
   **ถ้าไม่เปิด** ดาวน์โหลด zip จากเว็บ วางใน `inbox/`
2. **Close `slip_check.xlsx` in Excel.** / **ปิดไฟล์ Excel ก่อน**
3. Run / รัน:
   - Windows: `.venv\Scripts\Activate.ps1` then `python check_slips.py`
   - macOS: `source .venv/bin/activate` then `python check_slips.py`

   With Slip2Go on, it shows how many **new** slips it will send and your remaining tokens, then asks `y/N`.
   ถ้าเปิด Slip2Go จะบอกจำนวนสลิปใหม่ที่จะส่งและ token ที่เหลือ แล้วถาม y/N
4. Open `output/slip_check.xlsx`:
   - **To check**: work top to bottom. Money problems come first (wrong account, not found,
     fraud, under/overpaid), then other flags, then the random sample.
   - Check each one in your bank app, then fill `manual_status` / `checked_by` / `note` **in the All sheet**
     (the To check sheet only mirrors them).
   - **Light-green** rows are already confirmed by Slip2Go (right account, amount = fee, no flags).
     Nothing to do unless they are in the random sample (yellow).

   เปิด `output/slip_check.xlsx` ดูชีต **To check** ไล่จากบนลงล่าง เช็คในแอปธนาคาร
   แล้วกรอกผลใน**ชีต All** แถวสีเขียวอ่อน = Slip2Go ยืนยันแล้ว ไม่ต้องทำอะไร
5. **Underpaid** people: ask them to transfer the rest and re-upload with both slips
   (page 2 of the PDF, or both in one image). Next run shows **payment completed** in the Changes sheet.
   คนที่โอนขาด: ให้โอนส่วนที่เหลือแล้วอัปโหลดใหม่พร้อมสลิปทั้งสองใบ รอบถัดไปจะขึ้น payment completed
6. At the end: type the passbook total into **Summary → passbook total (fill in)**.
   The difference is calculated for you.
   สุดท้าย ใส่ยอดในสมุดบัญชีที่ Summary แล้วดูผลต่าง

It uses the zip it just downloaded, otherwise the newest zip in `inbox/` (by the date in the filename).
A download with the same content as a zip already in `inbox/` is not saved twice; a changed one is
saved as `payment-slips-DATE (1).zip`. Old zips are never overwritten.
- Skip the download once: `python check_slips.py --no-download`
- Use a specific zip: `python check_slips.py --zip inbox/payment-slips-2026-10-25.zip`

Your manual columns and the passbook cell are kept on every run. The previous Excel file
is copied to `output/history/` first.
ค่าที่กรอกเองถูกเก็บไว้ทุกครั้ง ไฟล์เก่าสำรองไว้ที่ `output/history/`

**Turn Slip2Go off**: set `enabled = false`. Everything works offline again, and results
already received stay visible.
ปิด Slip2Go: ตั้ง `enabled = false` ผลที่เคยเช็คแล้วยังแสดงอยู่

---

## 3. Slip2Go credits / การใช้ token

- Only the **QR text** of each slip is sent (bank code + transaction ref). No images, names or files.
  ส่งแค่ข้อความใน QR ไม่ส่งรูปหรือชื่อ
- Each slip is sent **once, ever**. Results are saved in `.cache/slip2go.json`, so re-runs and newer zips
  only send new slips. The same slip in two people's files is sent once.
  แต่ละสลิปส่งครั้งเดียว รันซ้ำไม่เสีย token สลิปซ้ำกันส่งครั้งเดียว
- Slips with a bad checksum, no QR, or a non-slip QR are **not** sent. Check those by hand.
- `max_calls_per_run` caps how many slips are sent per run. If tokens run out mid-run,
  the run stops sending, keeps what it got, and marks the rest `API_PENDING`.
- Don't delete `.cache/slip2go.json`, or slips will be paid for again.
  ห้ามลบ `.cache/slip2go.json` ไม่งั้นจะเสีย token ซ้ำ

---

## 4. What the columns mean / ความหมายคอลัมน์ (sheet All)

| column | meaning |
|---|---|
| slip_count | slip QRs found in the file (2 = top-up slip) |
| amount_paid | sum of slips confirmed by Slip2Go **and** paid to the faculty account |
| amount_detail | each slip, e.g. `1,000 + 500`; ✗ = not counted, ? = not checked |
| payment_status | ✔ PAID / UNDERPAID −500 / OVERPAID +200 / ✗ no valid slip / ⏳ not checked |
| receiver_check | ✔ bank, account, name of the receiver, or ✗ why it doesn't match |
| api_result | Slip2Go result per slip: OK / NOT FOUND / FRAUD / WRONG RECEIVER / pending |
| flags | problems (below) |

The **Slips** sheet has one row per slip (amount, time, sender, receiver, API code) for audit.

## 5. Flags / ความหมายของ flag

| Flag | Meaning | ความหมาย |
|---|---|---|
| WRONG_RECEIVER | money went to another account/name | โอนเข้าบัญชีอื่น |
| SLIP_FRAUD | Slip2Go says the slip is forged | สลิปปลอม |
| SLIP_NOT_FOUND | bank has no such transfer | ไม่พบรายการในธนาคาร |
| UNDERPAID / OVERPAID | total of their slips ≠ fee | โอนขาด / โอนเกิน |
| WRONG_DATE | transfer before `paid_after` (if `check_date`) | โอนก่อนวันที่กำหนด |
| API_PENDING | not checked with Slip2Go yet (cap, tokens, bank error) | ยังไม่ได้เช็คกับ Slip2Go |
| NO_QR | file exists but no QR found (or unreadable) | หา QR ไม่เจอ / เปิดไฟล์ไม่ได้ |
| NOT_SLIP_QR | QR is not a bank slip QR (e.g. PromptPay QR) | QR ไม่ใช่ของสลิป |
| BAD_CRC | slip QR checksum wrong → may be edited | checksum ไม่ตรง อาจถูกแก้ |
| DUPLICATE_REF | same slip as another person, or `(same file)` = same slip pasted twice | สลิปซ้ำกับคนอื่น / แปะซ้ำในไฟล์เดียว |
| DUPLICATE_IMAGE | identical file to another person's, or near-identical image where one has no readable slip QR | ไฟล์เดียวกับของคนอื่น หรือรูปเหมือนกันแต่อ่าน QR ไม่ได้ |
| FILE_MISSING | marked UPLOADED but file not in zip | บอกว่าอัปโหลดแต่ไม่มีไฟล์ |
| SLIP_CHANGED | re-uploaded since last run; stays until you set manual_status again | อัปโหลดใหม่ ต้องเช็คอีกรอบ |
| NOT_PAID | no upload yet (status, not a problem) | ยังไม่อัปโหลด |

Row colours: red = flagged · yellow = random sample, not verified · light green = confirmed by Slip2Go ·
green = you marked ✅ Verified · grey = not paid.

## 6. Random sample / สุ่มตรวจ

Clean slips get a random spot-check sample (`in_sample = yes`): at least `sample_size` people,
growing to `sample_fraction` of clean slips. People already sampled stay sampled.

To turn it off (for example when Slip2Go already verifies every slip), set `sample_size = 0` and
`sample_fraction = 0`. Then nobody is sampled, including people picked in earlier runs, and
"To check" only lists flagged rows.
ปิดการสุ่มตรวจ: ตั้ง `sample_size = 0` และ `sample_fraction = 0` ชีต To check จะเหลือเฉพาะแถวที่มี flag

## 7. Files / ไฟล์

- `inbox/`: zips you download (never modified)
- `output/slip_check.xlsx`: the main file you edit · `output/slip_check.csv`: All sheet as CSV
- `output/history/`: backups of previous Excel files
- `.cache/`: extracted files, decode cache, Slip2Go results and run state. **Don't delete it.**
  ห้ามลบ `.cache/`
- `slip2go_secret.txt`: your Slip2Go key. **Never share.**
- `backoffice_session.txt`: your backoffice login cookie. **Never share.**
- `config.toml`: your settings and account details (git-ignored). `config.example.toml` is the
  blank template that is committed.

This folder contains participants' financial documents. Don't share or upload it.
โฟลเดอร์นี้มีเอกสารการเงินของผู้สมัคร ห้ามแชร์หรืออัปโหลด
