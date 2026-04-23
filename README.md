นี่คือเอกสาร Implementation Explanation สำหรับโปรเจกต์ RPA ซึ่งทำหน้าที่ดึงข้อมูลอัตราแลกเปลี่ยนย้อนหลังจากเว็บไซต์ธนาคารแห่งประเทศไทย (BOT) ตามช่วงเวลาและสกุลเงินที่กำหนดในไฟล์ Config
**Project Overview: Daily Foreign Exchange Rates Scraper**
กระบวนการนี้ถูกออกแบบมาเพื่อดึงข้อมูลอัตราแลกเปลี่ยน (Exchange Rate) จากเว็บไซต์ BOT โดยรองรับการระบุเงื่อนไขผ่านไฟล์ JSON และทำการประมวลผลข้อมูลให้อยู่ในรูปแบบที่ต้องการก่อนส่งออกเป็นไฟล์ Excel และแจ้งเตือนผ่าน Email

1. Workflow Structure (Flowchart)
กระบวนการหลักใน Main.xaml ถูกจัดการในรูปแบบ Flowchart เพื่อควบคุม Error Handling และ Decision Making อย่างมีประสิทธิภาพ:

Get Initial Config: เริ่มต้นด้วยการอ่านไฟล์ Process_Config.json โดยใช้ Invoke Code (C#) เพื่อ Deserialize ข้อมูลลงใน Dictionary ตัวแปร Config

Prepare Environment: ตรวจสอบและสร้าง Folder ที่จำเป็น (Output, Exception, Process Path) และเรียก ClearFiles.xaml เพื่อลบไฟล์ขยะจากรอบก่อนหน้า

BOT Website Process: ขั้นตอนหลักในการเข้าถึง Browser:

จัดการ Cookies Consent อัตโนมัติ

ใช้ Inject JS Script เพื่อเลือกวันที่ (Since - To) และเลือกสกุลเงินตามที่ระบุใน Config

ดาวน์โหลดข้อมูลในรูปแบบ ZIP (CSV) และทำการ Unzip

Write Data to Excel: นำข้อมูลจาก CSV มาดึงเข้าสู่ DataTable และเขียนลงไฟล์ Excel แยกตาม Sheet ของแต่ละสกุลเงิน

End Process: เรียก SendMail.xaml เพื่อส่งไฟล์ผลลัพธ์ (หรือ Screenshot กรณี Error) และทำความสะอาดไฟล์ในเครื่อง

2. Component Breakdown
A. SubProcess: ProcessCurrency
ทำหน้าที่ประมวลผลข้อมูลดิบที่ได้จากเว็บไซต์:

Regex Matching: ใช้ Regular Expression ดึงค่า "Ratio" (เช่น 1:1 หรือ 1:100) จากหัวตาราง CSV

Data Normalization: หาก Ratio ไม่ใช่ 1 ระบบจะนำค่าอัตราแลกเปลี่ยนมาหารด้วย Ratio เพื่อให้ได้ค่ามาตรฐานต่อ 1 หน่วยสกุลเงิน

Data Mapping: เก็บ DataTable ของแต่ละสกุลเงินไว้ใน Dictionary เพื่อเตรียมเขียนลง Excel

B. SubProcess: SendMail
จัดการการส่งออกข้อมูล:

Success Path: ส่งอีเมลหัวข้อ "Success" พร้อมแนบไฟล์ Excel ผลลัพธ์

Fail Path: หากเกิด System Exception ระบบจะส่งอีเมลแจ้งเตือนพร้อมแนบ Screenshot ของหน้าจอที่เกิดปัญหา

Adaptive Mailer: รองรับการส่งผ่าน Outlook (ถ้ามีในเครื่อง) หากไม่มีจะสลับไปใช้ SMTP โดยจะแสดง Form ให้ผู้ใช้กรอกรหัสผ่าน (SecureString) เพื่อความปลอดภัย

3. Key Technical Features
Dynamic Date Picker: ใช้ JavaScript Injection ในการควบคุม UI ของปฏิทินในหน้าเว็บ BOT เพื่อความแม่นยำสูงกว่าการ Click/Type แบบธรรมดา

Non-Working Day Handling: ใน Main มีการใช้ LINQ เพื่อตรวจสอบวันหยุด โดยหากวันที่กำหนดไม่มีข้อมูลจาก BOT ระบบจะระบุในช่อง Remark ของ Excel ว่าเป็น "is not working day"

Error Recovery: มีการใช้ Try Catch ล้อมรอบจุดเสี่ยง เช่น การอ่านไฟล์ Config, การประมวลผล Browser และการจัดการ Excel เพื่อให้บอทสามารถจบกระบวนการ (End Process) ได้เสมอไม่ว่าจะสำเร็จหรือล้มเหลว

4. Configuration (Process_Config.json)
บอทจะดึงค่าสำคัญจากไฟล์นี้เพื่อให้ผู้ใช้ปรับเปลี่ยนได้โดยไม่ต้องแก้ Code:

BOT_website: URL หน้าอัตราแลกเปลี่ยน

Date_Start / Date_End: ช่วงวันที่ต้องการข้อมูล

ListCurrency: รายชื่อสกุลเงินที่ต้องการ (เช่น USD, JPY, EUR)

Email_To / Email_CC: ผู้รับผลลัพธ์
