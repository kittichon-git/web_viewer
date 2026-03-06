import os
import re
import requests
from datetime import datetime
from flask import Flask, render_template, request, jsonify
from jinja2 import Template

# ค้นหาตำแหน่งของไฟล์สคริปต์นี้
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, 'templates')

app = Flask(__name__, template_folder=TEMPLATE_DIR)

# ตั้งค่าบันทึกไฟล์รายงานไว้ที่ d:\web viewer
OUTPUT_DIR = r"d:\web viewer"

def parse_filename(filename):
    # รูปแบบ: 0317 มุกดาหาร... หรือ 0317สุราษฎร์ธานี...
    name = filename.replace('.pdf', '').strip()
    
    # ใช้ Regex ดึงตัวเลข 4 หรือ 8 หลักแรกออกมา
    match = re.search(r'^(\d{4,8})(.*)', name)
    if not match:
        return None
        
    date_raw = match.group(1)
    remaining = match.group(2).strip()
    
    # แยกส่วนที่เหลือด้วยช่องว่าง
    parts = remaining.split(' ')
    province = parts[0] if len(parts) > 0 and parts[0] else "ไม่ระบุจังหวัด"
    agency = parts[1] if len(parts) > 1 else "ไม่ระบุหน่วยงาน"
    items = " ".join(parts[2:]) if len(parts) > 2 else "ไม่ระบุรายการ"
    
    months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
              "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
    
    try:
        # รองรับทั้งแบบ 0317 (4 หลัก) และ 03172569 (8 หลัก)
        if len(date_raw) >= 4:
            mm = int(date_raw[0:2])
            dd = int(date_raw[2:4])
            # ถ้าไม่มีปี ให้ใช้ปีปัจจุบัน (พ.ศ.)
            if len(date_raw) >= 8:
                yy = int(date_raw[4:])
            else:
                yy = 2569 # ค่าเริ่มต้นตามปีปัจจุบันในระบบคุณ
            
            date_thai = f"{dd} {months[mm-1]} {yy}"
            sort_key = f"{yy}{mm:02d}{dd:02d}"
        else:
            return None
    except Exception as e:
        print(f"Error parsing date from {filename}: {e}")
        return None
        
    return {
        "date_thai": date_thai,
        "province": province,
        "agency": agency,
        "items": items,
        "sort_key": sort_key
    }

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    data = request.json
    script_url = data.get('script_url') # ลิงก์ Web App จาก Apps Script
    secret_key = data.get('secret_key') # รหัสผ่าน (เช่น AUCTION_SECRET_123)
    
    if not script_url or not secret_key:
        return jsonify({"status": "error", "message": "กรุณากรอกข้อมูลให้ครบถ้วน"})

    try:
        # ดึงข้อมูลจาก Apps Script Proxy (เพิ่ม Timeout เป็น 120 วินาที สำหรับโฟลเดอร์ขนาดใหญ่)
        response = requests.get(script_url, params={"key": secret_key}, timeout=120)
        res_data = response.json()
        
        if res_data.get('status') != 'success':
            return jsonify({"status": "error", "message": res_data.get('message', 'Unknown Error')})

        # คำนวณช่วงวันที่ (วันนี้ ถึง 90 วันข้างหน้า)
        now = datetime.now()
        # แปลงเป็นปี พ.ศ. สำหรับเปรียบเทียบ (ถ้าในชื่อไฟล์ใช้ พ.ศ.)
        current_year_thai = now.year + 543
        today_int = int(f"{current_year_thai}{now.month:02d}{now.day:02d}")
        
        # 90 วันข้างหน้า
        from datetime import timedelta
        future_date = now + timedelta(days=90)
        future_year_thai = future_date.year + 543
        future_int = int(f"{future_year_thai}{future_date.month:02d}{future_date.day:02d}")

        # Process ข้อมูลที่ได้รับจาก Proxy
        def process_list(raw_list, filter_90_days=False):
            processed = []
            import base64
            for f in raw_list:
                parsed = parse_filename(f['name'])
                if parsed:
                    # กรองเฉพาะ 90 วันข้างหน้า (ถ้าเปิดใช้งาน)
                    if filter_90_days:
                        file_date_int = int(parsed['sort_key'])
                        if file_date_int < today_int or file_date_int > future_int:
                            continue
                            
                    # เข้ารหัส Link เพื่อความปลอดภัย (Obfuscation)
                    parsed['obf_url'] = base64.b64encode(f['webViewLink'].encode()).decode()
                    processed.append(parsed)
            
            # เรียงตามวันที่
            processed.sort(key=lambda x: x['sort_key'])
            return processed

        new_files = process_list(res_data.get('new_files', []))
        all_files = process_list(res_data.get('all_files', []), filter_90_days=True)
        
        # สร้าง HTML ด้วย Template (ใช้ Path เต็ม)
        template_path = os.path.join(TEMPLATE_DIR, 'report_template.html')
        with open(template_path, 'r', encoding='utf-8') as f:
            template_str = f.read()
            
        now = datetime.now()
        timestamp = now.strftime("%d%m%Y%H%M")
        thai_year = now.year + 543
        
        months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
                  "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        thai_date_full = f"{now.day} {months[now.month-1]} {thai_year}"
        
        template = Template(template_str)
        html_content = template.render(
            new_files=new_files,
            all_files=all_files,
            thai_date=thai_date_full
        )
        
        # บันทึกไฟล์ที่ d:\web viewer
        filename = f"news_reports_{timestamp}.html"
        full_path = os.path.join(OUTPUT_DIR, filename)
        
        with open(full_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
            
        return jsonify({"status": "success", "path": full_path})
        
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})

if __name__ == '__main__':
    if not os.path.exists(OUTPUT_DIR):
        os.makedirs(OUTPUT_DIR)
    if not os.path.exists('templates'):
        os.makedirs('templates')
    app.run(debug=True, port=5000)
