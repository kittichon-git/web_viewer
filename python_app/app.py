import os
import re
import requests
import base64
from datetime import datetime, timedelta
from flask import Flask, render_template, request, jsonify
from jinja2 import Template

# ค้นหาตำแหน่งของไฟล์สคริปต์นี้
BASE_DIR = os.path.abspath(os.path.dirname(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, 'templates')

app = Flask(__name__, template_folder=TEMPLATE_DIR)

# ตั้งค่าบันทึกไฟล์รายงานไว้ที่ d:\web viewer
OUTPUT_DIR = r"d:\web viewer"

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    data = request.json
    report_type = data.get('type', 'standard') # 'standard' หรือ 'urgent'
    report_date_str = data.get('report_date') # 'YYYY-MM-DD'
    script_url = data.get('script_url')
    secret_key = "AUCTION_INTERNAL_SECRET_999" # ปรับเป็นค่าคงที่ภายใน
    
    # Folder IDs
    folder_new = data.get('folder_new')
    folder_all = data.get('folder_all')
    folder_urgent = data.get('folder_urgent')
    
    if not script_url or not secret_key:
        return jsonify({"status": "error", "message": "กรุณากรอก WebApp URL และ Secret Key"})

    try:
        # ส่ง Parameter ให้ GAS (รวมวันที่รายงานด้วย)
        params = {
            "key": secret_key,
            "folder_new": folder_new,
            "folder_all": folder_all,
            "folder_urgent": folder_urgent,
            "report_date": report_date_str
        }
        
        # ดึงข้อมูลจาก GAS
        response = requests.get(script_url, params=params, timeout=120)
        res_data = response.json()
        
        if res_data.get('status') != 'success':
            return jsonify({"status": "error", "message": res_data.get('message', 'Proxy Error')})

        # จัดการเรื่องวันที่สำหรับหัวตาราง (เหมือนเดิม)
        if report_date_str:
            target_date = datetime.strptime(report_date_str, '%Y-%m-%d')
        else:
            target_date = datetime.now()
            
        thai_year = target_date.year + 543
        months = ["มกราคม", "กุมภาพันธ์", "มีนาคม", "เมษายน", "พฤษภาคม", "มิถุนายน",
                  "กรกฎาคม", "สิงหาคม", "กันยายน", "ตุลาคม", "พฤศจิกายน", "ธันวาคม"]
        thai_date_full = f"{target_date.day} {months[target_date.month-1]} {thai_year}"

        # ฟังก์ชันเตรียมข้อมูล (แค่ Encode URL เพราะ GAS ส่ง Data ที่ Parse มาแล้ว)
        def prepare_data(data_list):
            for f in data_list:
                f['obf_url'] = base64.b64encode(f['webViewLink'].encode()).decode()
            return data_list

        # สร้างรายงานตามประเภท
        if report_type == 'urgent':
            urgent_files = prepare_data(res_data.get('urgent_files', []))
            print(f"DEBUG: Found {len(urgent_files)} urgent files")
            template_name = 'urgent_report_template.html'
            render_data = {"urgent_files": urgent_files, "thai_date": thai_date_full}
            file_prefix = "urgent_report"
        else:
            new_files = prepare_data(res_data.get('new_files', []))
            all_files = prepare_data(res_data.get('all_files', []))
            print(f"DEBUG: Found {len(new_files)} new files, {len(all_files)} all files")
            if new_files:
                print(f"DEBUG: First New File Details: {new_files[0]}")
            
            template_name = 'report_template.html'
            render_data = {"new_files": new_files, "all_files": all_files, "thai_date": thai_date_full}
            file_prefix = "news_reports"

        # โหลด Template และสร้าง HTML
        template_path = os.path.join(TEMPLATE_DIR, template_name)
        with open(template_path, 'r', encoding='utf-8') as f:
            template_str = f.read()
            
        template = Template(template_str)
        html_content = template.render(**render_data)
        
        # บันทึกไฟล์
        timestamp = datetime.now().strftime("%d%m%Y%H%M")
        save_filename = f"{file_prefix}_{timestamp}.html"
        full_path = os.path.join(OUTPUT_DIR, save_filename)
        
        with open(full_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
            
        # ตรวจสอบว่าต้องอัปโหลดขึ้น Git หรือไม่
        push_to_git = data.get('push_to_git', False)
        git_url = ""
        
        if push_to_git:
            import subprocess
            try:
                print(f"DEBUG: Pushing {save_filename} to git...")
                # Add file
                subprocess.run(['git', 'add', save_filename], cwd=OUTPUT_DIR, check=True)
                # Commit
                commit_msg = f"Auto-generate {save_filename}"
                subprocess.run(['git', 'commit', '-m', commit_msg], cwd=OUTPUT_DIR, check=True)
                # Push
                subprocess.run(['git', 'push', 'origin', 'master'], cwd=OUTPUT_DIR, check=True)
                
                # สร้าง URL สำหรับดูผ่าน GitHub Pages
                git_url = f"https://kittichon-git.github.io/web_viewer/{save_filename}"
                print(f"DEBUG: Successfully pushed to git. URL: {git_url}")
            except subprocess.CalledProcessError as e:
                print(f"ERROR pushing to git: {e}")
                return jsonify({"status": "error", "message": f"เกิดข้อผิดพลาดในการอัปโหลดขึ้น Git: {e}"})

        return jsonify({
            "status": "success", 
            "path": full_path, 
            "sheet_url": res_data.get('sheet_url', ''),
            "git_url": git_url
        })
        
    except Exception as e:
        return jsonify({"status": "error", "message": str(e)})

if __name__ == '__main__':
    if not os.path.exists(OUTPUT_DIR): os.makedirs(OUTPUT_DIR)
    app.run(debug=True, port=5000)
