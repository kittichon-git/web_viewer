import re
from googleapiclient.discovery import build

class DriveService:
    def __init__(self, api_key):
        # ใช้ API Key ในการสร้าง Service โดยตรง (ไม่ต้องใช้ OAuth)
        self.service = build('drive', 'v3', developerKey=api_key)

    def extract_id(self, url):
        # ดึง Folder ID จาก URL
        match = re.search(r'[-\w]{25,}', url)
        return match.group(0) if match else url

    def list_files(self, folder_id):
        # ค้นหาไฟล์ PDF ในโฟลเดอร์ (ต้องตั้งค่า Folder เป็น Public/Anyone with link)
        query = f"'{folder_id}' in parents and mimeType = 'application/pdf' and trashed = false"
        results = self.service.files().list(
            q=query, 
            fields="files(id, name, webViewLink)",
            pageSize=1000
        ).execute()
        return results.get('files', [])

    def get_folder_files_recursive(self, folder_id):
        all_files = []
        
        # ดึงไฟล์ในโฟลเดอร์ปัจจุบัน
        files = self.list_files(folder_id)
        all_files.extend(files)
        
        # ดึงโฟลเดอร์ย่อย (Recursive)
        query = f"'{folder_id}' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        subfolders = self.service.files().list(
            q=query, 
            fields="files(id, name)",
            pageSize=100
        ).execute().get('files', [])
        
        for folder in subfolders:
            all_files.extend(self.get_folder_files_recursive(folder['id']))
            
        return all_files
