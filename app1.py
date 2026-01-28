# app1.py (v52.0 - 改用 xlwt 輸出 .xls 格式以相容舊系統)

import os
import time
import requests
import json
import traceback
from datetime import datetime
from flask import Flask, request, render_template, jsonify, Response
import google.generativeai as genai
from mimetypes import guess_type
import fitz  # PyMuPDF
import io
import csv

# --- Google Drive API 相關套件 ---
from google.auth.transport.requests import Request as GoogleRequest
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload

# --- 環境變數管理套件 ---
from dotenv import load_dotenv

# --- 載入 .env 檔案 ---
load_dotenv()

# --- 函式庫 (更換為 xlwt) ---
import pandas as pd
import xlwt  # <--- 改用這個套件來產生 .xls

# --- 設定 ---
INPUT_FOLDER = "uploads"
PDF_CONVERSION_DPI = 300
app = Flask(__name__)

app.config['UPLOAD_FOLDER'] = INPUT_FOLDER
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)

# --- API Key 設定 ---
GEMINI_API_KEY = os.getenv('GOOGLE_API_KEY')

if not GEMINI_API_KEY:
    print("⚠️ 嚴重警告：未偵測到 GOOGLE_API_KEY！程式將無法辨識發票。")
else:
    os.environ['GOOGLE_API_KEY'] = GEMINI_API_KEY
    genai.configure(api_key=GEMINI_API_KEY)

# --- Google Drive 權限設定 ---
SCOPES = ['https://www.googleapis.com/auth/drive.file']

# --- 統一發票字軌設定 ---
INVOICE_PREFIX_MAP_2025 = { 
    'PT': '21', 'HT': '21', 'KT': '21', 'MT': '21', 'RT': '21', 'TT': '21',
    'HV': '22', 'HW': '22', 'HX': '22', 'HY': '22', 'KV': '22', 'KW': '22', 'KX': '22',
    'KY': '22', 'MV': '22', 'MW': '22', 'MX': '22', 'MY': '22', 'PV': '22', 'PW': '22',
    'PX': '22', 'PY': '22', 'RV': '22', 'RW': '22', 'RX': '22', 'RY': '22', 'TV': '22',
    'TW': '22', 'TX': '22', 'TY': '22', 'MW': '22'
}

INVOICE_PREFIX_MAP_2026 = { 
    'VT': '21', 'XV': '21', 'ZX': '21', 'CA': '21', 'EC': '21', 'GE': '21', 
    'VV': '22', 'XX': '22', 'ZZ': '22', 'CC': '22', 'EE': '22', 'GG': '22', 
    'VW': '22', 'XY': '22', 'AA': '22', 'CD': '22', 'EF': '22', 'GH': '22', 
    'VX': '22', 'XZ': '22', 'AB': '22', 'CE': '22', 'EG': '22', 'GJ': '22', 
    'VY': '22', 'YA': '22', 'AC': '22', 'CF': '22', 'EH': '22', 'GK': '22'
}

# --- Google Drive 輔助函式 ---
def get_drive_service():
    """取得 Google Drive Service (OAuth 2.0)"""
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(GoogleRequest())
        else:
            if os.path.exists('credentials.json'):
                flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)
            else:
                print("❌ 錯誤：找不到 credentials.json，無法使用 Google Drive 功能")
                return None
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return build('drive', 'v3', credentials=creds)

def download_file_by_id(service, file_id, file_name):
    """下載單一檔案"""
    save_path = os.path.join(app.config['UPLOAD_FOLDER'], file_name)
    request = service.files().get_media(fileId=file_id)
    fh = io.FileIO(save_path, 'wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
    print(f"已下載: {file_name}")
    return file_name

def download_files_from_drive_folder(folder_id):
    """(舊逻辑保留) 從指定 Drive 資料夾下載所有圖片"""
    service = get_drive_service()
    if not service: return []

    print(f"正在讀取雲端資料夾 ID: {folder_id} ...")
    query = f"'{folder_id}' in parents and (mimeType contains 'image/' or mimeType = 'application/pdf') and trashed = false"
    results = service.files().list(q=query, fields="files(id, name, mimeType)").execute()
    items = results.get('files', [])

    downloaded_files = []
    if not items:
        print("資料夾內沒有檔案。")
        return []

    print(f"找到 {len(items)} 個檔案，開始下載...")
    for item in items:
        downloaded_files.append(download_file_by_id(service, item['id'], item['name']))
    
    return downloaded_files

# --- 核心函式 (Prompt 嚴禁更動) ---
def extract_data_with_gemini_vision(image_bytes: bytes, mime_type: str) -> list:
    if not GEMINI_API_KEY:
        print("[Error] 缺少 API Key，跳過辨識。")
        return []

    image_part = {"mime_type": mime_type, "data": image_bytes}
    prompt = f"""
    你是一位頂尖的台灣發票資料分析師。你的任務是從眼前的發票圖片中，精準地擷取結構化資訊。

    **最終輸出指示 (非常重要):**
    你的回覆**必須**只包含一個乾淨的 JSON 物件，其結構為 `{{ "receipts": [...] }}`。**絕對不要**在 JSON 前後或內部加入任何說明文字、分析過程、註解或 ```json ``` 標籤。

    **擷取規則:**
    1.  **[視覺辨識優先]** 直接根據圖片上的視覺資訊進行判斷，忽略微小的污漬或印刷瑕疵。特別注意帶有斜線的'0'可能被看成'8'或'6'或'9'。
    2.  **發票號碼 (invoice_number):** 格式為 "XX12345678"。
    3.  **交易日期 (date):** 擷取年月日。如果是民國年(如114,115)，請轉換為西元年(2025,2026)，以此類推。最終格式為 "YYYY-MM-DD"。
    4.  **交易時間 (time):** 擷取時:分:秒，格式為 "HH:MM:SS"。如果圖片上沒有，則為空字串 ""。
    5.  **賣方統一編號 (seller_vat):** 8位數字。直接從圖片中的賣方資訊區塊（通常有公司印章或明確標示）尋找。**忽略印章上的裝飾性符號（如梅花✳、星號*等），只提取數字**。
        * **主要線索：** 優先尋找有「賣方、「統編」、「統一編號」、「NO.」、「#」等明確標示的8位數字，而在格式代碼21的三聯式收銀機發票中，這些統一編號有時會緊接在交易日期或時間之後，並用#做連結。
        * **位置線索：** 在格式代碼21的手開三聯式發票中，賣方統編常常出現在「專用章」、「TEL」、「負責人」字樣附近。
        * **順序線索：** 在格式代碼21的三聯式收銀機發票，或格式代碼22的二聯式收銀機發票中，如果沒有明確標示，它有時是發票號碼後的第二組8位數字。
        * **修正規則：** 如果辨識出的數字串超過8位（例如 68488162716），應由左至右每8位數字為一組，逐一比對於財政部稅籍登記資料之營業人名稱(businessNm)之前兩個字，是否與本張發票有的公司名稱前兩字相同(例如68488162無稅籍登記資料，改用下一組84881627，找到「鼎祥氣體有限公司」，其前兩字與「專用章」附近出現的「鼎祥氣體有限公司」前兩字一致，得出正確答案)。
    6.  **買方統一編號 (buyer_vat):** 尋找**除了**「賣方統一編號」之外的另一組8位數字，通常位於「買受人」字樣附近。如果沒有，則為 "N/A"。
    7.  **金額總計 (total_amount):** 發票上的含稅總金額，必須是整數。優先順序為：「總計」、「縂計」> 「合計」 > 「應收金額」，且三聯式發票通常有加總關係可以核對，例如: 格式代碼21的手開三聯式發票中的"銷售額合計"3800+"營業稅"190="總計3990"。
    8.  如果任何欄位找不到，請使用 "N/A" 作為值。

    **JSON 輸出格式範例:**
    {{
      "receipts": [
        {{ "invoice_number": "MW25046739", "date": "2025-05-27", "time": "13:10", "seller_vat": "49280041", "buyer_vat": "83251000", "total_amount": 465 }}
      ]
    }}
    """
    try:
        model = genai.GenerativeModel("gemini-3-flash-preview") 
        response = model.generate_content([prompt, image_part])
        cleaned_response_text = response.text.strip()
        json_start = cleaned_response_text.find('{')
        json_end = cleaned_response_text.rfind('}')
        if json_start != -1 and json_end != -1:
            json_str = cleaned_response_text[json_start:json_end+1]
            data = json.loads(json_str)
            return data.get("receipts", [])
        else:
            print("[Gemini Vision Warning] 回應中未找到有效的 JSON 物件。")
            return []
    except Exception as e:
        print(f"[Gemini Vision Error] 解析失敗: {e}")
        return []

def is_valid_vat_number(vat: str) -> bool:
    if not vat or not vat.isdigit() or len(vat) != 8: return False
    multipliers = [1, 2, 1, 2, 1, 2, 4, 1]; total = 0
    for i in range(8): product = int(vat[i]) * multipliers[i]; total += (product // 10) + (product % 10)
    if total % 10 == 0: return True
    if vat[6] == '7' and (total + 1) % 10 == 0: return True
    return False

def correct_vat_number(vat: str) -> str:
    if is_valid_vat_number(vat): return vat
    error_indices = [i for i, char in enumerate(vat) if char in ('8', '6')]
    for i in error_indices:
        temp_vat_list = list(vat); temp_vat_list[i] = '0'
        corrected_vat = "".join(temp_vat_list)
        if is_valid_vat_number(corrected_vat): print(f"統一編號修正成功: {vat} -> {corrected_vat}"); return corrected_vat
    return vat

# --- 增強版公司查詢 (含備援) ---
def get_company_info_from_fia_api(vat_number: str) -> dict:
    if not vat_number or vat_number == 'N/A' or not vat_number.isdigit(): 
        return {"name": "N/A", "address": ""}
    
    headers = {"User-Agent": "Mozilla/5.0"}
    
    # --- 策略 A: 財政部官方 API ---
    try:
        api_url = f"https://eip.fia.gov.tw/OAI/api/businessRegistration/{vat_number}"
        response = requests.get(api_url, headers=headers, timeout=5)
        if response.status_code == 200:
            data = response.json()
            company_name = data.get("businessNm", "")
            company_address = data.get("businessAddress", "")
            if company_name:
                full_width_chars = "０１２３４５６７８９ＡＢＣＤＥＦＧＨＩＪＫＬＭＮＯＰＱＲＳＴＵＶＷＸＹＺａｂｃｄｅｆｇｈｉｊｋｌｍｎｏpqrstuvwxyz"
                half_width_chars = "0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz"
                translation_table = str.maketrans(full_width_chars, half_width_chars)
                return {
                    "name": company_name.translate(translation_table), 
                    "address": company_address.translate(translation_table)
                }
    except Exception: pass

    # --- 策略 B: g0v API (通用解析版) ---
    try:
        # print(f"g0v 查詢: {vat_number}")
        g0v_url = f"https://company.g0v.ronny.tw/api/show/{vat_number}"
        response = requests.get(g0v_url, headers=headers, timeout=10)
        
        if response.status_code == 200:
            result = response.json()
            
            # g0v 回傳的結構通常是 {"data": { "來源A": {...}, "來源B": {...} }}
            if "data" in result and isinstance(result["data"], dict):
                all_sources = result["data"]
                
                # 我們遍歷所有來源 (例如 "財政部", "經濟部商業司"...)
                for source_name, info in all_sources.items():
                    if isinstance(info, dict):
                        # 嘗試抓取各種可能的名稱欄位
                        name = (info.get("公司名稱") or 
                                info.get("商業名稱") or 
                                info.get("營業人名稱") or  # <--- 針對普客二四這種財政部資料
                                info.get("名稱"))
                        
                        if name:
                            return {"name": name, "address": ""}
                            
    except Exception as e:
        print(f"g0v 查詢失敗: {e}")

    return {"name": "查無資料(連線失敗)", "address": ""}

def enrich_and_finalize_data(raw_receipts: list, source_filename: str) -> list:
    final_receipts = []
    for raw_receipt in raw_receipts:
        try: total = int(raw_receipt.get("total_amount", 0))
        except (ValueError, TypeError): total = 0
        tax_exclusive_amount = round(total / 1.05) if total > 0 else 0; tax_amount = total - tax_exclusive_amount if total > 0 else 0
        date_part = raw_receipt.get("date", "N/A"); day_of_week = "N/A"
        
        selected_map = INVOICE_PREFIX_MAP_2025
        if date_part != "N/A":
            try: 
                dt = datetime.strptime(date_part, '%Y-%m-%d')
                weekdays = ["一", "二", "三", "四", "五", "六", "日"]
                day_of_week = weekdays[dt.weekday()]
                if dt.year == 2026: selected_map = INVOICE_PREFIX_MAP_2026
            except ValueError: day_of_week = "格式錯誤"
        
        seller_vat = raw_receipt.get("seller_vat", "N/A"); buyer_vat = raw_receipt.get("buyer_vat", "N/A")
        corrected_seller_vat = correct_vat_number(seller_vat); corrected_buyer_vat = correct_vat_number(buyer_vat)
        seller_info = get_company_info_from_fia_api(corrected_seller_vat); time.sleep(0.5); buyer_info = get_company_info_from_fia_api(corrected_buyer_vat)
        invoice_number = raw_receipt.get("invoice_number", "N/A"); prefix = invoice_number[:2].upper() if invoice_number and len(invoice_number) == 10 else ""
        
        format_code_str = selected_map.get(prefix, '25');
        try: format_code_int = int(format_code_str)
        except (ValueError, TypeError): format_code_int = 25
        receipt = {
            "統一發票號碼": invoice_number, "格式": format_code_int,
            "交易日期": date_part, "星期": day_of_week, "交易時間": raw_receipt.get("time", ""),
            "賣方統一編號": corrected_seller_vat, "賣方名稱": seller_info["name"], "賣方營業地址": seller_info["address"],
            "買方統一編號": corrected_buyer_vat, "買方名稱": buyer_info["name"], "買方營業地址": buyer_info["address"],
            "未稅金額": tax_exclusive_amount, "進項稅額": tax_amount, "金額總計": total, "來源檔案": source_filename,
        }
        final_receipts.append(receipt)
    return final_receipts

# --- Routes ---
@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

# --- 新增：列出 Google Drive 檔案清單 ---
@app.route('/list_drive_files', methods=['POST'])
def list_drive_files():
    try:
        json_data = request.json or {}
        folder_id = json_data.get('folder_id') or os.getenv('GDRIVE_FOLDER_ID')
        if not folder_id: return jsonify({"error": "未提供 Folder ID 且 .env 中也未設定"}), 400
        
        service = get_drive_service()
        if not service: return jsonify({"error": "無法連接 Google Drive"}), 500

        print(f"正在讀取雲端資料夾清單 ID: {folder_id} ...")
        query = f"'{folder_id}' in parents and (mimeType contains 'image/' or mimeType = 'application/pdf') and trashed = false"
        # 增加 pageSize 讀取更多檔案 (預設100)
        results = service.files().list(q=query, pageSize=1000, fields="files(id, name, mimeType)").execute()
        items = results.get('files', [])
        return jsonify({"files": items})
    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

# --- 修改：從 Google Drive 匯入處理 (支援選擇檔案) ---
@app.route('/process_drive_folder', methods=['POST'])
def process_drive_folder():
    try:
        json_data = request.json or {}
        # 檢查是否有使用者指定的檔案清單
        selected_files = json_data.get('selected_files') # 預期格式: [{'id': '...', 'name': '...'}, ...]

        downloaded_files = []
        service = get_drive_service()
        
        if selected_files:
            # === 新流程：只下載指定的檔案 ===
            if not service: return jsonify({"error": "無法連接 Google Drive"}), 500
            print(f"收到指定處理檔案: {len(selected_files)} 個")
            for item in selected_files:
                try:
                    fname = download_file_by_id(service, item['id'], item['name'])
                    downloaded_files.append(fname)
                except Exception as e:
                    print(f"下載失敗 {item['name']}: {e}")
        else:
            # === 舊流程：下載資料夾全部 (Fallback) ===
            folder_id = json_data.get('folder_id') or os.getenv('GDRIVE_FOLDER_ID')
            if not folder_id: return jsonify({"error": "未提供 Folder ID 且 .env 中也未設定"}), 400
            downloaded_files = download_files_from_drive_folder(folder_id)

        if not downloaded_files:
            return jsonify({"error": "雲端資料夾為空、下載失敗或未選擇檔案"}), 404

        all_results = []
        # 處理檔案 (邏輯維持不變)
        for filename in downloaded_files:
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            try:
                mime_type = guess_type(filepath)[0]; raw_receipts = []
                if mime_type in ["image/jpeg", "image/png", "image/webp"]:
                    with open(filepath, "rb") as f: image_bytes = f.read()
                    raw_receipts = extract_data_with_gemini_vision(image_bytes, mime_type)
                elif mime_type == "application/pdf":
                    doc = fitz.open(filepath)
                    for page_num, page in enumerate(doc):
                        print(f"處理 PDF '{filename}' 的第 {page_num + 1} 頁...")
                        pix = page.get_pixmap(dpi=PDF_CONVERSION_DPI); img_bytes = pix.tobytes("png")
                        page_receipts = extract_data_with_gemini_vision(img_bytes, "image/png"); raw_receipts.extend(page_receipts)
                    doc.close()
                else: 
                    print(f"略過非支援檔案: {filename}"); continue
                    
                finalized_results = enrich_and_finalize_data(raw_receipts, filename); all_results.extend(finalized_results)
            except Exception as e:
                print(f"處理檔案 {filename} 時發生錯誤: {e}"); traceback.print_exc()
                all_results.append({"來源檔案": filename, "統一發票號碼": f"處理失敗: {e}",})
            finally:
                if os.path.exists(filepath): os.remove(filepath)

        return jsonify({"results": all_results})

    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

@app.route('/process_image', methods=['POST'])
def process_image():
    uploaded_files = request.files.getlist('receipt_image');
    if not uploaded_files or uploaded_files[0].filename == '': return jsonify({"error": "沒有選擇任何檔案"}), 400
    all_results = []
    for file in uploaded_files:
        filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename); file.save(filepath)
        try:
            mime_type = guess_type(filepath)[0]; raw_receipts = []
            if mime_type in ["image/jpeg", "image/png", "image/webp"]:
                with open(filepath, "rb") as f: image_bytes = f.read()
                raw_receipts = extract_data_with_gemini_vision(image_bytes, mime_type)
            elif mime_type == "application/pdf":
                doc = fitz.open(filepath)
                for page_num, page in enumerate(doc):
                    print(f"處理 PDF '{file.filename}' 的第 {page_num + 1} 頁...")
                    pix = page.get_pixmap(dpi=PDF_CONVERSION_DPI); img_bytes = pix.tobytes("png")
                    page_receipts = extract_data_with_gemini_vision(img_bytes, "image/png"); raw_receipts.extend(page_receipts)
                doc.close()
            else: raise Exception(f"不支援的檔案格式: {mime_type}")
            finalized_results = enrich_and_finalize_data(raw_receipts, file.filename); all_results.extend(finalized_results)
        except Exception as e:
            print(f"--- 處理檔案 {file.filename} 時發生嚴重錯誤 ---"); traceback.print_exc()
            all_results.append({"來源檔案": file.filename, "統一發票號碼": f"處理失敗: {e}",})
        finally:
            if os.path.exists(filepath): os.remove(filepath)
    return jsonify({"results": all_results})

@app.route('/generate_gv', methods=['POST'])
def generate_gv():
    # --- 修改說明: 使用 xlwt 產生 .xls 檔案 (Excel 97-2003) ---
    json_data = request.json
    results = json_data.get('results', [])
    account_payable_code = json_data.get('account_payable_code', '')

    if not results: return jsonify({"error": "沒有資料可供下載"}), 400

    def convert_format_code_to_type(code): return "Q" if str(code) == "22" else "I"

    # 表頭欄位
    header_row = [
        "序號", "公司別", "發票號碼", "稅籍編號", "統一編號", "記帳點",
        "發票/憑證類別", "格式代號", "單據憑證日期", 
        "傳票日期", "申報年月", "銷售人統一編號", "銷售人名稱", "課稅別", "進貨折讓區分", 
        "未稅金額", "進項稅額", "金額總計", "進項稅性質別", "扣抵代號",
        "彙總張數", "彙加註記", "應付立帳號碼", "備註" 
    ]

    data_for_excel = []; data_for_excel.append(header_row)

    # 準備資料
    for index, row_data in enumerate(results):
        transaction_date = row_data.get("交易日期", "")
        formatted_date_for_I_str = transaction_date.replace("-", "") if transaction_date else ""
        formatted_date_for_I_val = formatted_date_for_I_str if formatted_date_for_I_str else None

        format_code = row_data.get("格式", 25)
        try: format_code_int = int(format_code)
        except (ValueError, TypeError): format_code_int = 25
        format_code_val = str(format_code_int) if format_code_int is not None else None

        try: tax_exclusive = int(row_data.get("未稅金額", 0))
        except (ValueError, TypeError): tax_exclusive = 0
        try: tax = int(row_data.get("進項稅額", 0))
        except (ValueError, TypeError): tax = 0
        try: total = int(row_data.get("金額總計", 0))
        except (ValueError, TypeError): total = 0

        gv_dict = {
            "序號": index + 1,
            "公司別": "HD",
            "發票號碼": str(row_data.get("統一發票號碼", "")),
            "稅籍編號": "721401318",
            "統一編號": "03251000",
            "記帳點": 1,
            "發票/憑證類別": str(convert_format_code_to_type(format_code_int)),
            "格式代號": format_code_val,
            "單據憑證日期": formatted_date_for_I_val,
            "傳票日期": None,
            "申報年月": None,
            "銷售人統一編號": str(row_data.get("賣方統一編號", "")),
            "銷售人名稱": str(row_data.get("賣方名稱", "")),
            "課稅別": "1",
            "進貨折讓區分": None,
            "未稅金額": tax_exclusive if tax_exclusive != 0 else None,
            "進項稅額": tax if tax != 0 else None,
            "金額總計": total if total != 0 else None,
            "進項稅性質別": "126200",
            "扣抵代號": 1,
            "彙總張數": 0,
            "彙加註記": "N",
            "應付立帳號碼": str(account_payable_code) if account_payable_code else None,
            "備註": None
        }
        
        gv_dict["序號"] = str(gv_dict["序號"]) if gv_dict["序號"] is not None else None
        current_row_list = [gv_dict.get(header_name) for header_name in header_row]
        data_for_excel.append(current_row_list)

    # --- 建立 .xls 檔案 (使用 xlwt) ---
    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet("PURDATA")

    # 設定字體與樣式
    font_content = xlwt.Font()
    font_content.name = 'Microsoft JhengHei'
    font_content.height = 280 # 14pt (20 * 14)

    # 置中樣式 (標題)
    style_center = xlwt.XFStyle()
    alignment_center = xlwt.Alignment()
    alignment_center.horz = xlwt.Alignment.HORZ_CENTER
    alignment_center.vert = xlwt.Alignment.VERT_CENTER
    style_center.alignment = alignment_center
    style_center.font = font_content

    # 靠左樣式 (文字)
    style_left = xlwt.XFStyle()
    alignment_left = xlwt.Alignment()
    alignment_left.horz = xlwt.Alignment.HORZ_LEFT
    alignment_left.vert = xlwt.Alignment.VERT_CENTER
    style_left.alignment = alignment_left
    style_left.font = font_content

    # 靠右樣式 (數字)
    style_right = xlwt.XFStyle()
    alignment_right = xlwt.Alignment()
    alignment_right.horz = xlwt.Alignment.HORZ_RIGHT
    alignment_right.vert = xlwt.Alignment.VERT_CENTER
    style_right.alignment = alignment_right
    style_right.font = font_content

    # 寫入資料
    for row_idx, row_data in enumerate(data_for_excel):
        for col_idx, cell_value in enumerate(row_data):
            # 判斷樣式
            current_style = style_left
            if row_idx == 0: # 表頭
                current_style = style_center
            elif col_idx in [0, 7]: # 序號, 格式代號
                current_style = style_left
            elif isinstance(cell_value, (int, float)): # 數字
                current_style = style_right
            
            # 寫入儲存格
            val_to_write = cell_value if cell_value is not None else ''
            ws.write(row_idx, col_idx, val_to_write, current_style)

    # 自動調整欄寬 (簡易版估算)
    for col_idx in range(len(header_row)):
        max_len = 0
        for row_data in data_for_excel:
            val = str(row_data[col_idx]) if row_data[col_idx] is not None else ""
            # 計算長度 (中文字算2)
            curr_len = 0
            for char in val: curr_len += 2 if '\u4e00' <= char <= '\u9fff' else 1
            if curr_len > max_len: max_len = curr_len
        
        # xlwt 寬度單位是 1/256 個字元寬度
        ws.col(col_idx).width = 256 * (max_len + 2)

    output_buffer = io.BytesIO()
    wb.save(output_buffer)
    excel_data = output_buffer.getvalue()

    # 回傳 .xls MIME type
    return Response(excel_data, mimetype="application/vnd.ms-excel", headers={"Content-Disposition": "attachment;filename=GV_output.xls"})

# --- 新增功能：產生費用報支檔 (含彙加邏輯) ---
@app.route('/generate_expense_report', methods=['POST'])
def generate_expense_report():
    # --- 修改說明: 使用 xlwt 產生 .xls 檔案 (Excel 97-2003) ---
    json_data = request.json
    results = json_data.get('results', [])
    if not results: return jsonify({"error": "沒有資料可供下載"}), 400

    def convert_format_code_to_type(code): return "Q" if str(code) == "22" else "I"

    # 1. 依照格式代碼分組 (邏輯維持不變)
    groups = {} 
    for row in results:
        fmt_str = str(row.get("格式", 25))
        try: fmt = int(fmt_str)
        except: fmt = 25
        if fmt not in groups: groups[fmt] = []
        groups[fmt].append(row)

    processed_rows = []

    # 2. 處理每一組格式代碼 (邏輯維持不變)
    for fmt, items in groups.items():
        small_tax_items = [] 
        large_tax_items = [] 

        for item in items:
            try: tax = int(item.get("進項稅額", 0))
            except: tax = 0
            if tax < 500: small_tax_items.append(item)
            else: large_tax_items.append(item)

        for item in large_tax_items:
            item['_is_aggregated'] = "N"; item['_agg_count'] = 0; item['_final_fmt'] = fmt
            processed_rows.append(item)

        if small_tax_items:
            sum_tax_exclusive = 0; sum_tax = 0; sum_total = 0
            small_tax_items.sort(key=lambda x: (
                int(x.get("進項稅額", 0) if x.get("進項稅額") is not None else 0),
                int(x.get("未稅金額", 0) if x.get("未稅金額") is not None else 0),
                x.get("統一發票號碼", "")
            ), reverse=True)
            representative = small_tax_items[0].copy() 

            for s_item in small_tax_items:
                try: v1 = int(s_item.get("未稅金額", 0)); sum_tax_exclusive += v1
                except: pass
                try: v2 = int(s_item.get("進項稅額", 0)); sum_tax += v2
                except: pass
                try: v3 = int(s_item.get("金額總計", 0)); sum_total += v3
                except: pass
            
            representative["未稅金額"] = sum_tax_exclusive
            representative["進項稅額"] = sum_tax
            representative["金額總計"] = sum_total
            representative['_is_aggregated'] = "Y"
            representative['_agg_count'] = len(small_tax_items)
            
            if fmt == 21: representative['_final_fmt'] = 26
            elif fmt == 22: representative['_final_fmt'] = 27
            else: representative['_final_fmt'] = fmt 

            processed_rows.append(representative)

    # 3. 準備 Excel 輸出 (.xls)
    headers = [
        "公司別", "憑證類別", "格式代號", "發票號碼", "記帳點", 
        "統一編號", "單據憑證日期", "銷售人統一編號", "銷售人名稱", 
        "未稅金額", "進項稅金額", "金額總計", "彙加註記", "彙總張數", "進項稅性質別"
    ]

    wb = xlwt.Workbook(encoding='utf-8')
    ws = wb.add_sheet("ExpenseReport")

    # 設定字體 (12pt)
    font_content = xlwt.Font()
    font_content.name = 'Microsoft JhengHei'
    font_content.height = 240 # 12pt (20 * 12)

    # 靠左樣式
    style_left = xlwt.XFStyle()
    alignment_left = xlwt.Alignment()
    alignment_left.horz = xlwt.Alignment.HORZ_LEFT
    alignment_left.vert = xlwt.Alignment.VERT_CENTER
    style_left.alignment = alignment_left
    style_left.font = font_content

    # 靠右樣式
    style_right = xlwt.XFStyle()
    alignment_right = xlwt.Alignment()
    alignment_right.horz = xlwt.Alignment.HORZ_RIGHT
    alignment_right.vert = xlwt.Alignment.VERT_CENTER
    style_right.alignment = alignment_right
    style_right.font = font_content

    # 寫入表頭
    for col_idx, h in enumerate(headers):
        ws.write(0, col_idx, h, style_left)

    # 寫入資料
    for row_idx, item in enumerate(processed_rows, 1):
        transaction_date = item.get("交易日期", "")
        formatted_date = transaction_date.replace("-", "") if transaction_date else ""
        final_fmt = item.get('_final_fmt', 25)

        row_values = [
            "HD", 
            convert_format_code_to_type(final_fmt), 
            final_fmt, 
            item.get("統一發票號碼", ""), 
            1, 
            "03251000", 
            formatted_date, 
            item.get("賣方統一編號", ""), 
            item.get("賣方名稱", ""), 
            item.get("未稅金額", 0), 
            item.get("進項稅額", 0), 
            item.get("金額總計", 0), 
            item.get("_is_aggregated", "N"), 
            item.get("_agg_count", 0), 
            "126200" 
        ]

        for col_idx, val in enumerate(row_values):
            current_style = style_left
            if headers[col_idx] in ["未稅金額", "進項稅金額", "金額總計", "彙總張數"]:
                current_style = style_right
                try: val = int(val)
                except: pass
            else:
                val = str(val) # 強制轉字串

            ws.write(row_idx, col_idx, val, current_style)

    # 自動調整欄寬
    for col_idx in range(len(headers)):
        max_len = len(str(headers[col_idx])) * 2 # 表頭長度
        for r_idx, item in enumerate(processed_rows):
            # 簡易估算，這裡不需要太精準，只要夠寬即可
            pass 
        # 直接設定一個比較寬的預設值，xlwt 的自動調整比較麻煩
        ws.col(col_idx).width = 256 * 15 # 預設寬度

    output_buffer = io.BytesIO()
    wb.save(output_buffer)
    excel_data = output_buffer.getvalue()

    return Response(excel_data, mimetype="application/vnd.ms-excel", headers={"Content-Disposition": "attachment;filename=Expense_Report.xls"})

if __name__ == '__main__':
    print("--- 發票批次辨識與剖析程式 (v52.0 - 相容舊版 Excel .xls) ---")
    app.run(port=5000, debug=True)