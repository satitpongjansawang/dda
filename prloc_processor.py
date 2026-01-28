"""
PR Local Purchase Processor
ค้นหาเอกสารที่อนุมัติแล้วจาก SQL Server และเรียก Google Apps Script เพื่อแปลงเป็น CSV
"""

import pyodbc
import requests
import logging
from dataclasses import dataclass
from typing import Optional

# ===================== CONFIGURATION =====================
@dataclass
class Config:
    # SQL Server Connection
    SQL_SERVER: str = "your_server_name"
    SQL_DATABASE: str = "DASY-FLII"
    SQL_USERNAME: str = "your_username"
    SQL_PASSWORD: str = "your_password"
    SQL_DRIVER: str = "{ODBC Driver 17 for SQL Server}"

    # Google Apps Script Web App URL
    # ต้อง Deploy Google Apps Script เป็น Web App ก่อน
    GAS_WEB_APP_URL: str = "https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec"


CONFIG = Config()

# ===================== LOGGING =====================
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)
logger = logging.getLogger(__name__)


# ===================== DATABASE FUNCTIONS =====================
def get_db_connection() -> pyodbc.Connection:
    """สร้าง connection ไปยัง SQL Server"""
    connection_string = (
        f"DRIVER={CONFIG.SQL_DRIVER};"
        f"SERVER={CONFIG.SQL_SERVER};"
        f"DATABASE={CONFIG.SQL_DATABASE};"
        f"UID={CONFIG.SQL_USERNAME};"
        f"PWD={CONFIG.SQL_PASSWORD};"
        "TrustServerCertificate=yes;"
    )
    return pyodbc.connect(connection_string)


def get_approved_sinsei_code() -> Optional[str]:
    """
    ค้นหา SINSEI_CODE ของเอกสารที่อนุมัติแล้ว (JOUTAI_KBN = 1)

    Returns:
        SINSEI_CODE หรือ None ถ้าไม่พบ
    """
    query = """
        SELECT TOP 1 SINSEI_CODE
        FROM FLIISA.TR_SINSEI_DATA_HEADER
        WHERE AUTO_NO_CHR = 'PRLOC'
          AND JOUTAI_KBN = 1
          AND DATEPART(MILLISECOND, INS_DATE) != 0
          AND DATEPART(MILLISECOND, UPD_DATE) != 0
          AND DATEPART(SECOND, INS_DATE) != 0
          AND DATEPART(SECOND, UPD_DATE) != 0
        ORDER BY UPD_DATE ASC
    """

    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query)
            row = cursor.fetchone()

            if row:
                sinsei_code = row[0]
                logger.info(f"พบ SINSEI_CODE: {sinsei_code}")
                return sinsei_code
            else:
                logger.warning("ไม่พบเอกสารที่อนุมัติแล้ว")
                return None

    except pyodbc.Error as e:
        logger.error(f"Database error: {e}")
        raise


def get_google_drive_link(sinsei_code: str) -> Optional[str]:
    """
    ค้นหา Google Drive link จาก SINSEI_CODE

    Args:
        sinsei_code: รหัสเอกสาร

    Returns:
        Google Drive link หรือ None ถ้าไม่พบ
    """
    query = """
        SELECT KOUMOKU_VALUE
        FROM [DASY-FLII].[FLIISA].[TR_SINSEI_DATA_LNK]
        WHERE SINSEI_CODE = ?
          AND KOUMOKU_KEY = 'LNK1'
    """

    try:
        with get_db_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(query, (sinsei_code,))
            row = cursor.fetchone()

            if row:
                drive_link = row[0]
                logger.info(f"พบ Google Drive link: {drive_link}")
                return drive_link
            else:
                logger.warning(f"ไม่พบ link สำหรับ SINSEI_CODE: {sinsei_code}")
                return None

    except pyodbc.Error as e:
        logger.error(f"Database error: {e}")
        raise


# ===================== GOOGLE APPS SCRIPT FUNCTIONS =====================
def call_gas_convert_excel(file_link: str) -> dict:
    """
    เรียก Google Apps Script Web App เพื่อแปลง Excel เป็น CSV

    Args:
        file_link: Google Drive link ของไฟล์ Excel

    Returns:
        Response จาก Google Apps Script
    """
    try:
        # ส่ง POST request ไปยัง Web App
        payload = {
            "action": "convertExcel",
            "fileLink": file_link
        }

        response = requests.post(
            CONFIG.GAS_WEB_APP_URL,
            json=payload,
            timeout=120  # 2 minutes timeout
        )

        response.raise_for_status()
        result = response.json()

        if result.get("success"):
            logger.info(f"แปลงไฟล์สำเร็จ: {result.get('data', {}).get('csvName', 'N/A')}")
        else:
            logger.warning(f"แปลงไฟล์ไม่สำเร็จ: {result.get('message', 'Unknown error')}")

        return result

    except requests.exceptions.RequestException as e:
        logger.error(f"Error calling Google Apps Script: {e}")
        return {"success": False, "error": str(e)}


# ===================== MAIN PROCESS =====================
def process_approved_documents() -> dict:
    """
    กระบวนการหลัก: ค้นหาเอกสารที่อนุมัติแล้วและแปลงเป็น CSV

    Returns:
        ผลลัพธ์การประมวลผล
    """
    logger.info("=" * 50)
    logger.info("เริ่มกระบวนการ PR Local Purchase Processor")
    logger.info("=" * 50)

    # Step 1: ค้นหา SINSEI_CODE
    logger.info("Step 1: ค้นหาเอกสารที่อนุมัติแล้ว...")
    sinsei_code = get_approved_sinsei_code()

    if not sinsei_code:
        return {
            "success": False,
            "message": "ไม่พบเอกสารที่อนุมัติแล้ว",
            "step": 1
        }

    # Step 2: ค้นหา Google Drive link
    logger.info("Step 2: ค้นหา Google Drive link...")
    drive_link = get_google_drive_link(sinsei_code)

    if not drive_link:
        return {
            "success": False,
            "message": f"ไม่พบ Google Drive link สำหรับ SINSEI_CODE: {sinsei_code}",
            "step": 2,
            "sinsei_code": sinsei_code
        }

    # Step 3: เรียก Google Apps Script
    logger.info("Step 3: เรียก Google Apps Script เพื่อแปลงไฟล์...")
    gas_result = call_gas_convert_excel(drive_link)

    if gas_result.get("success"):
        logger.info("=" * 50)
        logger.info("กระบวนการเสร็จสิ้น - สำเร็จ")
        logger.info("=" * 50)
        return {
            "success": True,
            "message": "แปลงไฟล์สำเร็จ",
            "sinsei_code": sinsei_code,
            "drive_link": drive_link,
            "csv_result": gas_result.get("data")
        }
    else:
        return {
            "success": False,
            "message": f"แปลงไฟล์ไม่สำเร็จ: {gas_result.get('message', gas_result.get('error', 'Unknown'))}",
            "step": 3,
            "sinsei_code": sinsei_code,
            "drive_link": drive_link
        }


# ===================== ENTRY POINT =====================
if __name__ == "__main__":
    result = process_approved_documents()
    print("\n" + "=" * 50)
    print("ผลลัพธ์:")
    print("=" * 50)
    for key, value in result.items():
        print(f"  {key}: {value}")
