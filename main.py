from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
import sqlite3
import pandas as pd
import openpyxl
from datetime import datetime
import io
import os

app = FastAPI(title="ELWAHEB Massive Sourcing & Excel Exporter")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

DB_PATH = "sourcing.db"

def init_db():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    # جدول قاعدة البيانات العملاقة لمنتجات المستوردين والكميات
    c.execute('''CREATE TABLE IF NOT EXISTS catalog (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        product_name TEXT,
        carton_qty INTEGER DEFAULT 100,
        unit_cost REAL,
        importer_name TEXT,
        pickup_location TEXT,
        invoice_required TEXT DEFAULT 'نعم',
        barcode TEXT,
        status TEXT DEFAULT 'PENDING',
        asin TEXT,
        amazon_price REAL,
        total_amazon_fees REAL,
        net_profit_unit REAL,
        monthly_sales_range TEXT,
        is_golden INTEGER DEFAULT 0,
        updated_at TEXT
    )''')
    conn.commit()
    conn.close()

init_db()

class CatalogItem(BaseModel):
    product_name: str
    carton_qty: int = 100
    unit_cost: float
    importer_name: str
    pickup_location: str = "القاهرة"
    invoice_required: str = "نعم"
    barcode: str = ""

class ScanEvaluation(BaseModel):
    catalog_id: int
    asin: str
    amazon_title: str
    amazon_price: float
    total_amazon_fees: float
    monthly_sales_count: int
    is_pullable: bool

# 1. رفع ملف إكسيل ضخم من أي مستورد لإضافته لقاعدة البيانات العملاقة
@app.post("/upload-supplier-excel")
async def upload_supplier_excel(
    importer_name: str,
    pickup_location: str = "ابو النمرس",
    invoice_required: str = "نعم",
    file: UploadFile = File(...)
):
    contents = await file.read()
    df = pd.read_excel(io.BytesIO(contents))
    
    # توحيد أسماء الأعمدة تلقائياً مهما كان تنسيق ملف المستورد
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    added_count = 0
    for _, row in df.iterrows():
        name = str(row.get("المنتج") or row.get("الصنف") or row.get("name") or "").strip()
        cost = float(row.get("سعر الوحدة") or row.get("السعر") or row.get("price") or 0)
        qty = int(row.get("الكمية") or row.get("كمية الكرتونة") or row.get("qty") or 100)
        barcode = str(row.get("الباركود") or row.get("barcode") or "").strip()
        
        if name and cost > 0:
            c.execute('''INSERT INTO catalog 
                (product_name, carton_qty, unit_cost, importer_name, pickup_location, invoice_required, barcode, updated_at)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?)''',
                (name, qty, cost, importer_name, pickup_location, invoice_required, barcode, datetime.now().isoformat()))
            added_count += 1
            
    conn.commit()
    conn.close()
    return {"status": "success", "imported_products": added_count}

# 2. إضافة منتج فردي لقاعدة البيانات
@app.post("/add-catalog-item")
def add_catalog_item(item: CatalogItem):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''INSERT INTO catalog 
        (product_name, carton_qty, unit_cost, importer_name, pickup_location, invoice_required, barcode, updated_at)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)''',
        (item.product_name, item.carton_qty, item.unit_cost, item.importer_name, item.pickup_location, item.invoice_required, item.barcode, datetime.now().isoformat()))
    conn.commit()
    item_id = c.lastrowid
    conn.close()
    return {"status": "added", "id": item_id}

# 3. سحب المنتج التالي للبحث الآلي في أمازون (Auto-Pilot Queue)
@app.get("/next-pending-product")
def next_pending_product():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT id, product_name, carton_qty, unit_cost, importer_name, barcode FROM catalog WHERE status = 'PENDING' LIMIT 1")
    row = c.fetchone()
    if row:
        c.execute("UPDATE catalog SET status = 'SCANNING' WHERE id = ?", (row[0],))
        conn.commit()
    conn.close()
    if not row:
        return {"done": True, "message": "تم فحص جميع المنتجات في قاعدة البيانات!"}
    return {
        "done": False,
        "id": row[0],
        "product_name": row[1],
        "carton_qty": row[2],
        "unit_cost": row[3],
        "importer_name": row[4],
        "barcode": row[5]
    }

# 4. تقييم نتيجة الفحص الآلي وحفظ الفرص الذهبية
@app.post("/submit-scan-result")
def submit_scan_result(res: ScanEvaluation):
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute("SELECT unit_cost, carton_qty FROM catalog WHERE id = ?", (res.catalog_id,))
    row = c.fetchone()
    if not row:
        raise HTTPException(status_code=404, detail="Item not found")
        
    unit_cost, carton_qty = row
    packaging_shipping = 5.0
    net_profit = res.amazon_price - (unit_cost + res.total_amazon_fees + packaging_shipping)
    
    # شروط الفلترة الصارمة: ربح >= 100% + مبيعات >= 100 + متاح للسحب
    is_golden = 1 if (net_profit >= unit_cost and res.monthly_sales_count >= 100 and res.is_pullable) else 0
    sales_range = f"{res.monthly_sales_count}-{res.monthly_sales_count + 100}"
    
    c.execute('''UPDATE catalog SET 
        status = 'DONE', asin = ?, amazon_price = ?, total_amazon_fees = ?,
        net_profit_unit = ?, monthly_sales_range = ?, is_golden = ?, updated_at = ?
        WHERE id = ?''',
        (res.asin, res.amazon_price, res.total_amazon_fees, round(net_profit, 2),
         sales_range, is_golden, datetime.now().isoformat(), res.catalog_id))
    conn.commit()
    conn.close()
    return {"is_golden": bool(is_golden), "net_profit": round(net_profit, 2)}

# 5. تصدير الفرص الرابحة إلى ملف إكسيل مطابق تماماً لملف "شحنة رقم ١" (الأعمدة A إلى K)
@app.get("/export-golden-excel")
def export_golden_excel():
    conn = sqlite3.connect(DB_PATH)
    c = conn.cursor()
    c.execute('''SELECT product_name, carton_qty, monthly_sales_range, unit_cost, net_profit_unit,
                 asin, importer_name, pickup_location, invoice_required
                 FROM catalog WHERE is_golden = 1 ORDER BY net_profit_unit DESC''')
    rows = c.fetchall()
    conn.close()

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.sheet_view.rightToLeft = True

    # نفس العناوين الـ 11 الموجودة في ملفك بالضبط (A1 إلى K1)
    headers = [
        "المنتج", "الكمية", "المتوقع البيع", "سعر الوحدة", "هامش ربح القطعة",
        "ASIN", "المورد", "موقع الاستلام", "الفاتورة مطلوبة؟", "هامش الربح المقدر", "تكلفة المنتج"
    ]
    ws.append(headers)

    for idx, r in enumerate(rows, start=2):
        prod_name, qty, sales_rng, unit_cost, unit_profit, asin, importer, location, invoice = r
        # معادلات إكسيل ديناميكية للأعمدة J و K تماماً كما في جدولك
        est_total_profit = f"=B{idx}*E{idx}"  # الكمية × هامش ربح القطعة
        total_cost = f"=B{idx}*D{idx}"        # الكمية × سعر الوحدة
        ws.append([prod_name, qty, sales_rng, unit_cost, unit_profit, asin, importer, location, invoice, est_total_profit, total_cost])

    output_file = "ELWAHEB_Golden_Shipment.xlsx"
    wb.save(output_file)
    return FileResponse(output_file, filename=output_file, media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")