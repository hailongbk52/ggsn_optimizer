import os
import shutil
import sqlite3
import json
from fastapi import FastAPI, UploadFile, File, Form, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import FileResponse
from pydantic import BaseModel
from typing import Dict, Any, List, Optional
import openpyxl

app = FastAPI(title="GGSN Load Calculator API")

# Configure CORS
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

UPLOAD_DIR = os.path.join(os.path.dirname(__file__), "uploads")
os.makedirs(UPLOAD_DIR, exist_ok=True)

DB_FILE = os.path.join(os.path.dirname(__file__), "ggsn_persistent_store.db")

def init_db():
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    
    # Store settings / headers
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS system_settings (
            key TEXT PRIMARY KEY,
            value TEXT
        )
    """)
    
    # Store the 3 sub-tables
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS table_license (
            node TEXT PRIMARY KEY,
            vendor TEXT,
            lic_bear REAL,
            lic_throughput REAL,
            lic_bear_uctt REAL,
            lic_throughput_vhkt REAL
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS table_current (
            node TEXT PRIMARY KEY,
            vendor TEXT,
            bear_su_dung REAL,
            throughput REAL,
            bear_total_su_dung REAL,
            bear_ims REAL,
            ipv4_internet REAL,
            ipv4_ims REAL
        )
    """)
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS table_weight (
            node TEXT PRIMARY KEY,
            vendor TEXT,
            weight REAL,
            new_weight REAL,
            on_off INTEGER
        )
    """)
    
    # Check if empty, insert default initial nodes
    cursor.execute("SELECT COUNT(*) FROM table_license")
    if cursor.fetchone()[0] == 0:
        default_nodes = [
            ("GGPD04", "Huawei", 2500000, 100000, 2750000, 110000),
            ("GGPD05", "Huawei", 2500000, 100000, 2750000, 110000),
            ("GGPD06", "Huawei", 3500000, 120000, 3850000, 130000),
            ("GGPD07", "Huawei", 3500000, 120000, 3850000, 130000),
            ("GGHL04", "ZTE", 1090909, 45454, 1200000, 50000),
            ("GGHL05", "ZTE", 1090909, 45454, 1200000, 50000),
            ("GGHL06", "ZTE", 1090909, 45454, 1200000, 50000),
            ("GGHL07", "ZTE", 1090909, 45454, 1200000, 50000),
            ("GGHL11", "Ericsson", 2500000, 100000, 2750000, 110000),
            ("GGHL12", "Ericsson", 2500000, 100000, 2750000, 110000),
        ]
        cursor.executemany("INSERT INTO table_license VALUES (?, ?, ?, ?, ?, ?)", default_nodes)
        
        default_current = [
            ("GGPD04", "Huawei", 2333754, 88777.52, 2333754, 1166877, 2000000, 350000),
            ("GGPD05", "Huawei", 2306341, 88791.13, 2306341, 1153170, 2000000, 340000),
            ("GGPD06", "Huawei", 3075825, 116824.07, 3075825, 1537912, 3000000, 480000),
            ("GGPD07", "Huawei", 3048884, 116981.69, 3048884, 1524442, 3000000, 470000),
            ("GGHL04", "ZTE", 1030469, 40625.70, 1030469, 515234, 1000000, 200000),
            ("GGHL05", "ZTE", 1037096, 40537.58, 1037096, 518548, 1000000, 180000),
            ("GGHL06", "ZTE", 299166, 3101.05, 299166, 149583, 1000000, 190000),
            ("GGHL07", "ZTE", 1118598, 42880.41, 1118598, 559299, 1000000, 200000),
            ("GGHL11", "Ericsson", 2490029, 89261.57, 2490029, 1245014, 2000000, 360000),
            ("GGHL12", "Ericsson", 2474712, 87932.89, 2474712, 1237356, 2000000, 340000),
        ]
        cursor.executemany("INSERT INTO table_current VALUES (?, ?, ?, ?, ?, ?, ?, ?)", default_current)
        
        default_weight = [
            ("GGPD04", "Huawei", 90, 90, 1),
            ("GGPD05", "Huawei", 90, 90, 1),
            ("GGPD06", "Huawei", 120, 120, 1),
            ("GGPD07", "Huawei", 120, 120, 1),
            ("GGHL04", "ZTE", 60, 60, 1),
            ("GGHL05", "ZTE", 60, 60, 1),
            ("GGHL06", "ZTE", 60, 60, 1),
            ("GGHL07", "ZTE", 60, 60, 1),
            ("GGHL11", "Ericsson", 90, 90, 1),
            ("GGHL12", "Ericsson", 90, 90, 1),
        ]
        cursor.executemany("INSERT INTO table_weight VALUES (?, ?, ?, ?, ?)", default_weight)
        
        conn.commit()
    conn.close()

init_db()

class DBQueryRequest(BaseModel):
    connection_string: Optional[str] = None
    sql_query: str

class ExportRequest(BaseModel):
    file_name: str
    grid_data: List[Dict[str, Any]]

class SaveDataRequest(BaseModel):
    table_key: str
    rows: List[Dict[str, Any]]

@app.get("/api/get-all-data")
async def get_all_data():
    conn = sqlite3.connect(DB_FILE)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()
    
    # Load License
    cursor.execute("SELECT * FROM table_license")
    license_rows = [dict(r) for r in cursor.fetchall()]
    
    # Load Current
    cursor.execute("SELECT * FROM table_current")
    current_rows = [dict(r) for r in cursor.fetchall()]
    
    # Load Weight
    cursor.execute("SELECT * FROM table_weight")
    weight_rows = [dict(r) for r in cursor.fetchall()]
    
    # Translate DB keys back to UI template columns
    # UI License cols: ["Node", "Vendor", "License Bear", "License Throughput", "License Bear UCTT (110%)", "License Throughput VHKT"]
    license_ui = []
    for r in license_rows:
        license_ui.append({
            "Node": r["node"],
            "Vendor": r["vendor"] or "Huawei",
            "License Bear": r["lic_bear"],
            "License Throughput": r["lic_throughput"],
            "License Bear UCTT (110%)": r["lic_bear_uctt"],
            "License Throughput VHKT": r["lic_throughput_vhkt"]
        })
        
    # UI Current cols: ["Node", "Vendor", "Bear sử dụng", "Throughput", "Bear Total sử dụng", "Bear IMS", "Total IPv4 (v-internet)", "Total IPv4 (IMS)"]
    current_ui = []
    for r in current_rows:
        current_ui.append({
            "Node": r["node"],
            "Vendor": r["vendor"] or "Huawei",
            "Bear sử dụng": r["bear_su_dung"],
            "Throughput": r["throughput"],
            "Bear Total sử dụng": r["bear_total_su_dung"],
            "Bear IMS": r["bear_ims"],
            "Total IPv4 (v-internet)": r["ipv4_internet"],
            "Total IPv4 (IMS)": r["ipv4_ims"]
        })
        
    # UI Weight cols: ["Node", "Vendor", "Weight (AH)", "NEW WEIGHT (AK)", "ON/OFF"]
    weight_ui = []
    for r in weight_rows:
        weight_ui.append({
            "Node": r["node"],
            "Vendor": r["vendor"] or "Huawei",
            "Weight (AH)": r["weight"],
            "NEW WEIGHT (AK)": r["new_weight"],
            "ON/OFF": r["on_off"]
        })
        
    conn.close()
    return {
        "success": True,
        "license": license_ui,
        "current": current_ui,
        "weight": weight_ui
    }

@app.post("/api/save-table-data")
async def save_table_data(req: SaveDataRequest):
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    
    try:
        if req.table_key == "license":
            cursor.execute("DELETE FROM table_license")
            for r in req.rows:
                node = str(r.get("Node") or r.get("node") or "")
                if not node: continue
                cursor.execute("""
                    INSERT OR REPLACE INTO table_license (node, vendor, lic_bear, lic_throughput, lic_bear_uctt, lic_throughput_vhkt)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (
                    node,
                    r.get("Vendor") or r.get("vendor") or "Huawei",
                    float(r.get("License Bear") or r.get("lic_bear") or 0),
                    float(r.get("License Throughput") or r.get("lic_throughput") or 0),
                    float(r.get("License Bear UCTT (110%)") or r.get("lic_bear_uctt") or 0),
                    float(r.get("License Throughput VHKT") or r.get("lic_throughput_vhkt") or 0)
                ))
        elif req.table_key == "current":
            cursor.execute("DELETE FROM table_current")
            for r in req.rows:
                node = str(r.get("Node") or r.get("node") or "")
                if not node: continue
                cursor.execute("""
                    INSERT OR REPLACE INTO table_current (node, vendor, bear_su_dung, throughput, bear_total_su_dung, bear_ims, ipv4_internet, ipv4_ims)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                """, (
                    node,
                    r.get("Vendor") or r.get("vendor") or "Huawei",
                    float(r.get("Bear sử dụng") or r.get("bear_su_dung") or 0),
                    float(r.get("Throughput") or r.get("throughput") or 0),
                    float(r.get("Bear Total sử dụng") or r.get("bear_total_su_dung") or 0),
                    float(r.get("Bear IMS") or r.get("bear_ims") or 0),
                    float(r.get("Total IPv4 (v-internet)") or r.get("ipv4_internet") or 0),
                    float(r.get("Total IPv4 (IMS)") or r.get("ipv4_ims") or 0)
                ))
        elif req.table_key == "weight":
            cursor.execute("DELETE FROM table_weight")
            for r in req.rows:
                node = str(r.get("Node") or r.get("node") or "")
                if not node: continue
                cursor.execute("""
                    INSERT OR REPLACE INTO table_weight (node, vendor, weight, new_weight, on_off)
                    VALUES (?, ?, ?, ?, ?)
                """, (
                    node,
                    r.get("Vendor") or r.get("vendor") or "Huawei",
                    float(r.get("Weight (AH)") or r.get("weight") or 0),
                    float(r.get("NEW WEIGHT (AK)") or r.get("new_weight") or 0),
                    int(r.get("ON/OFF") or r.get("on_off") or 1)
                ))
        
        conn.commit()
        conn.close()
        return {"success": True}
    except Exception as e:
        conn.close()
        raise HTTPException(status_code=500, detail=str(e))

@app.post("/api/query-db")
async def query_db(req: DBQueryRequest):
    try:
        conn = sqlite3.connect(DB_FILE)
        cursor = conn.cursor()
        cursor.execute(req.sql_query)
        columns = [col[0] for col in cursor.description]
        rows = cursor.fetchall()
        
        result = []
        for r in rows:
            result.append(dict(zip(columns, r)))
            
        conn.close()
        return {
            "success": True,
            "columns": columns,
            "rows": result
        }
    except Exception as e:
        return {
            "success": False,
            "detail": f"Database query error: {str(e)}"
        }

@app.post("/api/export")
async def export_file(req: ExportRequest):
    # Search local download directory template first
    src_path = os.path.join("c:\\Users\\longvh3\\Downloads\\Tinh_tai_UCTT", "01072026_New_VN_Report_PS_Core_Resource_Tinhlq1.xlsx")
    if not os.path.exists(src_path):
        raise HTTPException(status_code=404, detail="Original template Excel file not found.")
            
    export_filename = f"exported_{req.file_name}"
    dest_path = os.path.join(UPLOAD_DIR, export_filename)
    shutil.copy(src_path, dest_path)
    
    try:
        wb = openpyxl.load_workbook(dest_path, data_only=False)
        ws = wb["Resource_GGSN_by_NOCPRO"]
        
        # Write modified values/formulas back
        for row in req.grid_data:
            r_idx = row.get("row_num")
            if not r_idx:
                continue
            for col_letter, cell_info in row.items():
                if col_letter in ["row_num", "Node"]:
                    continue
                if isinstance(cell_info, dict):
                    val = cell_info.get("value")
                    formula = cell_info.get("formula")
                    is_formula = cell_info.get("is_formula")
                    
                    c_idx = openpyxl.utils.column_index_from_string(col_letter)
                    
                    if is_formula and formula:
                        ws.cell(row=r_idx, column=c_idx, value=formula)
                    else:
                        try:
                            num_val = float(val)
                            ws.cell(row=r_idx, column=c_idx, value=num_val)
                        except (ValueError, TypeError):
                            ws.cell(row=r_idx, column=c_idx, value=val)
                            
        wb.save(dest_path)
        return FileResponse(
            dest_path, 
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=export_filename
        )
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to generate Excel export: {str(e)}")

if __name__ == "__main__":
    import uvicorn
    uvicorn.run("main:app", host="127.0.0.1", port=8000, reload=True)
