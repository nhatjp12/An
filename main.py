import os
import subprocess
import sys
import numpy as np
import pandas as pd
from fastapi import FastAPI, File, UploadFile, Request
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.templating import Jinja2Templates
from typing import List
import shutil
from finals import load_image, model, tokenizer
import json
from datetime import datetime

# Thêm JSON encoder để xử lý các kiểu dữ liệu từ numpy
class NumpyEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, (np.integer, np.int64)):
            return int(obj)
        elif isinstance(obj, (np.floating, np.float64)):
            return float(obj)
        elif isinstance(obj, np.ndarray):
            return obj.tolist()
        return super(NumpyEncoder, self).default(obj)

generation_config = dict(
    max_new_tokens=4096,
    do_sample=False,
    num_beams=3,
    repetition_penalty=3.5
)

prompt = '''<image>\nNhận diện hoá đơn trong ảnh. Chỉ trả về phần liệt kê từng mặt hàng hàng dưới dạng JSON, phần Ngày tạo đơn và Tên khách hàng chỉ in ra một lần:
[
  {
    "Ngày tạo đơn": "Ngày tạo đơn",
    "Tên khách hàng": "Tên khách hàng",
  }
  {
    "Tên mặt hàng": "Tên mặt hàng",
    "Đơn vị tính": "Đơn vị tính",
    "Số lượng": "Số lượng",
    "Đơn giá": "Đơn giá",
    "Thành tiền": "Thành tiền"
  },
]
'''
question = prompt

app = FastAPI()

UPLOAD_DIR = 'uploaded_images'
os.makedirs(UPLOAD_DIR, exist_ok=True)

app.mount("/static", StaticFiles(directory="static"), name="static")
templates = Jinja2Templates(directory="static")

def extract_data_to_excel():
    # Kiểm tra nếu file text.txt có nội dung
    if os.path.exists("text.txt") and os.path.getsize("text.txt") > 0:
        try:
            # Chạy file extract_to_excel.py để xuất dữ liệu ra Excel
            subprocess.run([sys.executable, "extract_to_excel.py"], check=True)
            return True
        except Exception as e:
            print(f"Lỗi khi chạy extract_to_excel.py: {e}")
            return False
    return False

@app.get("/", response_class=HTMLResponse)
def main(request: Request):
    return templates.TemplateResponse("index.html", {"request": request})

@app.get("/excel-data/")
async def get_excel_data():
    """Đọc dữ liệu từ file Excel và trả về dạng JSON."""
    try:
        # Kiểm tra xem file Excel có tồn tại không
        if not os.path.exists("output.xlsx"):
            return JSONResponse(content={"success": False, "message": "File Excel chưa được tạo"})
        
        # Đọc file Excel
        df = pd.read_excel("output.xlsx")
        
        # Chuyển đổi DataFrame thành JSON
        data = df.fillna("").to_dict(orient="records")
        
        # Chuyển đổi thủ công bằng json.dumps trước
        json_data = json.dumps({"success": True, "data": data}, cls=NumpyEncoder)
        return JSONResponse(content=json.loads(json_data))
    except Exception as e:
        return JSONResponse(content={"success": False, "message": f"Lỗi khi đọc file Excel: {str(e)}"})

@app.post("/process_images/")
async def process_images(files: List[UploadFile] = File(...)):
    results = []
    with open("text.txt", "a", encoding="utf-8") as f:
        for file in files:
            file_path = os.path.join(UPLOAD_DIR, file.filename)
            with open(file_path, "wb") as buffer:
                shutil.copyfileobj(file.file, buffer)
            try:
                pixel_values = load_image(file_path, max_num=6)
                response = model.chat(tokenizer, pixel_values, question, generation_config)
                results.append({
                    "filename": file.filename,
                    "result": response
                })
                f.write(f'{response}\n\n')
            except Exception as e:
                results.append({
                    "filename": file.filename,
                    "error": str(e)
                })
                f.write(f'ERROR: {str(e)}\n\n')
    
    # Sau khi xử lý xong các ảnh, chạy xuất Excel
    excel_generated = extract_data_to_excel()
    if excel_generated:
        for result in results:
            if "result" in result:
                result["excel_status"] = "Dữ liệu đã được xuất ra file Excel"
    
    return JSONResponse(content={"results": results})

@app.get("/dashboard-data/")
async def get_dashboard_data():
    """Phân tích dữ liệu Excel và tạo các thống kê cho Dashboard."""
    try:
        if not os.path.exists("output.xlsx"):
            return JSONResponse(content={"success": False, "message": "File Excel chưa được tạo"})
        
        # Đọc file Excel
        df = pd.read_excel("output.xlsx")
        
        # Đảm bảo các cột số liệu là dạng số
        for col in ['Số lượng', 'Đơn giá', 'Thành tiền']:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        
        # 1. Tổng doanh thu theo khách hàng
        revenue_by_customer = df.groupby('Tên khách hàng')['Thành tiền'].sum().reset_index()
        revenue_by_customer = revenue_by_customer.sort_values('Thành tiền', ascending=False)
        # Chuyển numpy types sang Python native types
        revenue_by_customer_list = [
            {"Tên khách hàng": row["Tên khách hàng"], "Thành tiền": float(row["Thành tiền"])}
            for _, row in revenue_by_customer.iterrows()
        ]
        
        # 2. Tổng doanh thu theo sản phẩm
        revenue_by_product = df.groupby('Tên mặt hàng')['Thành tiền'].sum().reset_index()
        revenue_by_product = revenue_by_product.sort_values('Thành tiền', ascending=False)
        revenue_by_product_list = [
            {"Tên mặt hàng": row["Tên mặt hàng"], "Thành tiền": float(row["Thành tiền"])}
            for _, row in revenue_by_product.iterrows()
        ]
        
        # 3. Tổng doanh thu theo tháng
        # Chuyển đổi cột ngày về định dạng datetime
        df['Ngày'] = pd.to_datetime(df['Ngày tạo đơn'], format='%d/%m/%Y', errors='coerce')
        df['Tháng'] = df['Ngày'].dt.strftime('%m/%Y')
        revenue_by_month = df.groupby('Tháng')['Thành tiền'].sum().reset_index()
        revenue_by_month_list = [
            {"Tháng": row["Tháng"], "Thành tiền": float(row["Thành tiền"])}
            for _, row in revenue_by_month.iterrows()
        ]
        
        # 4. Số lượng bán theo sản phẩm
        quantity_by_product = df.groupby('Tên mặt hàng')['Số lượng'].sum().reset_index()
        quantity_by_product = quantity_by_product.sort_values('Số lượng', ascending=False)
        quantity_by_product_list = [
            {"Tên mặt hàng": row["Tên mặt hàng"], "Số lượng": int(row["Số lượng"])}
            for _, row in quantity_by_product.iterrows()
        ]
        
        # 5. Thống kê đơn hàng theo khách hàng
        orders_by_customer = df.groupby('Tên khách hàng')['Mã tạo đơn'].nunique().reset_index()
        orders_by_customer.columns = ['Tên khách hàng', 'Số đơn hàng']
        orders_by_customer = orders_by_customer.sort_values('Số đơn hàng', ascending=False)
        orders_by_customer_list = [
            {"Tên khách hàng": row["Tên khách hàng"], "Số đơn hàng": int(row["Số đơn hàng"])}
            for _, row in orders_by_customer.iterrows()
        ]
        
        # 6. Giá trị đơn hàng trung bình
        df_orders = df.groupby(['Mã tạo đơn', 'Ngày tạo đơn', 'Tên khách hàng'])['Thành tiền'].sum().reset_index()
        avg_order_value = float(df_orders['Thành tiền'].mean())
        
        # Tổng hợp dữ liệu thống kê
        stats = {
            "total_revenue": float(df['Thành tiền'].sum()),
            "total_orders": int(df['Mã tạo đơn'].nunique()),
            "total_customers": int(df['Tên khách hàng'].nunique()),
            "total_products": int(df['Tên mặt hàng'].nunique()),
            "avg_order_value": avg_order_value,
            "revenue_by_customer": revenue_by_customer_list,
            "revenue_by_product": revenue_by_product_list,
            "revenue_by_month": revenue_by_month_list,
            "quantity_by_product": quantity_by_product_list,
            "orders_by_customer": orders_by_customer_list
        }
        
        # Chuyển đổi thủ công bằng json.dumps trước
        json_data = json.dumps({"success": True, "data": stats}, cls=NumpyEncoder)
        return JSONResponse(content=json.loads(json_data))
    except Exception as e:
        import traceback
        print(traceback.format_exc())
        return JSONResponse(content={"success": False, "message": f"Lỗi khi phân tích dữ liệu: {str(e)}"}) 