import json
import re
import pandas as pd
import hashlib
from datetime import datetime

def format_date(date_str):
    # Tìm ngày, tháng, năm trong chuỗi
    match = re.search(r'(\d{1,2})\s*tháng\s*(\d{1,2})\s*năm\s*(\d{4})', date_str)
    if match:
        day, month, year = match.groups()
        return f"{int(day):02d}/{int(month):02d}/{year}"
    # Nếu không khớp, trả về chuỗi gốc
    return date_str.strip()

def clean_customer_name(name):
    # Lấy phần trước dấu ( nếu có
    name = name.split('(')[0].strip()
    # Chuẩn hóa tên khách hàng đặc biệt
    if name == 'Thu Bàn':
        return 'Thu Bồn'
    return name

def correct_product_name(name):
    # Sửa lỗi chính tả phổ biến, có thể mở rộng thêm
    corrections = {
        "Nấm bào ngư": ["Nấm Bào Ngư", "Nấm bào ngư", "Nấm bào ngư xám"],
        "Nấm rơm": ["Nấm rơm", "Nấm Rơm"],
        "Nấm đông cô": ["Nấm đông cô", "Nấm Đông Cô"],
        "Nấm mộc nhĩ": ["Nấm mộc nhĩ", "Nấm Mộc Nhĩ"],
        # Thêm các tên khác nếu cần
    }
    for correct, variants in corrections.items():
        if name.strip() in variants:
            return correct
    return name.strip()

def format_number(val):
    # Loại bỏ dấu chấm, phẩy, chuyển về số nguyên nếu có thể
    if isinstance(val, int):
        return val
    val = str(val).replace('.', '').replace(',', '').strip()
    try:
        return int(val)
    except:
        return val

def format_price(val):
    # Loại bỏ dấu chấm, phẩy, chuyển về số nguyên
    num = format_number(val)
    try:
        num = int(num)
        if num < 10000:
            num = num * 1000
        return num
    except:
        return val

def generate_order_code(date, customer):
    # Tạo mã đơn hàng từ ngày và tên khách hàng
    combined = f"{date}-{customer}".lower()
    # Tạo hash ngắn để làm mã đơn hàng
    hash_str = hashlib.md5(combined.encode()).hexdigest()[:8]
    return f"DH-{hash_str.upper()}"

def extract_json_blocks(input_path):
    with open(input_path, "r", encoding="utf-8") as fin:
        content = fin.read()
        # Tìm tất cả các block JSON (dạng list)
        blocks = re.findall(r'(\[.*?\])', content, re.DOTALL)
        all_rows = []
        # Dictionary để lưu mã đơn hàng đã tạo
        order_codes = {}
        
        for block in blocks:
            try:
                data = json.loads(block)
                # Lấy thông tin ngày tạo đơn, tên khách hàng từ phần tử đầu tiên
                ngay_tao = format_date(data[0].get("Ngày tạo đơn", ""))
                ten_kh = clean_customer_name(data[0].get("Tên khách hàng", ""))
                
                # Tạo hoặc lấy mã đơn hàng dựa trên ngày và tên khách hàng
                order_key = f"{ngay_tao}-{ten_kh}"
                if order_key not in order_codes:
                    order_codes[order_key] = generate_order_code(ngay_tao, ten_kh)
                ma_don = order_codes[order_key]
                
                # Các phần tử tiếp theo là mặt hàng
                for item in data[1:]:
                    ten_hang = correct_product_name(item.get("Tên mặt hàng", ""))
                    don_vi = item.get("Đơn vị tính", "")
                    so_luong = format_number(item.get("Số lượng", ""))
                    don_gia = format_price(item.get("Đơn giá", ""))
                    # Thành tiền tính lại
                    try:
                        thanh_tien = int(so_luong) * int(don_gia)
                    except:
                        thanh_tien = ""
                    row = {
                        "Mã tạo đơn": ma_don,
                        "Ngày tạo đơn": ngay_tao,
                        "Tên khách hàng": ten_kh,
                        "Tên mặt hàng": ten_hang,
                        "Đơn vị tính": don_vi,
                        "Số lượng": so_luong,
                        "Đơn giá": don_gia,
                        "Thành tiền": thanh_tien
                    }
                    all_rows.append(row)
            except Exception as e:
                print(f"Lỗi parse block: {e}")
        return all_rows

def save_to_excel(rows, output_path):
    if not rows:
        print("Không có dữ liệu để xuất ra Excel")
        return
        
    # Thêm cột STT
    for i, row in enumerate(rows):
        row["STT"] = i + 1
    
    df = pd.DataFrame(rows, columns=[
        "STT", "Mã tạo đơn", "Ngày tạo đơn", "Tên khách hàng", "Tên mặt hàng", 
        "Đơn vị tính", "Số lượng", "Đơn giá", "Thành tiền"
    ])
    df.to_excel(output_path, index=False)

if __name__ == "__main__":
    rows = extract_json_blocks("text.txt")
    save_to_excel(rows, "output.xlsx")
    print("Đã xuất dữ liệu ra file output.xlsx") 