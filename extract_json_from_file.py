import re

def clean_output_file(input_path, output_path):
    with open(input_path, "r", encoding="utf-8") as fin, open(output_path, "w", encoding="utf-8") as fout:
        content = fin.read()
        # Tìm tất cả các block Assistant: ... (JSON)
        matches = re.findall(r'Assistant:\s*((\[.*?\]|\{.*?\}))', content, re.DOTALL)
        for match in matches:
            json_str = match[0].strip()
            fout.write(json_str + "\n\n")  # mỗi block cách nhau 1 dòng trống

if __name__ == "__main__":
    clean_output_file("output.txt", "output1.txt")
    print("Đã làm sạch file output.txt, chỉ còn lại tất cả các block JSON như mẫu text.txt.") 