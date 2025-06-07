import tkinter as tk
from tkinter import filedialog, messagebox
import base64
import datetime
import requests
import hmac
import hashlib
from openpyxl import Workbook
from requests.auth import AuthBase

REGION = "ap-southeast-1"
ENDPOINT = f"https://ocr.cn-north-4.myhuaweicloud.com/v2/{REGION}/ocr/smart-document-recognizer"

class HuaweiSigner(AuthBase):
    def __init__(self, ak, sk):
        self.ak = ak
        self.sk = sk

    def __call__(self, r):
        timestamp = datetime.datetime.utcnow().strftime('%Y%m%dT%H%M%SZ')
        r.headers['X-Sdk-Date'] = timestamp
        canonical_request = f"{r.method}\n{r.path_url}\n\nhost:{r.url.split('/')[2]}\nx-sdk-date:{timestamp}\n\nhost;x-sdk-date\n{hashlib.sha256(r.body or b'').hexdigest()}"
        string_to_sign = f"HMAC-SHA256\n{timestamp}\n{hashlib.sha256(canonical_request.encode()).hexdigest()}"
        signature = hmac.new(self.sk.encode(), string_to_sign.encode(), hashlib.sha256).hexdigest()
        r.headers['Authorization'] = f"HMAC-SHA256 Credential={self.ak}, SignedHeaders=host;x-sdk-date, Signature={signature}"
        return r

def call_huawei_ocr_api(image_path, ak, sk):
    with open(image_path, "rb") as f:
        image_data = base64.b64encode(f.read()).decode()

    payload = { "data": image_data }
    headers = { "Content-Type": "application/json" }
    auth = HuaweiSigner(ak, sk)
    response = requests.post(ENDPOINT, json=payload, headers=headers, auth=auth)
    if response.status_code == 200:
        return response.json()
    else:
        raise Exception(f"API Error: {response.status_code}\n{response.text}")

def save_tables_to_excel(json_data, output_file):
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "Table"

    table_blocks = json_data.get("result", [{}])[0].get("tables", [])
    if not table_blocks:
        raise Exception("未识别到表格")

    for cell in table_blocks[0].get("table_cells", []):
        row = cell["row"]
        col = cell["column"]
        text = cell["text"]
        sheet.cell(row=row + 1, column=col + 1, value=text)

    workbook.save(output_file)

def run_recognition(image_path, ak, sk, output_path):
    result = call_huawei_ocr_api(image_path, ak, sk)
    save_tables_to_excel(result, output_path)

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("华为OCR表格识别导出")

        tk.Label(root, text="图片文件:").grid(row=0, column=0, sticky="e")
        self.image_entry = tk.Entry(root, width=50)
        self.image_entry.grid(row=0, column=1)
        tk.Button(root, text="浏览", command=self.browse_image).grid(row=0, column=2)

        tk.Label(root, text="Access Key (AK):").grid(row=1, column=0, sticky="e")
        self.ak_entry = tk.Entry(root, width=50)
        self.ak_entry.grid(row=1, column=1, columnspan=2)

        tk.Label(root, text="Secret Key (SK):").grid(row=2, column=0, sticky="e")
        self.sk_entry = tk.Entry(root, width=50, show="*")
        self.sk_entry.grid(row=2, column=1, columnspan=2)

        tk.Label(root, text="输出文件路径:").grid(row=3, column=0, sticky="e")
        self.output_entry = tk.Entry(root, width=50)
        self.output_entry.grid(row=3, column=1)
        tk.Button(root, text="浏览", command=self.browse_output).grid(row=3, column=2)

        tk.Button(root, text="开始识别并导出", command=self.run).grid(row=4, column=1, pady=10)

    def browse_image(self):
        filename = filedialog.askopenfilename(filetypes=[("Images", "*.jpg *.jpeg *.png *.bmp *.tiff *.pdf")])
        if filename:
            self.image_entry.delete(0, tk.END)
            self.image_entry.insert(0, filename)

    def browse_output(self):
        filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel 文件", "*.xlsx")])
        if filename:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, filename)

    def run(self):
        image = self.image_entry.get()
        ak = self.ak_entry.get()
        sk = self.sk_entry.get()
        output = self.output_entry.get()
        if not all([image, ak, sk, output]):
            messagebox.showerror("错误", "请完整填写所有信息")
            return
        try:
            run_recognition(image, ak, sk, output)
            messagebox.showinfo("成功", f"已识别表格并保存至 {output}")
        except Exception as e:
            messagebox.showerror("失败", str(e))

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
