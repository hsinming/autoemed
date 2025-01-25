# 📌 eMedical 502 Chest X-Ray Automation

![GitHub Repo Stars](https://img.shields.io/github/stars/hsinming/autoemed?style=social)
![GitHub Forks](https://img.shields.io/github/forks/hsinming/autoemed?style=social)
![GitHub License](https://img.shields.io/github/license/hsinming/autoemed)

🚀 **專案簡介**  
本專案使用 `Helium` 和 `Selenium` 自動化 eMedical 502 Chest X-Ray 正常案例的網頁登錄流程。透過 Python 及 GUI 界面，使用者可以快速登入 eMedical 系統並批次處理 Excel 檔案中的 eMedical No.。

---

## 🛠 功能特性
- 🔄 自動登入 eMedical 系統
- 📂 批次處理 eMedical No. 並填寫 502 Chest X-Ray 表單
- 🔍 根據 eMedical No. 前綴自動判別國家（澳大利亞、紐西蘭、加拿大、美國）
- 📋 GUI 操作介面，便於使用
- 🚀 `Headless` 模式支援背景執行
- 📜 自動紀錄日誌以追蹤處理狀況

---

## 📦 安裝與使用方式

### 1️⃣ 安裝 Conda 與環境設定
請確保你的環境已安裝 `conda`，如果尚未安裝，可至 [Miniconda](https://docs.conda.io/en/latest/miniconda.html) 或 [Anaconda](https://www.anaconda.com/) 官方網站下載並安裝。

安裝完成後，請使用以下指令來建立 `autoemed` 虛擬環境並安裝所有依賴：
```bash
# 建立 Conda 環境
conda env create -f environment.yml

# 啟動 Conda 環境
conda activate autoemed
```

若要確保環境內所有依賴已正確安裝，可執行：
```bash
conda list
```

### 2️⃣ 使用
```bash
python main.py  # 啟動 GUI 介面
```

或者使用 Nuitka 打包成獨立執行檔：
```bash
python -m nuitka main.py
```

### 3️⃣ 安裝 Nuitka
如果尚未安裝 Nuitka，可以使用以下指令安裝：
```bash
pip install nuitka
```

---

## 📂 專案結構
```
📁 eMedicalAutomation
│── 📄 main.py        # 主程式
│── 📄 environment.yml  # Conda 依賴清單
│── 📄 README.md      # 本文件
│── 📄 LICENSE        # 授權協議
│── 📄 log.txt        # 日誌紀錄
```

---

## 📜 授權條款
本專案採用 **MIT License** 授權，詳情請參閱 [LICENSE](LICENSE) 檔案。

---

## 📞 聯絡方式
如果有任何問題，請開啟 [Issue](https://github.com/hsinming/autoemed/issues) 或直接聯絡我。

📧 Email: hsinming.chen@gmail.com
