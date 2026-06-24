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
- 📜 自動紀錄日誌以追蹤處理狀況

---

## 📦 安裝與使用方式

### 1️⃣ 安裝 uv 與環境設定
請確保你的環境已安裝 `uv`，如果尚未安裝，可參考 [uv 官方文件](https://docs.astral.sh/uv/getting-started/installation/) 安裝。

安裝完成後，執行以下指令同步所有依賴：
```bash
uv sync
```

### 2️⃣ 使用
```bash
uv run python main.py  # 啟動 GUI 介面
```

或者使用 Nuitka 打包成獨立執行檔（Windows）：
```powershell
.\build_app.ps1
```

---

## 📂 專案結構
```
📁 eMedicalAutomation
│── 📄 main.py           # 主程式
│── 📄 build_app.ps1     # Windows 打包腳本
│── 📄 pyproject.toml    # 依賴定義
│── 📄 README.md         # 本文件
│── 📄 LICENSE           # 授權協議
│── 📄 log.txt           # 日誌紀錄
```

---

## 📜 授權條款
本專案採用 **MIT License** 授權，詳情請參閱 [LICENSE](LICENSE) 檔案。

---

## 📞 聯絡方式
如果有任何問題，請開啟 [Issue](https://github.com/hsinming/autoemed/issues) 或直接聯絡我。

📧 Email: hsinming.chen@gmail.com
