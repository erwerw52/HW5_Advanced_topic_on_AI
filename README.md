# Auto PPT Generator

小型專案，用於從文檔自動生成 PPT，支援 txt/md/docx，使用 Google Gemini API 分析內容，並透過 python-pptx 生成簡報。

快速開始

1. 建立虛擬環境：

```powershell
python -m venv venv
venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

2. 設定環境變數：

```powershell
copy .env.example .env
# 將 .env 中的 GEMINI_API_KEY 改為實際的值
```

3. 執行 Gradio 本地介面：

```powershell
python main.py
```

若無 Gemini API Key，系統會使用本地 heuristic 摘要作為替代方案。