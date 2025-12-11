# 🎮 PPT 風格轉換器

一個遊戲風格的 PowerPoint 簡報風格轉換工具，讓你的簡報瞬間變身！

## 🚀 快速開始

### 1. 安裝依賴

```bash
pip install -r requirements.txt
```

### 2. 啟動應用程式

```bash
cd src
streamlit run app.py
```

### 3. 開始使用

1. 📤 上傳你的 .pptx 檔案
2. 🚀 點擊「開始魔法轉換」
3. 💾 自動生成兩種風格並下載

## 🎨 支援的風格

- **🌟 Maeve 風格** - 現代科技感設計
- **🎨 水彩有機形狀風格** - 藝術感十足的水彩風格

## 📁 專案結構

```
HW5_Advanced_topic_on_AI/
├── src/
│   ├── app.py              # Streamlit 網頁應用
│   └── process_ppt.py      # PPT 處理核心程式
├── ppt/
│   └── template/           # 模板資料夾
│       ├── Maeve.pptx
│       └── WatercolorOrganicShapes.pptx
├── requirements.txt
└── README.md
```

## ✨ 功能特色

- 🎮 遊戲風格介面設計
- 🎨 自動生成兩種風格模板
- 📊 轉換統計與成就系統
- 💫 動畫效果與視覺回饋
- 📥 支援重複下載兩種版型
- ⚡ 無需本地儲存，使用臨時目錄處理

## 🛠️ 技術棧

- **Streamlit** - 網頁應用框架
- **python-pptx** - PowerPoint 處理庫
- **Python 3.8+**

## 📝 注意事項

- 只支援 .pptx 格式的檔案
- 建議檔案大小小於 50MB
- 處理時間依檔案複雜度而定

## 🎯 使用提示

1. 確保模板檔案存在於 `ppt/template/` 資料夾
2. 上傳的簡報會保留原始內容和圖片
3. 轉換後會自動生成兩種風格
4. 生成的檔案可以重複下載，不會因為下載一次就消失

## 🤝 貢獻

歡迎提出問題和建議！

---

Made with 💜 by PPT Master