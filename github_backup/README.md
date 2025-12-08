# IQC 效率管理系統

透過數據量化分析，分析 IQC 檢驗效率、工作負載、時間管理分配，從而協助提升 IQC 效能與品質水平。

## 🚀 功能特色

- 📊 **IQC 檢驗效率監控** - 即時監控檢驗效率指標
- ⏱️ **工作負載分析** - 分析人員工作負載與時間分配
- 👤 **人員能力分析** - 五維雷達圖視覺化能力評估
- 🔍 **MRB 分析** - 品質異常分析追蹤
- 🕵️ **異常行為偵測** - 自動識別異常檢驗行為
- 🎯 **效率 vs. 品質四象限分析** - 平衡效率與品質指標

## 📦 安裝方式

### 1. 確保環境要求
- Python 3.8 或以上版本

### 2. 安裝依賴套件
```bash
pip install -r requirements.txt
```

## ▶️ 執行方式

### 方式一：直接執行
```bash
streamlit run iqc_monitor_Opus_testV2.py
```

### 方式二：使用啟動腳本
```bash
python run_app.py
```

## 📄 所需資料格式

系統需要以下 Excel 檔案作為資料輸入：

| 檔案類型 | 說明 |
|---------|------|
| IQC Report | IQC 檢驗報告資料 |
| PCB 建檔明細 | PCB 產品建檔資訊 |
| PCB 標準工時對應表 | 標準工時設定 |
| IQC 額外任務紀錄清單 | 額外工作任務記錄 |

## 🛠️ 技術架構

| 項目 | 技術 |
|------|------|
| 前端框架 | Streamlit |
| 資料處理 | Pandas, NumPy |
| 視覺化 | Plotly, Altair |
| Excel 處理 | openpyxl, xlsxwriter |

## 📁 專案結構

```
IQC-Efficiency-System/
├── README.md                      # 專案說明
├── requirements.txt               # Python 依賴套件
├── .gitignore                     # Git 忽略檔案
├── iqc_monitor_Opus_testV2.py     # 主程式
├── run_app.py                     # 啟動腳本
├── 使用說明.txt                   # 使用說明文件
├── 打包指南.md                    # 打包說明文件
└── assets/                        # 資源檔案
    ├── IQC1.png
    ├── IQC2.png
    ├── IQC3.png
    ├── IQC4.ico
    └── IQC5.ico
```

## 📝 注意事項

1. 首次執行前請確認已安裝所有依賴套件
2. 上傳的 Excel 檔案格式需符合系統要求
3. 圖片資源位於 `assets/` 資料夾中

## 📜 License

此專案僅供內部使用。

---
*最後更新：2025年12月*
