import os
import io
import zipfile
import pandas as pd
import streamlit as st
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import Paragraph
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

# 1. 網頁版面與佈局設定
st.set_page_config(page_title="標籤貼紙批次自動生成器", page_icon="🏷️", layout="centered")

st.title("🏷️ 標籤貼紙批次 PDF 生成工具")
st.write("請在下方上傳您的 Excel 檔案，系統會依原創完美格式批量生成獨立 PDF 並打包供您下載。")

# 2. 註冊字體：智慧識別本地 Windows 與雲端 Linux 環境
font_name = "Helvetica" # 預設安全牌

# 嘗試尋找 Windows 預設正黑體
win_font = r"C:\Windows\Fonts\msjh.ttc"
# 嘗試尋找 Linux (Streamlit 雲端) 常見的中文字體路徑
linux_font = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf" 

if os.path.exists(win_font):
    try:
        pdfmetrics.registerFont(TTFont("MSJH", win_font))
        font_name = "MSJH"
    except:
        pass
elif os.path.exists(linux_font):
    try:
        pdfmetrics.registerFont(TTFont("CustomFont", linux_font))
        font_name = "CustomFont"
    except:
        pass
else:
    # 如果雲端沒找到，直接註冊 ReportLab 自帶的 Helvetica (確保至少能順利開機跑完)
    font_name = "Helvetica"

# 3. 建立網頁檔案上傳器
uploaded_file = st.file_uploader("1. 請上傳您的標籤需求 Excel 檔案 (.xlsx 或 .csv)", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
        # 讀取上傳的資料
        if uploaded_file.name.lower().endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
            
        st.success(f"📈 成功讀取到 {len(df)} 筆品項資料！")
        
        # 資料預覽
        st.write("### 📊 資料預覽 (前 3 筆)：")
        st.dataframe(df.head(3))
        
        st.write("---")
        st.write("### 2. 點擊下方按鈕開始打包生成：")
        
        # 準備一個記憶體二進位緩衝區，用來存放最終的 ZIP 壓縮檔
        zip_buffer = io.BytesIO()
        
        # 當使用者點擊確認按鈕時才執行繪製
        if st.button("🚀 執行批次生成並打包下載", type="primary"):
            
            with st.spinner("正在為您精心繪製標籤並壓縮打包中，請稍候..."):
                
                # 建立記憶體 ZIP 壓縮檔
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    
                    # 4. 初始化 ReportLab 畫布參數 (1000x400 px -> 750x300 points)
                    width, height = 750, 300
                    styles = getSampleStyleSheet()
                    title_style = ParagraphStyle(
                        'LabelTitle',
                        fontName=font_name,
                        fontSize=24,
                        leading=32,
                        textColor='#1a365d',
                    )
                    
                    # 5. 開始迴圈繪製每一個獨立 PDF 標籤
                    for index, row in df.iterrows():
                        filename = str(row['檔名']).strip()
                        name = str(row['品名']).strip()
                        
                        expiry = str(row['效期']).strip()
                        if len(expiry) >= 10:
                            expiry = expiry[:10]
                            
                        days = str(row['保存天數']).strip()
                        
                        # 網頁版不寫入硬碟磁碟機，改用 io.BytesIO 存入記憶體
                        pdf_buffer = io.BytesIO()
                        c = canvas.Canvas(pdf_buffer, pagesize=(width, height))
                        
                        # ======= 完美還原您精調的排版邏輯 =======
                        # 畫外框
                        c.setStrokeColorRGB(0.1, 0.21, 0.36)
                        c.setLineWidth(5)
                        c.roundRect(20, 20, width - 40, height - 40, 12, stroke=1, fill=0)
                        
                        # 畫完整品名 (自動換行)
                        p = Paragraph(name, title_style)
                        p.wrapOn(c, width - 90, height - 50)
                        p.drawOn(c, 45, height - 85)
                        
                        # 分隔線
                        c.setLineWidth(2)
                        c.setStrokeColorRGB(0.1, 0.21, 0.36)
                        c.line(40, height - 100, width - 40, height - 100)
                        
                        # 內文
                        c.setFont(font_name, 20)
                        c.setFillColorRGB(0.3, 0.35, 0.4)
                        c.drawString(45, height - 165, "有效日期：")
                        c.setFillColorRGB(0, 0, 0)
                        c.drawString(165, height - 165, expiry)
                        
                        c.setFillColorRGB(0.3, 0.35, 0.4)
                        c.drawString(45, height - 225, "保存天數：")
                        c.setFillColorRGB(0, 0, 0)
                        c.drawString(165, height - 225, f"{days} 天")
                        
                        # 頁尾
                        c.setFillColorRGB(0.6, 0.6, 0.6)
                        c.setFont(font_name, 11)
                        c.drawRightString(width - 45, 35, "標籤自動生成系統")
                        
                        c.save()
                        # =======================================
                        
                        # 將繪製完成的單一 PDF 寫入 ZIP 壓縮包中
                        pdf_buffer.seek(0)
                        zip_file.writestr(f"{filename}.pdf", pdf_buffer.read())
                
                # 重新定位 ZIP 緩衝區到開頭以便下載
                zip_buffer.seek(0)
                
                st.success("✨ 所有標籤已成功完美繪製並打包完畢！")
                
                # 6. 亮出下載按鈕
                st.download_button(
                    label="📥 點擊下載標籤貼紙打包檔 (.zip)",
                    data=zip_buffer,
                    file_name="批次標籤貼紙成品.zip",
                    mime="application/zip",
                    use_container_width=True
                )
                
    except Exception as e:
        st.error(f"處理過程中發生失敗，請檢查檔案格式與表頭：\n{str(e)}")
