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

# 2. 註冊字體：優先讀取專案目錄下的 msjh.ttc
font_name = "Helvetica" 
if os.path.exists("msjh.ttc"):
    try:
        pdfmetrics.registerFont(TTFont("CustomMSJH", "msjh.ttc"))
        font_name = "CustomMSJH"
    except:
        pass
elif os.path.exists("msjh.ttf"):
    try:
        pdfmetrics.registerFont(TTFont("CustomMSJH", "msjh.ttf"))
        font_name = "CustomMSJH"
    except:
        pass

# 3. 建立網頁檔案上傳器
uploaded_file = st.file_uploader("1. 請上傳您的標籤需求 Excel 檔案 (.xlsx 或 .csv)", type=["xlsx", "xls", "csv"])

if uploaded_file is not None:
    try:
        if uploaded_file.name.lower().endswith('.csv'):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
            
        st.success(f"📈 成功讀取到 {len(df)} 筆品項資料！")
        
        st.write("---")
        st.write("### 2. 點擊下方按鈕開始打包生成：")
        
        zip_buffer = io.BytesIO()
        
        if st.button("🚀 執行批次生成並打包下載", type="primary"):
            with st.spinner("正在為您精心繪製標籤並壓縮打包中，請稍候..."):
                with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zip_file:
                    
                    # 4. 初始化 ReportLab 畫布參數 (750x300 points)
                    width, height = 750, 300
                    styles = getSampleStyleSheet()
                    
                    # 品名樣式 (維持原樣)
                    title_style = ParagraphStyle(
                        'LabelTitle',
                        fontName=font_name,
                        fontSize=24,
                        leading=32,
                        textColor='#1a365d',
                    )
                    
                    # 5. 開始迴圈繪製
                    for index, row in df.iterrows():
                        filename = str(row['檔名']).strip()
                        name = str(row['品名']).strip()
                        expiry = str(row['效期']).strip()
                        if len(expiry) >= 10:
                            expiry = expiry[:10]
                        days = str(row['保存天數']).strip()
                        
                        pdf_buffer = io.BytesIO()
                        c = canvas.Canvas(pdf_buffer, pagesize=(width, height))
                        
                        # ======= 完美排版邏輯 =======
                        # 畫外框
                        c.setStrokeColorRGB(0.1, 0.21, 0.36)
                        c.setLineWidth(5)
                        c.roundRect(20, 20, width - 40, height - 40, 12, stroke=1, fill=0)
                        
                        # 畫品名 (自動換行)
                        p = Paragraph(name, title_style)
                        p.wrapOn(c, width - 90, height - 50)
                        p.drawOn(c, 45, height - 85)
                        
                        # 分隔線
                        c.setLineWidth(2)
                        c.setStrokeColorRGB(0.1, 0.21, 0.36)
                        c.line(40, height - 100, width - 40, height - 100)
                        
                        # --- 調大部分：有效日期 (字體加大至 28) ---
                        c.setFont(font_name, 28) 
                        c.setFillColorRGB(0.3, 0.35, 0.4) # 標籤灰色
                        c.drawString(45, height - 170, "有效日期：")
                        c.setFillColorRGB(0, 0, 0) # 數值黑色
                        c.drawString(195, height - 170, expiry)
                        
                        # --- 調大部分：保存天數 (字體加大至 28) ---
                        c.setFillColorRGB(0.3, 0.35, 0.4)
                        c.drawString(45, height - 235, "保存天數：")
                        c.setFillColorRGB(0, 0, 0)
                        c.drawString(195, height - 235, f"{days} 天")
                        
                        # 頁尾
                        c.setFillColorRGB(0.6, 0.6, 0.6)
                        c.setFont(font_name, 11)
                        c.drawRightString(width - 45, 35, "標籤自動生成系統")
                        
                        c.save()
                        pdf_buffer.seek(0)
                        zip_file.writestr(f"{filename}.pdf", pdf_buffer.read())
                
                zip_buffer.seek(0)
                st.success("✨ 字體加大版標籤已打包完畢！")
                
                st.download_button(
                    label="📥 下載字體加大版標籤 (.zip)",
                    data=zip_buffer,
                    file_name="字體加大版標籤.zip",
                    mime="application/zip",
                    use_container_width=True
                )
                
    except Exception as e:
        st.error(f"發生錯誤：{str(e)}")
