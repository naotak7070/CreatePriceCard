import streamlit as st
import pandas as pd
import qrcode
from io import BytesIO
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.graphics.barcode import eanbc
from reportlab.graphics.shapes import Drawing
from reportlab.graphics import renderPDF
from reportlab.lib import utils
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import math

# フォント登録（Windows環境でメイリオを使う例）
# Streamlit Cloudの環境ではメイリオがない可能性があります。
# その場合はNotoSansCJKなど、CJKフォントをGitHubに同梱して使用することも可能です。
# ここでは仮にNotoSansCJK-Regular.ttfを同フォルダに配置している前提とします。
pdfmetrics.registerFont(TTFont('Meiryo', 'NotoSansJP-Regular.ttf'))

def is_empty_value(val):
    """valが空欄か判定する: NaNまたは空文字列ならTrue"""
    if pd.isna(val):
        return True
    if isinstance(val, str) and val.strip() == "":
        return True
    return False

def create_price_cards_from_df(df):
    # この関数はDataFrameを受け取り、PDFバイナリを返す
    # 前回のコードからExcel読み込み部分を除き、DataFrameをrecords化して利用します。
    
    pt_per_mm = 72.0/25.4
    page_width, page_height = A4  
    left_margin = 6 * pt_per_mm
    top_margin = 8.5 * pt_per_mm

    card_width = 66 * pt_per_mm
    card_height = 35 * pt_per_mm

    cols = 3
    rows = 8
    cards_per_page = cols * rows

    # PDFをメモリ上に生成
    pdf_buffer = BytesIO()
    c = canvas.Canvas(pdf_buffer, pagesize=A4)
    c.setFont("Meiryo", 8)

    needed_columns = ["id", "出展者名", "display_code", "jan", "name", "price", "retail_price"]
    records = df.to_dict(orient='records')

    card_count = 0
    for i, record in enumerate(records):
        # 全て空欄チェック
        if all(is_empty_value(record[col]) for col in needed_columns):
            continue

        card_index_on_page = card_count % cards_per_page
        col = card_index_on_page % cols
        row = card_index_on_page // cols

        # ページ切り替え
        if card_index_on_page == 0 and card_count != 0:
            c.showPage()
            c.setFont("Meiryo", 8)

        x = left_margin + col * card_width
        y = page_height - top_margin - (row+1)*card_height

        uuid = record['id']
        company = record['出展者名']
        display_code = record['display_code']
        jan_code = record['jan']
        product_name = record['name']
        msrp = record['price']
        sales_price = record['retail_price']

        # String化
        uuid = "" if is_empty_value(uuid) else str(uuid)
        company = "" if is_empty_value(company) else str(company)
        display_code = "" if is_empty_value(display_code) else str(display_code)
        product_name = "" if is_empty_value(product_name) else str(product_name)
        msrp = "" if is_empty_value(msrp) else str(msrp)
        sales_price = "" if is_empty_value(sales_price) else str(sales_price)

        # QRコード(UUID)
        if uuid.strip():
            qr = qrcode.QRCode(box_size=2, border=0)
            qr.add_data(uuid)
            qr.make()
            img_qr = qr.make_image(fill_color="black", back_color="white")

            qr_buf = BytesIO()
            img_qr.save(qr_buf, format='PNG')
            qr_buf.seek(0)
            qr_img = utils.ImageReader(qr_buf)

            qr_img_width = 15*mm
            qr_img_height = 15*mm
            c.drawImage(qr_img, x+2*mm, y+card_height - qr_img_height - 2*mm, 
                        width=qr_img_width, height=qr_img_height, preserveAspectRatio=True, mask='auto')

        text_left = x + 2*mm + (15*mm) + 2*mm
        text_top = y + card_height - 2*mm

        # テキスト表示
        c.drawString(text_left, text_top, company[:13])
        c.drawString(text_left, text_top - 10, product_name[:13])
        c.drawString(text_left, text_top - 20, display_code)
        c.drawString(text_left, text_top - 30, f"上代: {msrp}")
        c.drawString(text_left, text_top - 40, f"販売価格: {sales_price}")

        # JANコード
        if not is_empty_value(jan_code):
            jan_code = str(jan_code)
            if len(jan_code) < 13:
                jan_code = jan_code.zfill(13)
            elif len(jan_code) > 13:
                jan_code = jan_code[:13]

            if jan_code.isdigit():
                barcode_y_position = y + 5*mm
                barcode = eanbc.Ean13BarcodeWidget(jan_code)
                barcode.barHeight = card_height / 3.0
                barcode_drawing = Drawing(card_width, card_height/3)
                barcode_drawing.add(barcode)
                renderPDF.draw(barcode_drawing, c, x+2*mm, barcode_y_position)

        card_count += 1

    c.showPage()
    c.save()
    pdf_buffer.seek(0)
    return pdf_buffer.getvalue()

# StreamlitアプリUI
st.title("Price Card Generator")

uploaded_file = st.file_uploader("Excelファイルをアップロードしてください", type=["xlsx", "xls"])
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.write("アップロードされたファイル:")
    st.dataframe(df.head())

    if st.button("PDFを生成"):
        pdf_data = create_price_cards_from_df(df)
        st.success("PDFが生成されました。")
        st.download_button(
            label="PDFをダウンロード",
            data=pdf_data,
            file_name="output.pdf",
            mime="application/pdf"
        )
