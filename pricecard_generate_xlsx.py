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
pdfmetrics.registerFont(TTFont('Meiryo', 'C:/Windows/Fonts/meiryo.ttc'))

def is_empty_value(val):
    """valが空欄か判定する: NaNまたは空文字列ならTrue"""
    if pd.isna(val):
        return True
    if isinstance(val, str) and val.strip() == "":
        return True
    return False

def create_price_cards(excel_file, output_pdf):
    # ページ設定
    page_width, page_height = A4  
    pt_per_mm = 72.0/25.4

    left_margin = 6 * pt_per_mm
    top_margin = 8.5 * pt_per_mm

    card_width = 66 * pt_per_mm
    card_height = 35 * pt_per_mm

    # 列数、行数(各ページ)
    cols = 3
    rows = 8
    cards_per_page = cols * rows

    # PDF初期化
    c = canvas.Canvas(output_pdf, pagesize=A4)
    c.setFont("Meiryo", 8)  # 日本語表示可能なフォント

    # Excel読み込み
    df = pd.read_excel(excel_file)

    # 必要な列が存在するかチェック
    needed_columns = ["id", "出展者名", "display_code", "jan", "name", "price", "retail_price"]
    for col in needed_columns:
        if col not in df.columns:
            raise ValueError(f"Excelに必要な列 '{col}' が存在しません。")

    records = df.to_dict(orient='records')

    for i, record in enumerate(records):
        # すべて空欄かどうかチェック
        # 必要なフィールド全てが空なら、そのカードをスキップする
        if all(is_empty_value(record[col]) for col in needed_columns):
            # 全て空欄なのでスキップ
            continue

        # ページ/カード位置計算
        card_index_on_page = i % cards_per_page
        col = card_index_on_page % cols
        row = card_index_on_page // cols

        # ページ切り替え（新しいページを開始）
        if card_index_on_page == 0 and i != 0:
            c.showPage()
            c.setFont("Meiryo", 8)

        # カード座標
        x = left_margin + col * card_width
        y = page_height - top_margin - (row+1)*card_height

        # レコードから必要データを抽出
        uuid = record['id']
        company = record['出展者名']
        display_code = record['display_code']
        jan_code = record['jan']
        product_name = record['name']
        msrp = record['price']
        sales_price = record['retail_price']

        # 文字列化・NaN対応
        if not is_empty_value(uuid):
            uuid = str(uuid)
        else:
            uuid = ""  # QRコードが空のケース

        company = str(company) if not is_empty_value(company) else ""
        display_code = str(display_code) if not is_empty_value(display_code) else ""
        product_name = str(product_name) if not is_empty_value(product_name) else ""
        msrp = str(msrp) if not is_empty_value(msrp) else ""
        sales_price = str(sales_price) if not is_empty_value(sales_price) else ""

        # QRコード生成(UUIDが空でない場合のみ)
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
            # QRコード: 左上
            c.drawImage(qr_img, x+2*mm, y+card_height - qr_img_height - 2*mm, 
                        width=qr_img_width, height=qr_img_height, preserveAspectRatio=True, mask='auto')

        # テキスト情報(13文字まで)
        text_left = x + 2*mm + (15*mm) + 2*mm  # QR幅分考慮(15mm) + 2mm余白
        text_top = y + card_height - 2*mm
        
        # レイアウト: 会社名 → 商品名 → display_code → 上代 → 販売価格
        c.drawString(text_left, text_top, company[:13])
        c.drawString(text_left, text_top - 10, product_name[:13])
        c.drawString(text_left, text_top - 20, display_code)
        c.drawString(text_left, text_top - 30, f"上代: {msrp}")
        c.drawString(text_left, text_top - 40, f"販売価格: {sales_price}")

        # JANコード表示
        if not is_empty_value(jan_code):
            jan_code = str(jan_code)
            # バーコードとして有効な13桁数字かチェック・修正
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

    c.showPage()
    c.save()


if __name__ == "__main__":
    # 例: "input.xlsx" から "output.pdf" を生成
    create_price_cards("input.xlsx", "output.pdf")
