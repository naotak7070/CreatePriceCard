import csv
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.graphics.barcode import eanbc
from reportlab.graphics.shapes import Drawing
from reportlab.graphics import renderPDF
from reportlab.lib import utils
import qrcode
from io import BytesIO
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm

# フォント登録（Windows環境でメイリオを使う例）
pdfmetrics.registerFont(TTFont('Meiryo', 'C:/Windows/Fonts/meiryo.ttc'))

def create_price_cards(csv_file, output_pdf):
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

    # CSV読み込み
    with open(csv_file, 'r', encoding='utf-8') as f:
        reader = csv.DictReader(f)
        records = list(reader)

    for i, record in enumerate(records):
        page_index = i // cards_per_page
        card_index_on_page = i % cards_per_page
        col = card_index_on_page % cols
        row = card_index_on_page // cols

        # 新しいページの開始
        if card_index_on_page == 0 and i != 0:
            c.showPage()

        # カードの左下座標計算
        x = left_margin + col * card_width
        y = page_height - top_margin - (row+1)*card_height

        uuid = record['UUID']
        company = record['CompanyName']
        msrp = record['MSRP']
        sales_price = record['SalesPrice']
        product_name = record['ProductName']
        jan_code = record['JANCode']

        # QRコード生成(左上)
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
        # QRコード: 左上 (2mm内側に寄せる)
        c.drawImage(qr_img, x+2*mm, y+card_height - qr_img_height - 2*mm, 
                    width=qr_img_width, height=qr_img_height, preserveAspectRatio=True, mask='auto')

        # テキスト情報はQRコードの右側
        text_left = x + 2*mm + qr_img_width + 2*mm
        text_top = y + card_height - 2*mm
        
        # 商品名、会社名を13文字までに制限
        c.drawString(text_left, text_top, company[:13])            # 会社名(13文字まで)
        c.drawString(text_left, text_top - 10, product_name[:13])  # 商品名(13文字まで)
        c.drawString(text_left, text_top - 20, f"上代: {msrp}")
        c.drawString(text_left, text_top - 30, f"販売価格: {sales_price}")

        # JANバーコードを下から1/3の高さあたり(=高さを1/3に圧縮)
        # カード下から5mm上がった位置に描画
        barcode_y_position = y + 5*mm
        barcode = eanbc.Ean13BarcodeWidget(jan_code)
        # バーコードの高さをカード高さの1/3に設定 (高さ縮小)
        barcode.barHeight = card_height / 3.0
        
        # Drawing領域をバーコードより少し大きめに確保
        # 幅はカード幅分、高さは1/3にしたバーコード高さ分
        barcode_drawing = Drawing(card_width, card_height/3)
        barcode_drawing.add(barcode)
        
        # バーコードを描画 (左から2mm)
        renderPDF.draw(barcode_drawing, c, x+2*mm, barcode_y_position)

        # カード枠確認用（不要ならコメントアウト）
        # c.rect(x, y, card_width, card_height, stroke=1, fill=0)

    # 最終ページ出力
    c.showPage()
    c.save()

if __name__ == "__main__":
    create_price_cards("pricecard_template.csv", "output.pdf")
