import qrcode
import pandas as pd
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



# ===== フォント登録 =====
pdfmetrics.registerFont(TTFont('Meiryo', 'NotoSansJP-Regular.ttf'))

def is_empty_value(val):
    """valが空欄か判定する: NaNまたは空文字列ならTrue"""
    if pd.isna(val):
        return True
    if isinstance(val, str) and val.strip() == "":
        return True
    return False

def safe_str(val):
    """NaNや' nan' などを空文字にして返す"""
    if pd.isna(val):
        return ""
    # 文字列型の場合に 'nan' とかが含まれていれば空文字へ
    sval = str(val).strip()
    if sval.lower() == "nan":
        return ""
    return sval

def create_price_cards_from_df_24(df):
    """
    24枚シート(8行×3列)用のPDFを生成する例。
    （既存コードをそのまま引用）
    """
    pt_per_mm = 72.0 / 25.4
    page_width, page_height = A4
    left_margin = 6 * pt_per_mm
    top_margin = 8.5 * pt_per_mm

    card_width = 66 * pt_per_mm
    card_height = 35 * pt_per_mm

    cols = 3
    rows = 8
    cards_per_page = cols * rows

    pdf_buffer = BytesIO()
    c = canvas.Canvas(pdf_buffer, pagesize=A4)
    c.setFont("Meiryo", 8)

    needed_columns = ["id", "出展者名", "display_code", "jan", "name", "price", "retail_price"]
    records = df.to_dict(orient="records")

    card_count = 0
    # ★ 会社名を格納するリストを用意
    company_list = []
    
    for i, record in enumerate(records):
        if all(is_empty_value(record.get(col)) for col in needed_columns):
            continue

        card_index_on_page = card_count % cards_per_page
        col_idx = card_index_on_page % cols
        row_idx = card_index_on_page // cols

        if card_index_on_page == 0 and card_count != 0:
            c.showPage()
            c.setFont("Meiryo", 8)

        x = left_margin + col_idx * card_width
        y = page_height - top_margin - (row_idx + 1) * card_height

        uuid = safe_str(record.get('id', ''))
        company = safe_str(record.get('出展者名', ''))
        display_code = safe_str(record.get('number', ''))
        jan_code = safe_str(record.get('jan', ''))
        product_name = safe_str(record.get('name', ''))
        msrp = safe_str(record.get('retail_price', ''))  # 小売価格（上代）
        sales_price = safe_str(record.get('unit_price', ''))  # 卸単価
        lot = safe_str(record.get('lot', ''))  # 販売ロット

        # ★ 会社名リストに追加
        company_list.append(company)

        # QRコード
        if uuid:
            qr = qrcode.QRCode(box_size=2, border=0)
            qr.add_data(uuid)
            qr.make()
            img_qr = qr.make_image(fill_color="black", back_color="white")
            qr_buf = BytesIO()
            img_qr.save(qr_buf, format='PNG')
            qr_buf.seek(0)
            qr_img = utils.ImageReader(qr_buf)

            qr_img_width = 15 * mm
            qr_img_height = 15 * mm
            c.drawImage(qr_img, x + 2*mm, y + card_height - qr_img_height - 2*mm,
                        width=qr_img_width, height=qr_img_height, preserveAspectRatio=True, mask='auto')

        # テキスト
        text_left = x + (15 * mm) + 4 * mm
        text_top = y + card_height - 2*mm
        c.drawString(text_left, text_top, company[:13])
        c.drawString(text_left, text_top - 10, product_name[:13])
        c.drawString(text_left, text_top - 20, f"{display_code[:15]}")
        c.drawString(text_left, text_top - 30, f"上代: {msrp}")
        c.drawString(text_left, text_top - 40, f"販売価格: {sales_price} Lot: {lot}")

        # JANコード
        if jan_code.isdigit():
            if len(jan_code) < 13:
                jan_code = jan_code.zfill(13)
            elif len(jan_code) > 13:
                jan_code = jan_code[:13]

            barcode = eanbc.Ean13BarcodeWidget(jan_code)
            barcode.barHeight = card_height / 3.0
            x1, y1, x2, y2 = barcode.getBounds()
            barcode_width = x2 - x1
            barcode_height = y2 - y1

            barcode_x = x + card_width - barcode_width - 2*mm
            barcode_y = y + 2*mm

            barcode_drawing = Drawing(barcode_width, barcode_height)
            barcode_drawing.add(barcode)
            renderPDF.draw(barcode_drawing, c, barcode_x, barcode_y)

        card_count += 1

    c.showPage()
    c.save()
    pdf_buffer.seek(0)
    return pdf_buffer.getvalue(), company_list,24





def create_price_cards_from_df_18(df):
    """
    18枚シート(3列×6行)用のPDFを生成する。
    要件:
      - 上下余白: 21mm
      - 左右余白: 19mm
      - カード1枚のサイズ: 幅57.3mm × 高さ42.3mm
      - レイアウト: 3列 × 6行
    """
    pt_per_mm = 72.0 / 25.4
    page_width, page_height = A4
    
    # 余白(mm)
    left_margin_mm = 19
    right_margin_mm = 19
    top_margin_mm = 21
    bottom_margin_mm = 21

    left_margin = left_margin_mm * pt_per_mm
    top_margin = top_margin_mm * pt_per_mm

    # カードサイズ
    card_width_mm = 57.3
    card_height_mm = 42.3

    card_width = card_width_mm * pt_per_mm
    card_height = card_height_mm * pt_per_mm

    # 3列x6行 = 18面
    cols = 3
    rows = 6
    cards_per_page = cols * rows

    # PDFをメモリ上に生成
    pdf_buffer = BytesIO()
    c = canvas.Canvas(pdf_buffer, pagesize=A4)
    c.setFont("Meiryo", 8)

    needed_columns = ["id", "出展者名", "display_code", "jan", "name", "price", "retail_price"]
    records = df.to_dict(orient='records')

    card_count = 0
    # ★ 会社名を格納するリストを用意
    company_list = []

    for i, record in enumerate(records):
        # すべて空欄ならスキップ
        if all(is_empty_value(record.get(col)) for col in needed_columns):
            continue

        card_index_on_page = card_count % cards_per_page
        col_idx = card_index_on_page % cols
        row_idx = card_index_on_page // cols

        # ページ切り替え (新ページ)
        if card_index_on_page == 0 and card_count != 0:
            c.showPage()
            c.setFont("Meiryo", 8)

        # カードの左下座標
        # row=0 が最上段なので、通常は y = page_height - top_margin - (row+1)*card_height で計算
        # col=0 が最左列
        x = left_margin + col_idx * card_width
        y = page_height - top_margin - (row_idx + 1) * card_height

        # レコード抽出＆文字列化
        uuid = safe_str(record.get('id', ''))
        company = safe_str(record.get('出展者名', ''))
        display_code = safe_str(record.get('number', ''))
        jan_code = safe_str(record.get('jan', ''))
        product_name = safe_str(record.get('name', ''))
        msrp = safe_str(record.get('retail_price', ''))  # 小売価格（上代）
        sales_price = safe_str(record.get('unit_price', ''))  # 卸単価
        lot = safe_str(record.get('lot', ''))  # 販売ロット
        # ★ 会社名リストに追加
        company_list.append(company)

        # QRコード
        if uuid:
            qr = qrcode.QRCode(box_size=2, border=0)
            qr.add_data(uuid)
            qr.make()
            img_qr = qr.make_image(fill_color="black", back_color="white")

            qr_buf = BytesIO()
            img_qr.save(qr_buf, format='PNG')
            qr_buf.seek(0)
            qr_img = utils.ImageReader(qr_buf)

            qr_img_width = 15 * mm
            qr_img_height = 15 * mm
            c.drawImage(
                qr_img,
                x + 2*mm,
                y + card_height - qr_img_height - 2*mm,
                width=qr_img_width,
                height=qr_img_height,
                preserveAspectRatio=True,
                mask='auto'
            )

        # テキスト表示位置
        text_left = x + 2*mm + 15*mm + 2*mm
        text_top = y + card_height - 2*mm

        # テキスト表示
        c.drawString(text_left, text_top, company[:13])
        c.drawString(text_left, text_top - 10, product_name[:13])
        c.drawString(text_left, text_top - 20, f"{display_code[:15]}")
        # ★ここで大きなサイズに変更(例: 10pt)
        c.setFont("Meiryo", 10)
        c.drawString(text_left, text_top - 34, f"上代: {msrp}")
        c.drawString(text_left, text_top - 48, f"販売価格: {sales_price}")
        c.drawString(text_left, text_top - 62, f"Lot: {lot}")

        # 使い終わったら、元のサイズ(8pt)に戻す
        c.setFont("Meiryo", 8)



        # JANコード (右下)
        if jan_code.isdigit():
            # 長さ補正
            if len(jan_code) < 13:
                jan_code = jan_code.zfill(13)
            elif len(jan_code) > 13:
                jan_code = jan_code[:13]

            barcode = eanbc.Ean13BarcodeWidget(jan_code)
            barcode.barHeight = card_height / 3.0
            x1, y1, x2, y2 = barcode.getBounds()
            barcode_width = x2 - x1
            barcode_height = y2 - y1

            barcode_x = x + card_width - barcode_width - 2*mm
            barcode_y = y + 2*mm

            barcode_drawing = Drawing(barcode_width, barcode_height)
            barcode_drawing.add(barcode)
            renderPDF.draw(barcode_drawing, c, barcode_x, barcode_y)

        card_count += 1

    c.showPage()
    c.save()
    pdf_buffer.seek(0)
    return pdf_buffer.getvalue(),company_list,18


def create_price_cards_from_df_18_toc(df):
    """
    18枚シート(3列×6行)用のPDFを生成する。
    要件:
      - 上下余白: 21mm
      - 左右余白: 19mm
      - カード1枚のサイズ: 幅57.3mm × 高さ42.3mm
      - レイアウト: 3列 × 6行
    """
    pt_per_mm = 72.0 / 25.4
    page_width, page_height = A4
    
    # 余白(mm)
    left_margin_mm = 19
    right_margin_mm = 19
    top_margin_mm = 21
    bottom_margin_mm = 21

    left_margin = left_margin_mm * pt_per_mm
    top_margin = top_margin_mm * pt_per_mm

    # カードサイズ
    card_width_mm = 57.3
    card_height_mm = 42.3

    card_width = card_width_mm * pt_per_mm
    card_height = card_height_mm * pt_per_mm

    # 3列x6行 = 18面
    cols = 3
    rows = 6
    cards_per_page = cols * rows

    # PDFをメモリ上に生成
    pdf_buffer = BytesIO()
    c = canvas.Canvas(pdf_buffer, pagesize=A4)
    c.setFont("Meiryo", 8)

    needed_columns = ["タグ", "品番", "商品コード", "商品名", "商品単価"]
    records = df.to_dict(orient='records')

    card_count = 0
    # ★ 会社名を格納するリストを用意
    company_list = []

    for i, record in enumerate(records):
        # すべて空欄ならスキップ
        if all(is_empty_value(record.get(col)) for col in needed_columns):
            continue

        card_index_on_page = card_count % cards_per_page
        col_idx = card_index_on_page % cols
        row_idx = card_index_on_page // cols

        # ページ切り替え (新ページ)
        if card_index_on_page == 0 and card_count != 0:
            c.showPage()
            c.setFont("Meiryo", 8)

        # カードの左下座標
        # row=0 が最上段なので、通常は y = page_height - top_margin - (row+1)*card_height で計算
        # col=0 が最左列
        x = left_margin + col_idx * card_width
        y = page_height - top_margin - (row_idx + 1) * card_height

        # レコード抽出＆文字列化
        # uuid = safe_str(record.get('id', ''))
        company = safe_str(record.get('タグ', ''))
        display_code = safe_str(record.get('品番', ''))
        jan_code = safe_str(record.get('商品コード', ''))
        product_name = safe_str(record.get('商品名', ''))
        msrp = safe_str(record.get('商品単価', ''))  # 小売価格（上代）
        # ★ 会社名リストに追加
        company_list.append(company)
        

        # テキスト表示位置
        text_left = x + 2*mm + 15*mm + 2*mm
        text_top = y + card_height - 2*mm

        # テキスト表示
        c.drawString(text_left, text_top, company[:13])
        c.drawString(text_left, text_top - 10, product_name[:13])
        c.drawString(text_left, text_top - 20, f"{display_code[:15]}")
        # ★ここで大きなサイズに変更(例: 10pt)
        c.setFont("Meiryo", 14)
        c.drawString(text_left, text_top - 36, f"税抜 {msrp} 円")
        # c.drawString(text_left, text_top - 50, f"Lot: {lot}")

        # 使い終わったら、元のサイズ(8pt)に戻す
        c.setFont("Meiryo", 8)



        # JANコード (右下)
        if jan_code.isdigit():
            # 長さ補正
            if len(jan_code) < 13:
                jan_code = jan_code.zfill(13)
            elif len(jan_code) > 13:
                jan_code = jan_code[:13]

            barcode = eanbc.Ean13BarcodeWidget(jan_code)
            barcode.barHeight = card_height / 3.0
            x1, y1, x2, y2 = barcode.getBounds()
            barcode_width = x2 - x1
            barcode_height = y2 - y1

            barcode_x = x + card_width - barcode_width - 2*mm
            barcode_y = y + 2*mm

            barcode_drawing = Drawing(barcode_width, barcode_height)
            barcode_drawing.add(barcode)
            renderPDF.draw(barcode_drawing, c, barcode_x, barcode_y)

        card_count += 1

    c.showPage()
    c.save()
    pdf_buffer.seek(0)
    return pdf_buffer.getvalue(),company_list,18

def create_price_cards_from_df_24_toc(df):
    """
    24枚シート(8行×3列)用のPDFを生成する例。
    （既存コードをそのまま引用）
    """
    pt_per_mm = 72.0 / 25.4
    page_width, page_height = A4
    left_margin = 6 * pt_per_mm
    top_margin = 8.5 * pt_per_mm

    card_width = 66 * pt_per_mm
    card_height = 35 * pt_per_mm

    cols = 3
    rows = 8
    cards_per_page = cols * rows

    pdf_buffer = BytesIO()
    c = canvas.Canvas(pdf_buffer, pagesize=A4)
    c.setFont("Meiryo", 8)

    needed_columns = ["タグ", "品番", "商品コード", "商品名", "商品単価"]
    records = df.to_dict(orient='records')

    card_count = 0
    # ★ 会社名を格納するリストを用意
    company_list = []

    for i, record in enumerate(records):
        if all(is_empty_value(record.get(col)) for col in needed_columns):
            continue

        card_index_on_page = card_count % cards_per_page
        col_idx = card_index_on_page % cols
        row_idx = card_index_on_page // cols

        if card_index_on_page == 0 and card_count != 0:
            c.showPage()
            c.setFont("Meiryo", 8)

        x = left_margin + col_idx * card_width
        y = page_height - top_margin - (row_idx + 1) * card_height

        # レコード抽出＆文字列化
        # uuid = safe_str(record.get('id', ''))
        company = safe_str(record.get('タグ', ''))
        display_code = safe_str(record.get('品番', ''))
        jan_code = safe_str(record.get('商品コード', ''))
        product_name = safe_str(record.get('商品名', ''))
        msrp = safe_str(record.get('商品単価', ''))  # 小売価格（上代）
        # ★ 会社名リストに追加
        company_list.append(company)
        


        # テキスト
        text_left = x + (15 * mm) + 4 * mm
        text_top = y + card_height - 2*mm
        c.drawString(text_left, text_top, company[:13])
        c.drawString(text_left, text_top - 10, product_name[:13])
        c.drawString(text_left, text_top - 20, f"{display_code[:15]}")
          # ★ここで大きなサイズに変更(例: 10pt)
        c.setFont("Meiryo", 14)
        c.drawString(text_left, text_top - 36, f" {msrp} 円")
        # 使い終わったら、元のサイズ(8pt)に戻す
        c.setFont("Meiryo", 8)


        # JANコード
        if jan_code.isdigit():
            if len(jan_code) < 13:
                jan_code = jan_code.zfill(13)
            elif len(jan_code) > 13:
                jan_code = jan_code[:13]

            barcode = eanbc.Ean13BarcodeWidget(jan_code)
            barcode.barHeight = card_height / 3.0
            x1, y1, x2, y2 = barcode.getBounds()
            barcode_width = x2 - x1
            barcode_height = y2 - y1

            barcode_x = x + card_width - barcode_width - 2*mm
            barcode_y = y + 2*mm

            barcode_drawing = Drawing(barcode_width, barcode_height)
            barcode_drawing.add(barcode)
            renderPDF.draw(barcode_drawing, c, barcode_x, barcode_y)

        card_count += 1

    c.showPage()
    c.save()
    pdf_buffer.seek(0)
    return pdf_buffer.getvalue(),company_list,24