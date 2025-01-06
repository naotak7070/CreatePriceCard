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
import math




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
    company_list = []  # ★ 会社名を格納するリストを用意

    for i, record in enumerate(records):
        if all(is_empty_value(record.get(col)) for col in needed_columns):
            continue

        card_index_on_page = card_count % cards_per_page
        col_idx = card_index_on_page % cols
        row_idx = card_index_on_page // cols

        if card_index_on_page == 0 and card_count != 0:
            c.showPage()  # 新しいページを作成
            c.setFont("Meiryo", 8)

        x = left_margin + col_idx * card_width
        y = page_height - top_margin - (row_idx + 1) * card_height

        uuid = safe_str(record.get('id', ''))
        company = safe_str(record.get('出展者名', ''))
        display_code = safe_str(record.get('number', ''))

        jan_value = record.get('jan', '')
        jan_code = None

        if isinstance(jan_value, float):
            if not math.isnan(jan_value):
                jan_code = int(jan_value)
        elif isinstance(jan_value, int):
            jan_code = jan_value
        elif isinstance(jan_value, str) and jan_value.isdigit():
            jan_code = int(jan_value)

        product_name = safe_str(record.get('name', ''))
        msrp = safe_str(record.get('retail_price', ''))  # 小売価格（参考上代）
        msrp_text = "オープン" if msrp == "" or msrp == "0" else msrp
        sales_price = safe_str(record.get('unit_price', ''))  # 卸単価
        lot = safe_str(record.get('lot', ''))  # 販売ロット

        company_list.append(company)  # ★ 会社名リストに追加

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
            c.drawImage(qr_img, x + 2 * mm, y + card_height - qr_img_height - 2 * mm,
                        width=qr_img_width, height=qr_img_height, preserveAspectRatio=True, mask='auto')

        # テキスト
        text_left = x + (15 * mm) + 4 * mm
        text_top = y + card_height - 3.5 * mm

        c.drawString(text_left, text_top, company[:13])
        c.drawString(text_left, text_top - 10, product_name[:13])
        c.drawString(text_left, text_top - 20, f"{display_code[:15]}")
        c.drawString(text_left, text_top - 30, f"参考上代: {msrp_text}")
        c.drawString(text_left, text_top - 40, f"卸販売単価: {sales_price}, ロット: {lot}")

        # JANコードの処理
        if jan_code is not None:  # jan_codeがNoneでない場合のみ処理を行う
            jan_code_str = str(jan_code)

            if len(jan_code_str) < 13:
                jan_code_str = jan_code_str.zfill(13)
            elif len(jan_code_str) > 13:
                jan_code_str = jan_code_str[:13]

            barcode = eanbc.Ean13BarcodeWidget(jan_code_str)
            barcode.barHeight = card_height / 3.0
            x1, y1, x2, y2 = barcode.getBounds()
            barcode_width = x2 - x1
            barcode_height = y2 - y1

            barcode_x = x + card_width - barcode_width - 2 * mm
            barcode_y = y + 2 * mm

            barcode_drawing = Drawing(barcode_width, barcode_height)
            barcode_drawing.add(barcode)
            renderPDF.draw(barcode_drawing, c, barcode_x, barcode_y)

        card_count += 1  # カード生成ごとにインクリメント

    c.save()  # PDF保存
    pdf_buffer.seek(0)
    return pdf_buffer.getvalue(), company_list, 24



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
    top_margin_mm = 21

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
    default_font_size = 8
    c.setFont("Meiryo", default_font_size)

    needed_columns = ["id", "出展者名", "display_code", "jan", "name", "price", "retail_price"]
    records = df.to_dict(orient="records")

    card_count = 0
    company_list = []  # ★ 会社名を格納するリストを用意

    for record in records:
        # すべて空欄ならスキップ
        if all(is_empty_value(record.get(col)) for col in needed_columns):
            continue

        card_index_on_page = card_count % cards_per_page
        col_idx = card_index_on_page % cols
        row_idx = card_index_on_page // cols

        # ページ切り替え (新ページ)
        if card_index_on_page == 0 and card_count != 0:
            c.showPage()
            c.setFont("Meiryo", default_font_size)

        # カードの左下座標
        x = left_margin + col_idx * card_width
        y = page_height - top_margin - (row_idx + 1) * card_height

        # レコード抽出＆文字列化
        uuid = safe_str(record.get("id", ""))
        company = safe_str(record.get("出展者名", ""))
        display_code = safe_str(record.get("number", ""))
        product_name = safe_str(record.get("name", ""))
        msrp = safe_str(record.get("retail_price", ""))
        sales_price = safe_str(record.get("unit_price", ""))
        lot = safe_str(record.get("lot", ""))

        # JANコードの取得と変換
        jan_code = parse_jan_code(record.get("jan"))

        # ★ 会社名リストに追加
        company_list.append(company)

        # QRコード
        if uuid:
            draw_qr_code(c, uuid, x, y, card_width, card_height)

        # テキスト表示
        text_left = x + 2 * mm + 15 * mm + 2 * mm
        text_top = y + card_height - 2 * mm
        c.drawString(text_left, text_top, company[:13])
        c.drawString(text_left, text_top - 10, product_name[:13])
        c.drawString(text_left, text_top - 20, f"{display_code[:15]}")

        # 一時的なフォントサイズ変更
        c.setFont("Meiryo", 10)
        c.drawString(text_left, text_top - 34, f"参考上代: {msrp if msrp else 'オープン'}")
        c.drawString(text_left, text_top - 48, f"卸販売単価: {sales_price}")
        c.drawString(text_left, text_top - 62, f"ロット数: {lot}")
        c.setFont("Meiryo", default_font_size)

        # JANコード
        if jan_code:
            draw_barcode(c, jan_code, x, y, card_width, card_height)

        card_count += 1

    c.showPage()
    c.save()
    pdf_buffer.seek(0)
    return pdf_buffer.getvalue(), company_list, 18


def parse_jan_code(jan_value):
    """JANコードを適切にパースする"""
    if isinstance(jan_value, float) and not math.isnan(jan_value):
        return str(int(jan_value)).zfill(13)
    elif isinstance(jan_value, int):
        return str(jan_value).zfill(13)
    elif isinstance(jan_value, str) and jan_value.isdigit():
        return jan_value.zfill(13)
    return None


def draw_qr_code(c, uuid, x, y, card_width, card_height):
    """QRコードを描画する"""
    qr = qrcode.QRCode(box_size=2, border=0)
    qr.add_data(uuid)
    qr.make()
    img_qr = qr.make_image(fill_color="black", back_color="white")
    qr_buf = BytesIO()
    img_qr.save(qr_buf, format="PNG")
    qr_buf.seek(0)
    qr_img = utils.ImageReader(qr_buf)

    qr_img_width = 15 * mm
    qr_img_height = 15 * mm
    c.drawImage(
        qr_img,
        x + 2 * mm,
        y + card_height - qr_img_height - 2 * mm,
        width=qr_img_width,
        height=qr_img_height,
        preserveAspectRatio=True,
        mask="auto",
    )


def draw_barcode(c, jan_code, x, y, card_width, card_height):
    """バーコードを描画する"""
    barcode = eanbc.Ean13BarcodeWidget(jan_code)
    barcode.barHeight = card_height / 3.0
    x1, y1, x2, y2 = barcode.getBounds()
    barcode_width = x2 - x1
    barcode_height = y2 - y1

    barcode_x = x + card_width - barcode_width - 2 * mm
    barcode_y = y + 2 * mm

    barcode_drawing = Drawing(barcode_width, barcode_height)
    barcode_drawing.add(barcode)
    renderPDF.draw(barcode_drawing, c, barcode_x, barcode_y)



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
        msrp_str = safe_str(record.get('商品単価', ''))  # 小売価格（参考上代）
        #税込価格換算
        # 1) 数値として変換できない場合は 0 とみなす or 例外処理をする
        try:
            # float()なら小数点も扱える
            msrp_val = float(msrp_str)
        except ValueError:
            msrp_val = 0

        # 2) 計算して丸める
        salesprice_intax = round(msrp_val * 1.1)  # 小数点以下を四捨五入する
        
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
        c.drawString(text_left, text_top - 36, f"税込 {salesprice_intax} 円")
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
        msrp_str = safe_str(record.get('商品単価', ''))  # 小売価格（参考上代）
        # 1) 数値として変換できない場合は 0 とみなす or 例外処理をする
        try:
            # float()なら小数点も扱える
            msrp_val = float(msrp_str)
        except ValueError:
            msrp_val = 0

        # 2) 計算して丸める
        salesprice_intax = round(msrp_val * 1.1)  # 小数点以下を四捨五入する
        
        
        # ★ 会社名リストに追加
        company_list.append(company)
        


        # テキスト
        text_left = x + (15 * mm) + 4 * mm
        text_top = y + card_height - 3.5*mm
        c.drawString(text_left, text_top, company[:13])
        c.drawString(text_left, text_top - 10, product_name[:13])
        c.drawString(text_left, text_top - 20, f"{display_code[:15]}")
          # ★ここで大きなサイズに変更(例: 10pt)
        c.setFont("Meiryo", 14)
        c.drawString(text_left, text_top - 36, f" 税込{salesprice_intax} 円")
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




def create_price_cards_from_df_24_fuku(df):
    """
    24枚シート(8行×3列)用のPDFを生成する例。
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
    company_list = []  # ★ 会社名を格納するリストを用意

    for i, record in enumerate(records):
        if all(is_empty_value(record.get(col)) for col in needed_columns):
            continue

        card_index_on_page = card_count % cards_per_page
        col_idx = card_index_on_page % cols
        row_idx = card_index_on_page // cols

        if card_index_on_page == 0 and card_count != 0:
            c.showPage()  # 新しいページを作成
            c.setFont("Meiryo", 8)

        x = left_margin + col_idx * card_width
        y = page_height - top_margin - (row_idx + 1) * card_height

        uuid = safe_str(record.get('id', ''))
        company = safe_str(record.get('出展者名', ''))
        company='福花園種苗（株）'
        display_code = safe_str(record.get('number', ''))

        jan_value = record.get('jan', '')
        jan_code = None

        if isinstance(jan_value, float):
            if not math.isnan(jan_value):
                jan_code = int(jan_value)
        elif isinstance(jan_value, int):
            jan_code = jan_value
        elif isinstance(jan_value, str) and jan_value.isdigit():
            jan_code = int(jan_value)

        product_name = safe_str(record.get('name', ''))
        msrp = safe_str(record.get('retail_price', ''))  # 小売価格（参考上代）
        msrp_text = "オープン" if msrp == "" or msrp == "0" else msrp
        sales_price = safe_str(record.get('unit_price', ''))  # 卸単価
        lot = safe_str(record.get('lot', ''))  # 販売ロット

        company_list.append(company)  # ★ 会社名リストに追加

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
            c.drawImage(qr_img, x + 2 * mm, y + card_height - qr_img_height - 2 * mm,
                        width=qr_img_width, height=qr_img_height, preserveAspectRatio=True, mask='auto')

        # テキスト
        text_left = x + (15 * mm) + 4 * mm
        text_top = y + card_height - 3.5 * mm

        # c.drawString(text_left, text_top, company[:13])
        c.drawString(text_left, text_top, product_name[:13])
        c.drawString(text_left, text_top - 10, product_name[14:27])
        c.drawString(text_left, text_top - 20, f"{display_code[:15]}")
        c.drawString(text_left, text_top - 30, f"参考上代: {msrp_text}")
        c.drawString(text_left, text_top - 40, f"卸販売単価: {sales_price}, ロット: {lot}")

        # JANコードの処理
        if jan_code is not None:  # jan_codeがNoneでない場合のみ処理を行う
            jan_code_str = str(jan_code)

            if len(jan_code_str) < 13:
                jan_code_str = jan_code_str.zfill(13)
            elif len(jan_code_str) > 13:
                jan_code_str = jan_code_str[:13]

            barcode = eanbc.Ean13BarcodeWidget(jan_code_str)
            barcode.barHeight = card_height / 3.0
            x1, y1, x2, y2 = barcode.getBounds()
            barcode_width = x2 - x1
            barcode_height = y2 - y1

            barcode_x = x + card_width - barcode_width - 2 * mm
            barcode_y = y + 2 * mm

            barcode_drawing = Drawing(barcode_width, barcode_height)
            barcode_drawing.add(barcode)
            renderPDF.draw(barcode_drawing, c, barcode_x, barcode_y)

        card_count += 1  # カード生成ごとにインクリメント

    c.save()  # PDF保存
    pdf_buffer.seek(0)
    return pdf_buffer.getvalue(), company_list, 24



def create_price_cards_from_df_18_fuku(df):
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
    top_margin_mm = 21

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
    default_font_size = 8
    c.setFont("Meiryo", default_font_size)

    needed_columns = ["id", "出展者名", "display_code", "jan", "name", "price", "retail_price"]
    records = df.to_dict(orient="records")

    card_count = 0
    company_list = []  # ★ 会社名を格納するリストを用意

    for record in records:
        # すべて空欄ならスキップ
        if all(is_empty_value(record.get(col)) for col in needed_columns):
            continue

        card_index_on_page = card_count % cards_per_page
        col_idx = card_index_on_page % cols
        row_idx = card_index_on_page // cols

        # ページ切り替え (新ページ)
        if card_index_on_page == 0 and card_count != 0:
            c.showPage()
            c.setFont("Meiryo", default_font_size)

        # カードの左下座標
        x = left_margin + col_idx * card_width
        y = page_height - top_margin - (row_idx + 1) * card_height

        # レコード抽出＆文字列化
        uuid = safe_str(record.get("id", ""))
        company = safe_str(record.get("出展者名", ""))
        company='福花園種苗（株）'
        display_code = safe_str(record.get("number", ""))
        product_name = safe_str(record.get("name", ""))
        msrp = safe_str(record.get("retail_price", ""))
        sales_price = safe_str(record.get("unit_price", ""))
        lot = safe_str(record.get("lot", ""))

        # JANコードの取得と変換
        jan_code = parse_jan_code(record.get("jan"))

        # ★ 会社名リストに追加
        company_list.append(company)

        # QRコード
        if uuid:
            draw_qr_code(c, uuid, x, y, card_width, card_height)

        # テキスト表示
        text_left = x + 2 * mm + 15 * mm + 2 * mm
        text_top = y + card_height - 2 * mm
        # c.drawString(text_left, text_top, company[:13])
        c.drawString(text_left, text_top, product_name[:13])
        c.drawString(text_left, text_top - 10, product_name[13:27])
        c.drawString(text_left, text_top - 20, f"{display_code[:15]}")

        # 一時的なフォントサイズ変更
        c.setFont("Meiryo", 10)
        c.drawString(text_left, text_top - 34, f"参考上代: {msrp if msrp else 'オープン'}")
        c.drawString(text_left, text_top - 48, f"卸販売単価: {sales_price}")
        c.drawString(text_left, text_top - 62, f"ロット数: {lot}")
        c.setFont("Meiryo", default_font_size)

        # JANコード
        if jan_code:
            draw_barcode(c, jan_code, x, y, card_width, card_height)

        card_count += 1

    c.showPage()
    c.save()
    pdf_buffer.seek(0)
    return pdf_buffer.getvalue(), company_list, 18

