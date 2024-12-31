import streamlit as st
import pandas as pd

# 関数ファイルからインポート
from pricecards import (
    create_price_cards_from_df_24,
    create_price_cards_from_df_18
)

st.title("PriceCardApp for to B")
st.write("印刷の際は「実際のサイズ」で印刷すること")

# ラジオボタンでシートタイプを選択（デフォルトを18枚シートにする: index=0）
sheet_option = st.radio(
    "作成するシートタイプを選択してください",
    ["18枚シート", "24枚シート"], 
    index=0  # ← これで「18枚シート」をデフォルトに
)

uploaded_file = st.file_uploader("Excelファイルをアップロードしてください", type=["xlsx", "xls"])
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.write("アップロードされたファイル:")
    st.dataframe(df.head())

    if st.button("PDFを生成"):
        # 選択に応じて処理を分岐
        if sheet_option == "18枚シート":
            pdf_data,company_list,layout = create_price_cards_from_df_18(df)
        else:
            pdf_data,company_list,layout = create_price_cards_from_df_24(df)

        st.success("PDFが生成されました。")
        st.download_button(
            label="PDFをダウンロード",
            data=pdf_data,
            file_name=f"output_toB_{company_list[0]}_{layout}.pdf",
            mime="application/pdf"
        )

