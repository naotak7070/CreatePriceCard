import streamlit as st
import pandas as pd

# 関数ファイルからインポート
from pricecards import create_price_cards_from_df_18_toc

st.title("PriceCardApp for to C")

# ラジオボタンでシートタイプを選択（デフォルトを18枚シートにする: index=0）

uploaded_file = st.file_uploader("Excelファイル（スマレジインポートデータ）をアップロードしてください", type=["xlsx", "xls"])
if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    st.write("アップロードされたファイル:")
    st.dataframe(df.head())

    if st.button("PDFを生成"):
        pdf_data = create_price_cards_from_df_18_toc(df)
    
        st.success("PDFが生成されました。")
        st.download_button(
            label="PDFをダウンロード",
            data=pdf_data,
            file_name="output.pdf",
            mime="application/pdf"
        )
