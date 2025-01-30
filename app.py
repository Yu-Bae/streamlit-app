import streamlit as st
import pandas as pd
import plotly.express as px

st.title('Aplikasi Streamlit dengan Dataset Excel dan Plotly')

# Pastikan nama file sesuai dataset lokal Anda
file_name = "Dec 2024.xlsx"

try:
    # Membaca file Excel dengan memastikan kolom ID tetap sebagai string
    df = pd.read_excel(file_name, dtype={"ID PRODUK": str})

    # Menampilkan dataframe
    st.write("### Data yang dimuat:", df)

    # Dropdown untuk memilih kolom filter yang digunakan pada semua chart
    filter_column = st.selectbox(
        "Pilih kolom data untuk ditampilkan di semua Chart:",
        ["AMOUNT REAL", "QTY REAL", "TOTAL LITER"]
    )

    # ================== Chart 1: By Depo ==================
    filtered_data_depo = df.groupby("NAMA DEPO")[filter_column].sum().reset_index()
    filtered_data_depo[filter_column] = filtered_data_depo[filter_column].round(0).astype(int)

    fig_depo = px.bar(
        filtered_data_depo,
        x="NAMA DEPO",
        y=filter_column,
        title=f"{filter_column} by Depo",
        text_auto=True,
        color=filter_column,
        labels={"NAMA DEPO": "Nama Depo", filter_column: f"Total {filter_column}"},
        template="simple_white"
    )
    st.plotly_chart(fig_depo, use_container_width=True)

    # ================== Chart 2: By Channel ==================
    filtered_data_channel = df.groupby("CHANNEL")[filter_column].sum().reset_index()
    filtered_data_channel[filter_column] = filtered_data_channel[filter_column].round(0).astype(int)

    fig_channel = px.pie(
        filtered_data_channel,
        names="CHANNEL",
        values=filter_column,
        title=f"{filter_column} by Channel",
        hole=0.3
    )
    st.plotly_chart(fig_channel, use_container_width=True)

    # ================== Chart 3: By SKU Report ==================
    filtered_data_sku = df.groupby("SKU REPORT")[filter_column].sum().reset_index()
    filtered_data_sku[filter_column] = filtered_data_sku[filter_column].round(0).astype(int)

    fig_sku = px.bar(
        filtered_data_sku,
        x="SKU REPORT",
        y=filter_column,
        title=f"{filter_column} by SKU Report",
        color=filter_column,
        text_auto=True,
        template="plotly_dark"
    )
    fig_sku.update_traces(textposition="outside")
    st.plotly_chart(fig_sku, use_container_width=True)

except FileNotFoundError:
    st.error(f"File '{file_name}' tidak ditemukan.")
except ValueError as e:
    st.error(f"Error saat membaca data: {e}")
