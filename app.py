import streamlit as st
import pandas as pd
import plotly.express as px

st.title('Aplikasi Streamlit dengan Dataset Excel dan Plotly')

# Pastikan nama file sesuai dataset lokal Anda
file_name = "Dec 2024.xlsx"

try:
    # Membaca file Excel dengan memastikan kolom ID tetap sebagai string
    df = pd.read_excel(file_name, dtype={"ID PRODUK": str})

    # Dropdown untuk memilih kolom filter yang digunakan pada semua chart
    filter_column = st.selectbox(
        "Pilih kolom data untuk ditampilkan di semua Chart:",
        ["AMOUNT REAL", "QTY REAL", "TOTAL LITER"]
    )

    # ================== Chart 1: By Depo ==================
    filtered_data_depo = df.groupby("NAMA DEPO")[filter_column].sum().reset_index()
    filtered_data_depo[filter_column] = filtered_data_depo[filter_column].astype(int)

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

    # Menonaktifkan pemformatan otomatis dan menambahkan format ribuan
    fig_depo.update_layout(
        yaxis_tickformat=",.0f"  # Menambahkan koma sebagai pemisah ribuan
    )

    st.plotly_chart(fig_depo, use_container_width=True)

    # ================== Chart 2: By Channel ==================
    filtered_data_channel = df.groupby("CHANNEL")[filter_column].sum().reset_index()
    filtered_data_channel[filter_column] = filtered_data_channel[filter_column].astype(int)

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
    filtered_data_sku[filter_column] = filtered_data_sku[filter_column].astype(int)

    fig_sku = px.bar(
        filtered_data_sku,
        x="SKU REPORT",
        y=filter_column,
        title=f"{filter_column} by SKU Report",
        color=filter_column,
        text_auto=True,
        template="plotly_dark"
    )

    # Menonaktifkan pemformatan otomatis dan menambahkan format ribuan
    fig_sku.update_layout(
        yaxis_tickformat=",.0f"  # Menambahkan koma sebagai pemisah ribuan
    )

    fig_sku.update_traces(textposition="outside")
    st.plotly_chart(fig_sku, use_container_width=True)

    # ================== Tabel Top 10 NAMA PELANGGAN ==================
    customer_sales = df.groupby(["NAMA DEPO", "NAMA PELANGGAN"])[["AMOUNT REAL"]].sum().reset_index()
    top_10_customers = customer_sales.sort_values(by="AMOUNT REAL", ascending=False).head(10)
    top_10_customers['Ranking'] = top_10_customers.reset_index().index + 1
    top_10_customers['AMOUNT REAL'] = top_10_customers['AMOUNT REAL'].astype(int)

    # Menambahkan koma sebagai pemisah ribuan untuk angka di tabel
    top_10_customers['AMOUNT REAL'] = top_10_customers['AMOUNT REAL'].apply(lambda x: f"{x:,}")

    # Menampilkan tabel Top 10 NAMA PELANGGAN
    st.write("### Top 10 NAMA PELANGGAN berdasarkan Penjualan")
    st.write(top_10_customers[['Ranking', 'NAMA DEPO', 'NAMA PELANGGAN', 'AMOUNT REAL']].set_index('Ranking'))

except FileNotFoundError:
    st.error(f"File '{file_name}' tidak ditemukan.")
except ValueError as e:
    st.error(f"Error saat membaca data: {e}")
