import streamlit as st
import pandas as pd
import plotly.express as px
import os

# Memuat file CSS eksternal
with open("assets/style.css") as f:
    st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

st.title('Dashboard Sales Overview AQUA Division MTD Dec 2024')

# Pastikan nama file sesuai dataset lokal Anda
file_name = "Dec 2024.xlsx"

try:
    # Membaca file Excel dengan memastikan kolom ID tetap sebagai string
    df = pd.read_excel(file_name, dtype={"ID PRODUK": str})

    # Dropdown untuk memilih kolom filter yang digunakan pada semua chart
    filter_column = st.selectbox(
        "Pilih kolom data untuk ditampilkan di semua Chart:",
        ["Jumlah Penjualan", "Jumlah Kuantitas", "Total Liter"]
    )

    # Menerjemahkan pilihan dropdown ke nama kolom yang sesuai
    column_map = {
        "Jumlah Penjualan": "AMOUNT REAL",
        "Jumlah Kuantitas": "QTY REAL",
        "Total Liter": "TOTAL LITER"
    }
    filter_column_name = column_map[filter_column]

    # ================== Chart 1: By Depo ==================
    filtered_data_depo = df.groupby("NAMA DEPO")[filter_column_name].sum().reset_index()
    filtered_data_depo[filter_column_name] = filtered_data_depo[filter_column_name].astype(int)

    fig_depo = px.bar(
        filtered_data_depo,
        x="NAMA DEPO",
        y=filter_column_name,
        title=f"{filter_column_name} by Depo",
        text_auto=True,
        color=filter_column_name,
        labels={"NAMA DEPO": "Nama Depo", filter_column_name: f"Total {filter_column_name}"},
        template="plotly_white"
    )

    fig_depo.update_layout(
        yaxis_tickformat=",.0f"  # Menambahkan koma sebagai pemisah ribuan
    )

    st.plotly_chart(fig_depo, use_container_width=True)

    # ================== Chart 2: By Channel (Bar Chart) ==================
    filtered_data_channel = df.groupby("CHANNEL")[filter_column_name].sum().reset_index()
    filtered_data_channel[filter_column_name] = filtered_data_channel[filter_column_name].astype(int)

    fig_channel = px.bar(
        filtered_data_channel,
        x="CHANNEL",
        y=filter_column_name,
        title=f"{filter_column_name} by Channel",
        text_auto=True,
        color=filter_column_name,
        labels={"CHANNEL": "Channel", filter_column_name: f"Total {filter_column_name}"},
        template="plotly_white"
    )

    fig_channel.update_layout(
        yaxis_tickformat=",.0f"  # Menambahkan koma sebagai pemisah ribuan
    )

    st.plotly_chart(fig_channel, use_container_width=True)

    # ================== Chart 3: By SKU Report ==================
    filtered_data_sku = df.groupby("SKU REPORT")[filter_column_name].sum().reset_index()
    filtered_data_sku[filter_column_name] = filtered_data_sku[filter_column_name].astype(int)

    fig_sku = px.bar(
        filtered_data_sku,
        x="SKU REPORT",
        y=filter_column_name,
        title=f"{filter_column_name} by SKU Report",
        color=filter_column_name,
        text_auto=True,
        template="plotly_dark"
    )

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

    # ================== Pie Chart Week by QTY REAL ==================
    filtered_data_week = df.groupby("WEEK")["QTY REAL"].sum().reset_index()
    filtered_data_week["QTY REAL"] = filtered_data_week["QTY REAL"].astype(int)

    fig_week = px.pie(
        filtered_data_week,
        names="WEEK",
        values="QTY REAL",
        title="Week by QTY REAL",
        template="plotly_white"
    )

    # Menyesuaikan ukuran font dan warna teks pada Pie Chart agar Bold
    fig_week.update_traces(
        textfont=dict(size=15, color="black", family="Arial, sans-serif", weight="bold"),  # Menambahkan bold
        textinfo="percent+label",  # Menampilkan persentase dan label
        marker=dict(colors=['#FF7F0E', '#1F77B4', '#2CA02C', '#D62728', '#9467BD'])  # Warna sektornya
    )

    # ================== Membagi Kolom untuk Tabel dan Pie Chart ==================
    col1, col2 = st.columns([2, 1])  # Kolom pertama lebih lebar (2) untuk tabel, kolom kedua (1) untuk pie chart

    with col1:
        # Menampilkan tabel Top 10 NAMA PELANGGAN
        st.write("### Top 10 NAMA PELANGGAN berdasarkan Penjualan")
        st.write(top_10_customers[['Ranking', 'NAMA DEPO', 'NAMA PELANGGAN', 'AMOUNT REAL']].set_index('Ranking'))

    with col2:
        # Menampilkan Pie Chart
        st.plotly_chart(fig_week, use_container_width=True)

except FileNotFoundError:
    st.error(f"File '{file_name}' tidak ditemukan.")
except ValueError as e:
    st.error(f"Error saat membaca data: {e}")
except Exception as e:
    st.error(f"Terjadi kesalahan: {e}")
