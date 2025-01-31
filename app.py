import streamlit as st
import pandas as pd
import plotly.express as px
import os

# emojis: https://www.webfx.com/tools/emoji-cheat-sheet/
st.set_page_config(page_title="Sales Trend - AQUA Dec 2024", page_icon=":bar_chart:", layout="wide")

# Memuat file CSS eksternal
with open("assets/style.css") as f:
    st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

st.title('Sales Trend - AQUA Division Dec 2024')

# Pastikan nama file sesuai dataset lokal Anda
file_name = "mst_sales.xlsx"

try:
    # Membaca file Excel dengan memastikan kolom ID tetap sebagai string
    df = pd.read_excel(file_name, dtype={"ID PRODUK": str})

    # Pastikan TANGGAL diubah menjadi format datetime
    df['TANGGAL'] = pd.to_datetime(df['TANGGAL'], errors='coerce')

    # Format tanggal menjadi hanya Hari dan Bulan (misalnya: 1 Dec, 2 Dec, dst)
    df['TANGGAL'] = df['TANGGAL'].dt.strftime('%d %b')  # Format: '01 Dec'

    # TOP KPI's
    id_customer = df["ID PELANGGAN"].nunique()
    ao_customer = int(round(id_customer, 0))
    total_sales = int(df["TOTAL LITER"].sum())
    average_sales_by_liter = round(df["TOTAL LITER"].mean(), 2)

    # Kolom untuk meratakan angka
    left_column, middle_column, right_column = st.columns(3)

    # Menampilkan total LITER, Average Sales per Day, dan Total Active Outlet
    with left_column:
        st.markdown("<div class='center-text'>Total LITER:</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='center-text'>{total_sales:,}</div>", unsafe_allow_html=True)
    with middle_column:
        st.markdown("<div class='center-text'>Average Sales per Day:</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='center-text'>{average_sales_by_liter}</div>", unsafe_allow_html=True)
    with right_column:
        st.markdown("<div class='center-text'>Total Active Outlet:</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='center-text'>{ao_customer:,}</div>", unsafe_allow_html=True)

    st.markdown("""---""")

    # Dropdown untuk memilih kolom filter yang digunakan pada semua chart
    filter_column = st.selectbox(
        "Filter Total: ",
        ["in Sales", "in Quantity", "in Liter"]
    )

    # Menerjemahkan pilihan dropdown ke nama kolom yang sesuai
    column_map = {
        "in Sales": "AMOUNT REAL",
        "in Quantity": "QTY REAL",
        "in Liter": "TOTAL LITER"
    }
    filter_column_name = column_map[filter_column]

    # ================== Sales per Day by DEPO ==================
    # Mengelompokkan data berdasarkan NAMA DEPO dan TANGGAL dan menjumlahkan nilai yang dipilih
    sales_per_day_by_depo = df.groupby(['NAMA DEPO', 'TANGGAL'])[filter_column_name].sum().reset_index()

    # Memutar data sehingga TANGGAL menjadi kolom
    sales_per_day_depo_pivot = sales_per_day_by_depo.pivot_table(
        index='NAMA DEPO',  # Baris berdasarkan Nama Depo
        columns='TANGGAL',  # Kolom berdasarkan Tanggal
        values=filter_column_name,  # Nilai yang ditampilkan adalah kolom yang dipilih
        aggfunc='sum',  # Menjumlahkan nilai jika ada beberapa entri untuk satu depo dan tanggal
        fill_value=0  # Jika tidak ada data, isi dengan 0
    )

    # ================== Sales per Day by SKU ==================
    # Mengelompokkan data berdasarkan SKU REPORT dan TANGGAL dan menjumlahkan nilai yang dipilih
    sales_per_day_by_sku = df.groupby(['SKU REPORT', 'TANGGAL'])[filter_column_name].sum().reset_index()

    # Memutar data sehingga TANGGAL menjadi kolom
    sales_per_day_sku_pivot = sales_per_day_by_sku.pivot_table(
        index='SKU REPORT',  # Baris berdasarkan SKU REPORT
        columns='TANGGAL',  # Kolom berdasarkan Tanggal
        values=filter_column_name,  # Nilai yang ditampilkan adalah kolom yang dipilih
        aggfunc='sum',  # Menjumlahkan nilai jika ada beberapa entri untuk satu SKU dan tanggal
        fill_value=0  # Jika tidak ada data, isi dengan 0
    )

    # ================== Menampilkan Tabel Sales per Day ==================
    col1, col2 = st.columns(2)  # Membagi layar menjadi dua kolom

    with col1:
        st.markdown(f"<h3 style='text-align: center;'>Sales per Day by DEPO - Filter: {filter_column}</h3>", unsafe_allow_html=True)
        st.dataframe(sales_per_day_depo_pivot, use_container_width=True)

    with col2:
        st.markdown(f"<h3 style='text-align: center;'>Sales per Day by SKU - Filter: {filter_column}</h3>", unsafe_allow_html=True)
        st.dataframe(sales_per_day_sku_pivot, use_container_width=True)

    st.markdown("""---""")

    # ================== Chart 1: By Depo ==================
    filtered_data_depo = df.groupby("NAMA DEPO")[filter_column_name].sum().reset_index()
    filtered_data_depo[filter_column_name] = filtered_data_depo[filter_column_name].astype(int)

    fig_depo = px.bar(
        filtered_data_depo,
        x="NAMA DEPO",
        y=filter_column_name,
        title="Sales by Branch",
        text_auto=True,
        color=filter_column_name,
        labels={"NAMA DEPO": "Nama Depo", filter_column_name: f"Total {filter_column_name}"},
        template="plotly_white"
    )

    fig_depo.update_layout(
        yaxis_tickformat=",.0f"  # Menambahkan koma sebagai pemisah ribuan
    )

    # Menonaktifkan legenda
    showlegend=False  # Menonaktifkan legenda

    st.plotly_chart(fig_depo, use_container_width=True)

    # ================== Chart 2: By Channel (Bar Chart) ==================
    filtered_data_channel = df.groupby("CHANNEL")[filter_column_name].sum().reset_index()
    filtered_data_channel[filter_column_name] = filtered_data_channel[filter_column_name].astype(int)

    fig_channel = px.bar(
        filtered_data_channel,
        x="CHANNEL",
        y=filter_column_name,
        title="Sales by Channel",
        text_auto=True,
        color=filter_column_name,
        labels={"CHANNEL": "Channel", filter_column_name: f"Total {filter_column_name}"},
        template="plotly_white"
    )

    fig_channel.update_layout(
        yaxis_tickformat=",.0f"  # Menambahkan koma sebagai pemisah ribuan
    )

    # Menonaktifkan legenda
    showlegend=False  # Menonaktifkan legenda

    st.plotly_chart(fig_channel, use_container_width=True)

    # ================== Chart 3: By SKU Report ==================
    filtered_data_sku = df.groupby("SKU REPORT")[filter_column_name].sum().reset_index()
    filtered_data_sku[filter_column_name] = filtered_data_sku[filter_column_name].astype(int)

    fig_sku = px.bar(
        filtered_data_sku,
        x="SKU REPORT",
        y=filter_column_name,
        title="Sales by SKU",
        color=filter_column_name,
        text_auto=True,
        template="plotly_dark"
    )

    fig_sku.update_layout(
        yaxis_tickformat=",.0f"  # Menambahkan koma sebagai pemisah ribuan
    )

    fig_sku.update_traces(textposition="outside")

    # Menonaktifkan legenda
    showlegend=False  # Menonaktifkan legenda

    st.plotly_chart(fig_sku, use_container_width=True)

    # ================== Tabel Top 25 NAMA PELANGGAN ==================
    customer_sales = df.groupby(["NAMA DEPO", "ID PELANGGAN", "NAMA PELANGGAN","TOTAL LITER","QTY REAL"])[["AMOUNT REAL"]].sum().reset_index()
    top_10_customers = customer_sales.sort_values(by="AMOUNT REAL", ascending=False).head(25)
    top_10_customers['Ranking'] = top_10_customers.reset_index().index + 1

    # Mengubah 'AMOUNT REAL' menjadi format angka dengan koma
    top_10_customers['AMOUNT REAL'] = pd.to_numeric(top_10_customers['AMOUNT REAL'], errors='coerce')
    top_10_customers['AMOUNT REAL'] = top_10_customers['AMOUNT REAL'].apply(lambda x: f"{x:,.0f}" if pd.notna(x) else '0')

    # ================== Pie Chart Week by QTY REAL ==================
    filtered_data_week = df.groupby("WEEK")["QTY REAL"].sum().reset_index()
    filtered_data_week["QTY REAL"] = filtered_data_week["QTY REAL"].astype(int)

    fig_week = px.pie(
        filtered_data_week,
        names="WEEK",
        values="QTY REAL",
        title="Sales by Week",
        template="plotly_white"
    )

    # Menyesuaikan ukuran font dan warna teks pada Pie Chart agar Bold
    fig_week.update_traces(
        textfont=dict(size=13, color="black", family="Arial, sans-serif", weight="bold"),
        textinfo="percent+label",
        marker=dict(colors=['#FF7F0E', '#1F77B4', '#2CA02C', '#D62728', '#9467BD']),
        showlegend=False  # Menonaktifkan legenda
    )

    # Mengatur ukuran Pie Chart dengan menambahkan 'height' dan 'autosize'
    fig_week.update_layout(
        height=600,  # Mengatur tinggi Pie Chart
        width=800,   # Mengatur lebar Pie Chart
        autosize=True  # Agar menyesuaikan dengan lebar container
    )

    # ================== Membagi Kolom untuk Tabel dan Pie Chart ==================
    col1, col2 = st.columns([2, 1])

    with col1:
        # Menampilkan tabel Top 25 Customer dengan Ranking rata tengah
        st.write("### Top 25 Customer")
        st.write(
            top_10_customers[['Ranking', 'NAMA DEPO','ID PELANGGAN', 'NAMA PELANGGAN', 'AMOUNT REAL','TOTAL LITER','QTY REAL']].set_index('Ranking')
        )

    with col2:
        # Menampilkan Pie Chart
        st.plotly_chart(fig_week, use_container_width=True)

except FileNotFoundError:
    st.error(f"File '{file_name}' tidak ditemukan.")
except ValueError as e:
    st.error(f"Error saat membaca data: {e}")
except Exception as e:
    st.error(f"Terjadi kesalahan: {e}")
