import streamlit as st
import pandas as pd
import plotly.express as px
import os

# 📂 Membaca file Excel
file_name = "mst_sales.xlsx"
df = pd.read_excel(file_name, parse_dates=["TANGGAL"])  # Pastikan kolom TANGGAL terbaca sebagai datetime

# 📂 Membaca dataset target (misal target_channel.xlsx)
target_channel  = pd.read_excel("target_channel.xlsx")
target_branch   = pd.read_excel("target_branch.xlsx")
target_sku      = pd.read_excel("target_sku.xlsx")

# 🗓️ Ekstrak MONTH dan YEAR dari kolom TANGGAL
latest_date = df["TANGGAL"].max()  # Ambil tanggal terbaru dalam dataset
month_year = latest_date.strftime("%b %Y")  # Format menjadi "Dec 2024"

# 🏷️ Set Title dengan nilai dari data
st.set_page_config(page_title=f"Sales Trend - AQUA {month_year}", page_icon=":bar_chart:", layout="wide")

# Memuat file CSS eksternal
with open("assets/style.css") as f:
    st.markdown(f"<style>{f.read()}</style>", unsafe_allow_html=True)

# 🏷️ Gunakan month_year dalam st.title
st.markdown(f"""
    <h1 style="color: #010169; text-align: center;">
        Sales Trend - AQUA Division {month_year}
    </h1>
    """, unsafe_allow_html=True)

# Pastikan nama file sesuai dataset lokal Anda
# file_name = "mst_sales.xlsx"

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
    # Sum liter per tanggal dulu, kemudian ambil rata-ratanya
    # Jumlahkan liter per tanggal terlebih dahulu
    total_liter_per_day = df.groupby('TANGGAL')['TOTAL LITER'].sum().reset_index()
    # Ambil rata-rata dari total liter per tanggal
    average_sales_by_liter = round(total_liter_per_day['TOTAL LITER'].mean(), 2)
    # Menggunakan format string untuk menambahkan koma sebagai pemisah ribuan
    formatted_average_sales_by_liter = f"{average_sales_by_liter:,.2f}"

    # Kolom untuk meratakan angka
    left_column, middle_column, right_column = st.columns(3)

    # Menampilkan total LITER, Average Sales per Day, dan Total Active Outlet
    with left_column:
        st.markdown("<div class='center-text' style='color: #1C79F2'>Total Liter</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='center-text'>{total_sales:,}</div>", unsafe_allow_html=True)
    with middle_column:
        st.markdown("<div class='center-text' style='color: #1C79F2'>Average Sales per Day (Liter)</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='center-text'>{formatted_average_sales_by_liter}</div>", unsafe_allow_html=True)
    with right_column:
        st.markdown("<div class='center-text' style='color: #1C79F2'>Total Active Outlet</div>", unsafe_allow_html=True)
        st.markdown(f"<div class='center-text'>{ao_customer:,}</div>", unsafe_allow_html=True)

    st.markdown("""---""")

    # Menambahkan styling pada teks "Filter Total"
    st.markdown('<div style="font-size:15px; font-weight:bold; color:#A31B17;">Filter Total:</div>', unsafe_allow_html=True)

    # Dropdown untuk memilih kolom filter yang digunakan pada semua chart
    filter_column = st.selectbox(
        "",
        ["by Amount", "by Quantity", "by Liter"]
    )
    # Menerjemahkan pilihan dropdown ke nama kolom yang sesuai
    column_map = {
        "by Amount": "AMOUNT REAL",
        "by Quantity": "QTY REAL",
        "by Liter": "TOTAL LITER"
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
        st.markdown(f"<h3 style='text-align: center; color: #1C79F2'>SPD by Branch {filter_column}</h3>", unsafe_allow_html=True)
        sales_per_day_depo_pivot = sales_per_day_depo_pivot.applymap(lambda x: f"{int(x):,}" if isinstance(x, (int, float)) else x)
        st.dataframe(sales_per_day_depo_pivot, use_container_width=True)

    with col2:
        st.markdown(f"<h3 style='text-align: center; color: #1C79F2'>SPD by SKU {filter_column}</h3>", unsafe_allow_html=True)
        sales_per_day_sku_pivot = sales_per_day_sku_pivot.applymap(lambda x: f"{int(x):,}" if isinstance(x, (int, float)) else x)
        st.dataframe(sales_per_day_sku_pivot, use_container_width=True)

    st.markdown("""---""")

        # ================== Chart 1: By Depo ==================
    df_target_branch = pd.merge(
        df.groupby("NAMA DEPO")[filter_column_name].sum().reset_index(),
        target_branch,
        on="NAMA DEPO",
        how="left"
    )


        # Kondisi untuk memilih kolom target sesuai filter
    if filter_column == "by Quantity":
        target_col = "TARGET DEPO QTY"
        y_columns = [filter_column_name, target_col]
        title_chart = "Perbandingan Penjualan per Branch (Quantity)"
    elif filter_column == "by Liter":
        target_col = "TARGET DEPO LITER"
        y_columns = [filter_column_name, target_col]
        title_chart = "Perbandingan Penjualan per Branch (Liter)"
    else:
        y_columns = [filter_column_name]
        target_col = None
        title_chart = "Perbandingan Penjualan per Branch"

    fig_depo_comparison = px.bar(
        df_target_branch,
        x="NAMA DEPO",
        y=y_columns,
        barmode="group",
        title=title_chart,
        text_auto=True,
        labels={
            "NAMA DEPO": "Branch",
            filter_column_name: f"Total {filter_column_name}",
            target_col: target_col if target_col else ""
        },
        template="plotly_white"
    )

    fig_depo_comparison.update_layout(
        yaxis_title=None,
        yaxis_tickformat=",.0f",
        title_font=dict(size=20, color="#1C79F2"),  # Ukuran font diperbesar
    )

    st.plotly_chart(fig_depo_comparison, use_container_width=True)

    # ================== Chart 2: By Channel (Bar Chart) ==================
    # Menggabungkan data penjualan dengan target berdasarkan kolom CHANNEL
    df_target_channel = pd.merge(
        df.groupby("CHANNEL")[filter_column_name].sum().reset_index(),
        target_channel,
        on="CHANNEL",
        how="left"
    )

    # Kondisi untuk memilih kolom target sesuai filter
    if filter_column == "by Quantity":
        target_col = "TARGET CH QTY"
        y_columns = [filter_column_name, target_col]
        title_chart = "Perbandingan Penjualan per Channel (Quantity)"
    elif filter_column == "by Liter":
        target_col = "TARGET CH LITER"
        y_columns = [filter_column_name, target_col]
        title_chart = "Perbandingan Penjualan per Channel (Liter)"
    else:
        y_columns = [filter_column_name]
        target_col = None
        title_chart = "Perbandingan Penjualan per Channel"

    # Membuat chart dengan kondisi yang ditentukan
    fig_channel_comparison = px.bar(
        df_target_channel,
        x="CHANNEL",
        y=y_columns,
        barmode="group",
        title=title_chart,
        text_auto=True,
        labels={
            "CHANNEL": "Channel",
            filter_column_name: f"Total {filter_column_name}",
            target_col: target_col if target_col else ""
        },
        template="plotly_white"
    )

    fig_channel_comparison.update_layout(
        yaxis_title=None,
        yaxis_tickformat=",.0f",
        title_font=dict(size=20, color="#1C79F2"),  # Ukuran font diperbesar
    )

    st.plotly_chart(fig_channel_comparison, use_container_width=True)

    # ================== Chart 3: By SKU Report ==================
    df_target_sku = pd.merge(
        df.groupby("SKU REPORT")[filter_column_name].sum().reset_index(),
        target_sku,
        on="SKU REPORT",
        how="left"
    )

        # Kondisi untuk memilih kolom target sesuai filter
    if filter_column == "by Quantity":
        target_col = "TARGET SKU QTY"
        y_columns = [filter_column_name, target_col]
        title_chart = "Perbandingan Penjualan per SKU (Quantity)"
    elif filter_column == "by Liter":
        target_col = "TARGET SKU LITER"
        y_columns = [filter_column_name, target_col]
        title_chart = "Perbandingan Penjualan per SKU (Liter)"
    else:
        y_columns = [filter_column_name]
        target_col = None
        title_chart = "Perbandingan Penjualan per SKU"

    fig_sku_comparison = px.bar(
        df_target_sku,
        x="SKU REPORT",
        y=y_columns,
        barmode="group",
        title=title_chart,
        text_auto=True,
        labels={
            "SKU REPORT": "SKU",
            filter_column_name: f"Total {filter_column_name}",
            target_col: target_col if target_col else ""
        },
        template="plotly_white"
    )

    fig_sku_comparison.update_layout(
        yaxis_title=None,
        yaxis_tickformat=",.0f",
        title_font=dict(size=20, color="#1C79F2"),  # Ukuran font diperbesar
    )

    st.plotly_chart(fig_sku_comparison, use_container_width=True)

    # ================== Tabel Top 10 NAMA PELANGGAN ==================
    # Mengelompokkan data untuk mendapatkan total AMOUNT REAL, TOTAL LITER, dan QTY REAL per pelanggan
    customer_sales = df.groupby(["NAMA DEPO", "ID PELANGGAN", "NAMA PELANGGAN"])[["AMOUNT REAL", "TOTAL LITER", "QTY REAL"]].sum().reset_index()
    # Mengurutkan berdasarkan AMOUNT REAL dan mengambil top 25
    top_10_customers = customer_sales.sort_values(by="AMOUNT REAL", ascending=False).head(25)
    # Menambahkan kolom NO (Ranking)
    top_10_customers['NO'] = top_10_customers.reset_index().index + 1
    # Mengonversi AMOUNT REAL menjadi integer
    top_10_customers['AMOUNT REAL'] = top_10_customers['AMOUNT REAL'].astype(int)
    # Menambahkan koma sebagai pemisah ribuan untuk AMOUNT REAL
    top_10_customers['AMOUNT REAL'] = top_10_customers['AMOUNT REAL'].apply(lambda x: f"{int(x):,}")  # Mengubah menjadi integer dan format ribuan
    top_10_customers['TOTAL LITER'] = top_10_customers['TOTAL LITER'].apply(lambda x: f"{int(x):,}")
    top_10_customers['QTY REAL'] = top_10_customers['QTY REAL'].apply(lambda x: f"{int(x):,}")

    # ================== Pie Chart Week by QTY REAL ==================
    filtered_data_week = df.groupby("WEEK")["QTY REAL"].sum().reset_index()
    filtered_data_week["QTY REAL"] = filtered_data_week["QTY REAL"].astype(int)

    fig_week = px.pie(
        filtered_data_week,
        names="WEEK",
        values="QTY REAL",
        template="plotly_white"
    )

    # Menyesuaikan ukuran font dan warna teks pada Pie Chart agar Bold
    fig_week.update_traces(
        textfont=dict(size=13, color="black", family="Arial, sans-serif", weight="bold"),
        textinfo="percent+label",
        marker=dict(colors=['#A4DDED', '#6B9AC4', '#F1C40F', '#2ECC71', '#E74C3C']),
        showlegend=False  # Menonaktifkan legenda
    )

    # Mengatur ukuran Pie Chart dengan menambahkan 'height' dan 'autosize'
    fig_week.update_layout(
        height=400,  # Mengatur tinggi Pie Chart
        width=800,   # Mengatur lebar Pie Chart
        autosize=True,  # Agar menyesuaikan dengan lebar container
        margin=dict(t=7)  # Mengurangi margin atas agar lebih dekat dengan judul
    )

    # ================== Membagi Kolom untuk Tabel dan Pie Chart ==================
    col1, col2 = st.columns([2, 1])

    with col1:
        # Menampilkan label dengan penyesuaian jarak
        st.markdown('<h3 style="color:#1C79F2; margin-bottom: 0;">Top 25 Customer</h3>', unsafe_allow_html=True)
        # ✅ Menampilkan tabel dengan Ranking sebagai index
        renamed_df = top_10_customers.rename(columns={
            'AMOUNT REAL': 'AMOUNT',
            'TOTAL LITER': 'LITER',
            'QTY REAL': 'QUANTITY'
        })

        st.write(renamed_df[['NO', 'NAMA DEPO', 'NAMA PELANGGAN', 'AMOUNT', 'LITER', 'QUANTITY']].set_index('NO'))

    with col2:
        # Menampilkan label dengan penyesuaian jarak
        st.markdown('<h3 style="color:#1C79F2; margin-bottom: 0;">Sales by Week</h3>', unsafe_allow_html=True)
        # Menampilkan Pie Chart
        st.plotly_chart(fig_week, use_container_width=True)

except FileNotFoundError:
    st.error(f"File '{file_name}' tidak ditemukan.")
except ValueError as e:
    st.error(f"Error saat membaca data: {e}")
except Exception as e:
    st.error(f"Terjadi kesalahan: {e}")
