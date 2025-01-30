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

    # Dropdown untuk memilih kolom filter untuk Depo
    filter_column = st.selectbox(
        "Pilih kolom data untuk ditampilkan per Depo:",
        ["AMOUNT REAL", "QTY REAL", "TOTAL LITER"]
    )

    # Validasi kolom yang dipilih
    if filter_column in df.columns:
        # Agregasi data sesuai dengan kolom yang dipilih
        filtered_data = df.groupby("NAMA DEPO")[filter_column].sum().reset_index()

        # Membulatkan hasil ke angka bulat (tanpa desimal)
        filtered_data[filter_column] = filtered_data[filter_column].round(0).astype(int)

        # Membuat Bar Chart interaktif dengan Plotly
        fig_depo = px.bar(
            filtered_data,
            x="NAMA DEPO",
            y=filter_column,
            title=f"{filter_column} by Depo",
            text_auto=True,
            color=filter_column,
            labels={"NAMA DEPO": "Nama Depo", filter_column: f"Total {filter_column}"},
            template="simple_white"  # Menggunakan template sederhana
        )

        # Update untuk memperbesar ukuran bar, angka, dan sumbu
        fig_depo.update_traces(textposition="outside", textfont_size=14)  # Angka di luar bar
        fig_depo.update_layout(
            xaxis_title="Nama Depo",
            yaxis_title=f"Total {filter_column}",
            xaxis_tickangle=-45,
            yaxis=dict(tickformat=",", tickfont_size=14),  # Perbesar angka Y
            xaxis=dict(tickfont_size=14),  # Perbesar angka X
            title=dict(font_size=24),  # Perbesar ukuran judul
            font=dict(size=16),  # Perbesar ukuran font seluruhnya
        )

        # Menampilkan Plotly Chart untuk Depo di Streamlit
        st.plotly_chart(fig_depo, use_container_width=True)  # Menyesuaikan lebar chart
    else:
        st.warning(f"Kolom '{filter_column}' tidak ditemukan.")

    # Menghitung total penjualan per NAMA PELANGGAN
    customer_sales = df.groupby("NAMA PELANGGAN")[["AMOUNT REAL"]].sum().reset_index()

    # Mengurutkan berdasarkan penjualan terbesar dan mengambil top 10
    top_10_customers = customer_sales.sort_values(by="AMOUNT REAL", ascending=False).head(10)

    # Menambahkan kolom Ranking
    top_10_customers['Ranking'] = top_10_customers.reset_index().index + 1

    # Membulatkan AMOUNT REAL ke angka bulat (tanpa desimal) untuk tabel
    top_10_customers['AMOUNT REAL'] = top_10_customers['AMOUNT REAL'].round(0).astype(int)

    # Menampilkan tabel Top 10 NAMA PELANGGAN
    st.write("### Top 10 NAMA PELANGGAN berdasarkan Penjualan")
    st.write(top_10_customers[['Ranking', 'NAMA PELANGGAN', 'AMOUNT REAL']].set_index('Ranking'))

except FileNotFoundError:
    st.error(f"File '{file_name}' tidak ditemukan.")
except ValueError as e:
    st.error(f"Error saat membaca data: {e}")
