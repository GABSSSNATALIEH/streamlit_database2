import pickle
import streamlit as st
import numpy as np

# Membaca model scaler yang telah disimpan
scaler_iq = pickle.load(open('scaler_iq.sav', 'rb'))

# Judul aplikasi
st.title("Test TOEFL - Analisis Nilai IQ")

# Input untuk Skor Mentah
skor_mentah = st.text_input("Masukkan Skor Mentah", value="0")

# Keterangan hasil prediksi
keterangan_iq = ''

# Tombol untuk memproses data
if st.button("Hitung Nilai IQ"):
    try:
        # Konversi skor mentah ke tipe float
        skor_mentah_float = float(skor_mentah)

        # Standarisasi skor mentah
        skor_mentah_standar = scaler_iq.transform([[skor_mentah_float]])[0][0]

        # Hitung nilai IQ
        nilai_iq = (skor_mentah_standar * 15) + 100

        # Tentukan kategori IQ
        if nilai_iq >= 110:
            keterangan_iq = "Di Atas Rata-Rata"
            outcome = 3
        elif nilai_iq >= 92:
            keterangan_iq = "Rata-Rata"
            outcome = 2
        elif nilai_iq >= 56:
            keterangan_iq = "Di Bawah Rata-Rata"
            outcome = 1
        else:
            keterangan_iq = "Defisiensi"
            outcome = 0

        # Tampilkan hasil analisis
        st.success(f"Nilai IQ Anda: {nilai_iq:.2f}")
        st.info(f"Keterangan: {keterangan_iq} (Outcome: {outcome})")

    except ValueError:
        st.error("Pastikan input Skor Mentah berupa angka yang valid.")