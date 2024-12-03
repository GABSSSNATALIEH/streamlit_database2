import numpy as np
import pandas as pd
import pickle
from sklearn.preprocessing import StandardScaler

# Load data dari file Excel
file_excel = 'data_iq.xlsx'
try:
    data = pd.read_excel(file_excel)
    print(f"Data berhasil dimuat dari {file_excel}")
except FileNotFoundError:
    print(f"File {file_excel} tidak ditemukan. Pastikan file ada di lokasi yang benar.")
    exit()

# Periksa nama kolom
print("Nama Kolom Awal:", data.columns)

# Bersihkan nama kolom jika ada spasi
data.columns = data.columns.str.strip()

# Isi ulang nilai NaN hanya pada kolom numerik
numeric_cols = data.select_dtypes(include=[np.number]).columns
if not numeric_cols.empty:
    data[numeric_cols] = data[numeric_cols].fillna(data[numeric_cols].mean())

# Pastikan kolom 'Skor Mentah' ada dan gunakan nama kolom yang sesuai
if 'Skor Mentah' in data.columns:
    try:
        # Menstandarisasi skor mentah
        scaler = StandardScaler()
        data['Skor_Mentah_Standar'] = scaler.fit_transform(data[['Skor Mentah']])

        # Hitung Nilai IQ
        data['Nilai_IQ'] = (data['Skor_Mentah_Standar'] * 15) + 100

        # Tambahkan kolom keterangan berdasarkan nilai IQ
        def keterangan_iq(iq):
            if iq >= 110:
                return 'Di Atas Rata-Rata'
            elif iq >= 92:
                return 'Rata-Rata'
            elif iq >= 56:
                return 'Di Bawah Rata-Rata'
            else:
                return 'Defisiensi'

        data['Keterangan'] = data['Nilai_IQ'].apply(keterangan_iq)

        # Tentukan Outcome berdasarkan nilai IQ
        data['Outcome'] = data['Nilai_IQ'] >= 100

        # Simpan hasil ke file Excel baru
        output_file = 'hasil_iq.xlsx'
        data.to_excel(output_file, index=False)
        print(f"Proses selesai. Data hasil disimpan ke {output_file}")

        # Simpan model standarisasi ke file pickle
        pickle.dump(scaler, open('scaler_iq.sav', 'wb'))
        print("Model standarisasi disimpan ke scaler_iq.sav")

    except Exception as e:
        print(f"Terjadi kesalahan saat memproses data: {e}")
else:
    print("Kolom 'Skor Mentah' tidak ditemukan dalam data. Periksa file Excel Anda.")
