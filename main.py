import pandas as pd
import streamlit as st
from io import BytesIO

def main():
    st.title("KODE POS")

    st.write("Pilih file Excel untuk data OSS:")
    uploaded_file_data1 = st.file_uploader("Upload File", type=["xlsx", "xls"])

    st.write("Pilih file Excel untuk ListKodePOS:")
    uploaded_file_listkodepos = st.file_uploader("Upload File", type=["xlsx", "xls"], key="listkodepos")

    if uploaded_file_data1 is not None and uploaded_file_listkodepos is not None:
        # Baca semua sheet dalam file Data1
        xls = pd.ExcelFile(uploaded_file_data1)
        sheets = []
        for sheet_name in xls.sheet_names:
            sheet = pd.read_excel(xls, sheet_name=sheet_name)
            sheets.append(sheet)

        # Gabungkan semua sheet menjadi satu DataFrame
        data1 = pd.concat(sheets, ignore_index=True)

        # Baca file ListKodePOS_filtered.xlsx
        list_kodepos = pd.read_excel(uploaded_file_listkodepos)

        # Lowercase semua kata
        data1['KELURAHAN_PERSEROAN'] = data1['KELURAHAN_PERSEROAN'].str.lower()
        list_kodepos['kelurahan'] = list_kodepos['kelurahan'].str.lower()

        # Menghitung jumlah kode pos unik untuk setiap kelurahan di listkodepos
        kodepos_count = list_kodepos.groupby('kelurahan').kode_pos.nunique().reset_index()
        kodepos_count.columns = ['kelurahan', 'kode_pos_count']

        # Menggabungkan listkodepos dengan kodepos_count untuk mengetahui kelurahan yang memiliki lebih dari satu kode pos
        list_kodepos = list_kodepos.merge(kodepos_count, on='kelurahan', how='left')

        # Mengambil kelurahan yang hanya memiliki satu kode pos
        single_kodepos = list_kodepos[list_kodepos.kode_pos_count == 1].drop(columns=['kode_pos_count'])

        # Pastikan single_kodepos hanya memiliki satu baris per kelurahan
        single_kodepos = single_kodepos.drop_duplicates(subset=['kelurahan'])

        # Menggabungkan data1 dengan single_kodepos untuk mendapatkan kode pos yang sesuai
        data_merged = data1.merge(single_kodepos, left_on='KELURAHAN_PERSEROAN', right_on='kelurahan', how='left')

        # Mengisi KODE_POS_PERSEROAN dengan kode pos yang sesuai
        data_merged['KODE_POS_PERSEROAN'] = data_merged['kode_pos']

        # Menghapus kolom tambahan
        data_final = data_merged.drop(columns=['kelurahan', 'kode_pos'])

        # Memisahkan data yang KODE_POS_PERSEROAN null dan tidak null
        data_null = data_final[data_final['KODE_POS_PERSEROAN'].isna()]
        data_not_null = data_final[data_final['KODE_POS_PERSEROAN'].notna()]

        st.write(f"Total records: {len(data_final)}")
        st.write(f"Not Null: {len(data_not_null)}")
        st.write(f"Null: {len(data_null)}")

        st.write("Sample Data:")
        st.write(data_final.head())

        # Tombol untuk mengunduh file hasil
        st.write("Download hasil:")
        output_null = to_excel(data_null)
        output_not_null = to_excel(data_not_null)

        st.download_button(
            label="Download file hasil (Kode Pos Belum Terisi)",
            data=output_null,
            file_name="hasil_oss_belum_terisi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.download_button(
            label="Download file hasil (Kode Pos Terisi)",
            data=output_not_null,
            file_name="hasil_oss_terisi.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

def to_excel(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sheet1')
    processed_data = output.getvalue()
    return processed_data

if __name__ == '__main__':
    main()
