import pandas as pd
import streamlit as st
import os

# Konfigurasi halaman Streamlit
st.set_page_config(page_title="Gabung & Filter Laporan 5G", page_icon="📊", layout="wide")

st.title("📊 Aplikasi Penggabung & Filter Laporan 5G & Traffic PLMN")
st.write("Silakan upload beberapa file Excel laporan Nokia Anda melalui panel di sebelah kiri.")

# --- SIDEBAR: KONTROL FILTER & UPLOAD FILE ---
st.sidebar.header("⚙️ Pengaturan & Data")

# Widget Upload File (Mendukung banyak file sekaligus)
uploaded_files = st.sidebar.file_uploader(
    "Upload File Excel (.xlsx)", 
    type=["xlsx"], 
    accept_multiple_files=True,
    help="Pilih satu atau beberapa file Excel laporan 5G sekaligus"
)

st.sidebar.divider()

# Input Site ID
filter_site_id = st.sidebar.text_input("Masukkan Site ID (Opsional)", value="09TJK0152", help="Kosongkan jika ingin mengambil semua Site ID")

# Input Rentang Tanggal
start_date = st.sidebar.date_input("Tanggal Awal", value=pd.to_datetime("2026-01-01"))
end_date = st.sidebar.date_input("Tanggal Akhir", value=pd.to_datetime("2026-12-31"))

# Tombol Eksekusi Utama
run_button = st.sidebar.button("🚀 Jalankan Proses Penggabungan", type="primary")

# --- PROSES UTAMA KETIKA TOMBOL DIKLIK ---
if run_button:
    if not uploaded_files:
        st.error("❌ Silakan upload minimal 1 file Excel terlebih dahulu melalui panel di sidebar!")
    else:
        with st.spinner("Sedang memproses file Excel, mohon tunggu..."):
            
            list_summary = []
            list_5g = []
            list_traffic = []

            for uploaded_file in uploaded_files:
                file_name = uploaded_file.name
                try:
                    # Membaca file Excel yang di-upload ke dalam memori
                    excel_file = pd.ExcelFile(uploaded_file)
                    
                    # 1. Proses Sheet 'Summary'
                    if 'Summary' in excel_file.sheet_names:
                        df_sum = pd.read_excel(excel_file, sheet_name='Summary')
                        if filter_site_id and 'Site ID' in df_sum.columns:
                            df_sum = df_sum[df_sum['Site ID'].astype(str).str.contains(filter_site_id, na=False)]
                        if not df_sum.empty:
                            list_summary.append(df_sum)

                    # 2. Proses Sheet '5G' (dengan skiprows=2 sesuai format template Nokia)
                    if '5G' in excel_file.sheet_names:
                        df_5g = pd.read_excel(excel_file, sheet_name='5G', skiprows=2)
                        cols = df_5g.columns.tolist()
                        site_col = cols[1] if len(cols) > 1 else None
                        date_col = cols[4] if len(cols) > 4 else None
                        
                        # Filter Site ID
                        if filter_site_id and site_col:
                            df_5g = df_5g[df_5g[site_col].astype(str).str.contains(filter_site_id, na=False)]
                        
                        # Filter Tanggal
                        if date_col:
                            df_5g[date_col] = pd.to_datetime(df_5g[date_col], errors='coerce')
                            df_5g = df_5g[(df_5g[date_col] >= pd.to_datetime(start_date)) & (df_5g[date_col] <= pd.to_datetime(end_date))]
                        
                        if not df_5g.empty:
                            df_5g['Sumber_File'] = file_name
                            list_5g.append(df_5g)
                    
                    # 3. Proses Sheet '5G_Traffic_PLMN'
                    if '5G_Traffic_PLMN' in excel_file.sheet_names:
                        df_traffic = pd.read_excel(excel_file, sheet_name='5G_Traffic_PLMN')
                        if 'Time' in df_traffic.columns:
                            df_traffic['Time'] = pd.to_datetime(df_traffic['Time'], errors='coerce')
                            df_traffic = df_traffic[(df_traffic['Time'] >= pd.to_datetime(start_date)) & (df_traffic['Time'] <= pd.to_datetime(end_date))]
                        
                        if not df_traffic.empty:
                            df_traffic['Sumber_File'] = file_name
                            list_traffic.append(df_traffic)
                            
                except Exception as e:
                    st.error(f"Gagal memproses file {file_name}: {e}")

            # Simpan hasil gabungan ke file Excel secara virtual di memori
            output_file_name = 'Hasil_Laporan_Filtered.xlsx'
            
            with pd.ExcelWriter(output_file_name, engine='openpyxl') as writer:
                if list_summary:
                    pd.concat(list_summary, ignore_index=True).to_excel(writer, sheet_name='Summary', index=False)
                if list_5g:
                    pd.concat(list_5g, ignore_index=True).to_excel(writer, sheet_name='5G', index=False)
                if list_traffic:
                    pd.concat(list_traffic, ignore_index=True).to_excel(writer, sheet_name='5G_Traffic_PLMN', index=False)

            st.success("🎉 Proses penggabungan dan filter selesai!")

            # --- TOMBOL DOWNLOAD & PREVIEW ---
            if os.path.exists(output_file_name):
                with open(output_file_name, "rb") as f:
                    st.download_button(
                        label="📥 Download File Hasil Excel",
                        data=f,
                        file_name=output_file_name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
                # Tampilkan Preview Data Sheet 5G jika ada
                if list_5g:
                    st.subheader("Preview Data Sheet '5G' (Gabungan)")
                    preview_df = pd.concat(list_5g, ignore_index=True)
                    st.dataframe(preview_df.head(50))
