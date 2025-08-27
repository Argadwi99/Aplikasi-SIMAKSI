import streamlit as st
import docx
from docx import Document
import io
from datetime import date # Import 'date' untuk default tanggal

# Fungsi untuk mengisi template Word (tidak berubah)
def fill_word_template(template_path, data_dict):
    try:
        doc = Document(template_path)
        for p in doc.paragraphs:
            for key, value in data_dict.items():
                if key in p.text:
                    inline = p.runs
                    for i in range(len(inline)):
                        if key in inline[i].text:
                            text = inline[i].text.replace(key, value)
                            inline[i].text = text
        
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    for p in cell.paragraphs:
                        for key, value in data_dict.items():
                            if key in p.text:
                                inline = p.runs
                                for i in range(len(inline)):
                                    if key in inline[i].text:
                                        text = inline[i].text.replace(key, value)
                                        inline[i].text = text
        return doc
    except Exception as e:
        st.error(f"Gagal memproses template: {e}")
        return None

# --- Tampilan Aplikasi Web ---
st.title("Aplikasi Pembuat Dokumen SIMAKSI")
st.write("Silakan isi form di bawah ini untuk membuat dokumen secara otomatis.")

# Membuat form input
with st.form("simaksi_form"):
    nama_lengkap = st.text_input("Nama Lengkap:")
    tujuan_kegiatan = st.text_input("Tujuan Kegiatan:")
    lokasi_konservasi = st.text_input("Lokasi Kawasan Konservasi:")
    
    # --- PERUBAHAN DI SINI: Menggunakan st.date_input ---
    tanggal_mulai_obj = st.date_input("Tanggal Mulai:", value=date.today(), format="DD/MM/YYYY")
    tanggal_selesai_obj = st.date_input("Tanggal Selesai:", value=date.today(), format="DD/MM/YYYY")
    
    jumlah_peserta = st.text_input("Jumlah Peserta/Pengikut:")
    
    # Tombol submit
    submitted = st.form_submit_button("Buat Dokumen Sekarang")

# Jika tombol ditekan, proses dokumen
if submitted:
    # --- PERUBAHAN DI SINI: Mengubah format tanggal ---
    tanggal_mulai = tanggal_mulai_obj.strftime("%d/%m/%Y")
    tanggal_selesai = tanggal_selesai_obj.strftime("%d/%m/%Y")

    if not all([nama_lengkap, tujuan_kegiatan, lokasi_konservasi, tanggal_mulai, tanggal_selesai, jumlah_peserta]):
        st.warning("Harap isi semua kolom.")
    else:
        with st.spinner("Sedang membuat dokumen..."):
            # Siapkan data dari form
            data_to_fill = {
                '[nama_lengkap]': nama_lengkap,
                '[tujuan_kegiatan]': tujuan_kegiatan,
                '[lokasi_konservasi]': lokasi_konservasi,
                '[tanggal_mulai]': tanggal_mulai,
                '[tanggal_selesai]': tanggal_selesai,
                '[jumlah_peserta]': jumlah_peserta,
            }

            # Proses template
            template_name = "Draft SIMAKSI_Kosong.docx"
            filled_doc = fill_word_template(template_name, data_to_fill)

            if filled_doc:
                # Simpan dokumen ke memori virtual
                bio = io.BytesIO()
                filled_doc.save(bio)
                
                st.success("Dokumen berhasil dibuat!")
                
                # Buat tombol download
                st.download_button(
                    label="📥 Unduh Dokumen Word",
                    data=bio.getvalue(),
                    file_name=f"SIMAKSI_{nama_lengkap.replace(' ', '_')}.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )