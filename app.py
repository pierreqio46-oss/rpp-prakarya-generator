# 1. Pastikan library sudah diinstal di terminal sebelum menjalankan skrip ini
# (Jalankan: pip install google-genai python-docx di terminal)

import os
from google import genai
from docx import Document
from docx.shared import Mm, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH

# 2. Masukkan API Key Anda di bawah ini
API_KEY = "AQ.Ab8RN6ItIeuU_NQY1rhUaM5f-Sl7GBRH8xdFv-OTpV4jFcMNgQ"

def generate_rpp():
    # Input interaktif untuk pengguna
    mapel = input("1. Masukkan Mata Pelajaran: ")
    kelas_fase = input("2. Masukkan Kelas/Fase: ")
    elemen = input("3. Masukan elemen Eksplorasi dan Observasi, Desain/Perencanaan, Produksi, Evaluasi dan Refleksi: ")
    sekolah = input("4. masukan nama sekolah: ")
    penulis = input('5. masukan nama penulis modul: ')
    topik_bahasan = input("6. Masukkan Topik/Materi Pembelajaran: ")
    tujuan_pembelajaran = input("7. Masukan Tujuan Pembelajaran: ")
    waktu = input("8. Masukan berapa pertemuan (contoh: 4): ")
    kktp = input("9. Masukan KKTP: ")
    dimensi_profil_lulusan = input("10. masukan dimensi Profil Lulusan: ")
    praktik_pedagogis = input("11. masukan praktik pedagogis (misal: PBL, PJBL, dan lain-lain): ")
    lingkungan_pembelajaran = input("12. masukan lingkungan pembelajaran: ")
    kemitraan_pembelajaran = input("13. masukan kemitraan: ")
    pemanfaatan_digital = input("14. masukan Pemanfaatan Digital: ")
    persiapan_pembelajaran = input("15. Masukan persiapan guru: ")
    kepala_sekolah = input("masukan nama kepala sekolah: ")
    nip_kepala_sekolah = input("masukan Nip Kepala sekolah: ")
    nip_penulis = input("Masukan Nip Penulis: ")

    client = genai.Client(api_key=API_KEY)

    # Modifikasi Prompt agar menghasilkan seluruh pertemuan secara detail
    prompt_text = f"""
    Buatkan Rencana Pelaksanaan Pembelajaran (RPP) / Modul Ajar berbasis Pembelajaran Mendalam (Deep Learning) secara LENGKAP untuk {waktu} pertemuan. 
    Gunakan dokumen di tautan https://drive.google.com/drive/folders/1E8R2MOdiz7vpQzgLtky2wswr5hoc2-Z5?usp=sharing sebagai sumber referensi utama, gunakan pula buku buku elektronik dari kementrian pendidikan yang mendukung 
    Setiap 1 pertemuan terdiri dari 120 menit. Bagi waktu di setiap pertemuan agar sesuai dengan kegiatan awal, kegiatan inti, dan penutup.
    
    Detail Kelas:
    - Mata Pelajaran: {mapel}
    - Kelas/Fase: {kelas_fase}
    - Elemen: {elemen}
    - Sekolah: {sekolah}
    - Penulis: {penulis}
    - Topik/Pokok Bahasan: {topik_bahasan}
    - Tujuan Pembelajaran: {tujuan_pembelajaran}
    - Total Waktu Rencana: {waktu} Pertemuan
    - Kriteria Ketercapaian Pembelajaran: {kktp}
    - Dimensi Profil Lulusan: {dimensi_profil_lulusan}
    - Praktik Pedagogis: {praktik_pedagogis}
    - Lingkungan Pembelajaran: {lingkungan_pembelajaran}
    - Kemitraan Pembelajaran: {kemitraan_pembelajaran}
    - Pemanfaatan Digital: {pemanfaatan_digital}
    - Persiapan Pembelajaran: {persiapan_pembelajaran}
    
    Struktur RPP harus mengikuti susunan berikut:
    Judul Modul (Satu judul menarik yang merangkum seluruh pertemuan)
    1. Identitas RPP
       A. Identifikasi (Dimensi profil lulusan)
       B. Desain Pembelajaran
    
    2. Langkah Pembelajaran:
       PENTING: Tuliskan secara detail langkah pembelajaran untuk Pertemuan 1, Pertemuan 2, hingga Pertemuan ke-{waktu}. Jangan disingkat atau digabung! Setiap pertemuan wajib menjabarkan sub-topik yang berkesinambungan serta memuat:
       - Kegiatan Awal (15 menit: memuat Apersepsi, Motivasi, Asesmen Diagnostik, Tujuan Pembelajaran, Manfaat Pembelajaran)
       - Kegiatan Inti (90 menit: gunakan tahapan pembelajaran pada praktik pedagogis yang dipilih, perhatikan aspek Meaningful, Eksplorasi Mendalam, dan Diskusi/Kolaborasi)
       - Kegiatan Akhir (15 menit: berisi kesimpulan dan refleksi, Fokus pada aspek Joyful, Refleksi, dan Apresiasi)
    
    3. Asesmen/Penilaian Formatif (Sediakan instrumen penilaian atau refleksi yang spesifik untuk masing-masing pertemuan)
    4. Lembar Kerja Murid (LKM/LKPD) untuk tiap-tiap pertemuan secara terpisah.
    5. Asesmen Sumatif (Buat 20 soal pilihan ganda HOTS berbasis materi dari seluruh pertemuan beserta kunci jawabannya di bagian akhir modul).
    6. Referensi/Daftar Pustaka.
    
    Di akhir modul cantumkan tanda tangan format kiri-kanan:
    Sebelah kiri: Kepala Sekolah: {kepala_sekolah} (NIP: {nip_kepala_sekolah})
    Sebelah kanan: Penulis: {penulis} (NIP: {nip_penulis})
    """

    print(f"\n[Sistem] Sedang merancang RPP Pembelajaran Mendalam untuk {waktu} Pertemuan... Mohon tunggu.\n")

    try:
        response = client.models.generate_content(
            model='gemini-2.5-flash',
            contents=prompt_text,
        )
        
        # --- PROSES PEMBUATAN FILE DOCX (WORD) UKURAN A4 ---
        doc = Document()
        
        # Mengatur Ukuran Kertas ke A4 (210mm x 297mm)
        section = doc.sections[0]
        section.page_width = Mm(210)
        section.page_height = Mm(297)
        section.top_margin = Mm(25)
        section.bottom_margin = Mm(25)
        section.left_margin = Mm(25)
        section.right_margin = Mm(25)
        
        # Mengatur Gaya Tulisan Default (Arial, 11pt)
        style = doc.styles['Normal']
        font = style.font
        font.name = 'Arial'
        font.size = Pt(11)
        
        # Memasukkan hasil AI ke dalam dokumen Word baris demi baris
        rpp_text = response.text
        for line in rpp_text.split('\n'):
            if line.startswith('# '):
                p = doc.add_heading(line.replace('# ', ''), level=1)
            elif line.startswith('## '):
                p = doc.add_heading(line.replace('## ', ''), level=2)
            elif line.startswith('### '):
                p = doc.add_heading(line.replace('### ', ''), level=3)
            else:
                p = doc.add_paragraph(line)
                
        # Nama file output
        nama_file = f"RPP_{mapel.replace(' ', '_')}_{waktu}_Pertemuan.docx"
        doc.save(nama_file)
        
        print("--- PROSES SELESAI ---")
        print(f"[Sistem] Berhasil membuat dokumen Word resmi berukuran A4 berisi {waktu} pertemuan!")
        print(f"[Sistem] File tersimpan dengan nama: {nama_file}")
        print("[Sistem] Silakan unduh melalui ikon 'Folder (Files)' di sebelah kiri layar Google Colab Anda.")
        
    except Exception as e:
        print(f"Terjadi kesalahan: {e}")

# Menjalankan fungsi generator
generate_rpp()
