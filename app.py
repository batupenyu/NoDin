import os
import tempfile
import streamlit as st
from fill_nota_dinas import NotaDinasFiller

st.set_page_config(page_title="Nota Dinas Filler", layout="wide")
st.title("Pengisi Nota Dinas")
st.markdown("Isi form berikut untuk menghasilkan file Nota Dinas berdasarkan template.")

if "peserta" not in st.session_state:
    st.session_state.peserta = [{"nama": "", "nip": "", "jabatan": ""}]

st.header("Header")
col1, col2 = st.columns(2)
with col1:
    kepada = st.text_input("Kepada", value="Kepala Sekolah SMKN 1 Koba", key="kepada")
    dari = st.text_input("Dari", value="SMK Negeri 1 Koba", key="dari")
    nomor = st.text_input("Nomor (bagian tengah)", value="800.1.11.1/............/SMKN 1 Kb/Dindik/2026", placeholder="misal: 001", key="nomor")
    lampiran = st.text_input("Lampiran", value="-", key="lampiran")
with col2:
    tanggal = st.date_input("Tanggal", key="tanggal")
    perihal = st.text_input("Perihal", value="", key="perihal")

st.header("Isi Nota Dinas")
kegiatan = st.text_input("Kegiatan (mengikuti ...)", value="", key="kegiatan")
dasar_nomor = st.text_input("Dasar Pelaksanaan - Nomor", value="", placeholder="misal: 001", key="dasar_nomor")
maksud_tujuan = st.text_area("Maksud dan Tujuan", value="", height=100, key="maksud_tujuan")

st.subheader("Pelaksanaan")
pcol1, pcol2, pcol3 = st.columns(3)
with pcol1:
    pel_tanggal = st.date_input("Tanggal Pelaksanaan", key="pel_tanggal")
with pcol2:
    pel_jam = st.text_input("Jam", value="", key="pel_jam")
with pcol3:
    pel_tempat = st.text_input("Tempat", value="", key="pel_tempat")

kesimpulan = st.text_area("Kesimpulan", value="", height=100, key="kesimpulan")

# --- Kelola Peserta (dipindah ke bawah field Kesimpulan) ---
st.header("Kelola Peserta")
for idx, p in enumerate(st.session_state.peserta):
    nama_s = st.text_input("Nama", key=f"peserta_nama_{idx}", label_visibility="collapsed", placeholder="isi nama")
    nip_s = st.text_input("NIP", key=f"peserta_nip_{idx}", label_visibility="collapsed", placeholder="isi NIP")
    jabatan_s = st.text_input("Jabatan", key=f"peserta_jab_{idx}", label_visibility="collapsed", placeholder="isi jabatan")
    if st.button("Hapus", key=f"peserta_del_{idx}"):
        if len(st.session_state.peserta) > 1:
            st.session_state.peserta.pop(idx)
            st.rerun()

if st.button("Tambah Peserta"):
    st.session_state.peserta.append({"nama": "", "nip": "", "jabatan": ""})
    st.rerun()

submitted = st.button("Generate Nota Dinas", use_container_width=True)

if submitted:
    # Collect peserta data
    peserta_data = []
    for idx in range(len(st.session_state.peserta)):
        peserta_data.append({
            "nama": st.session_state.get(f"peserta_nama_{idx}", ""),
            "nip": st.session_state.get(f"peserta_nip_{idx}", ""),
            "jabatan": st.session_state.get(f"peserta_jab_{idx}", ""),
        })

    try:
        with tempfile.TemporaryDirectory() as tmpdir:
            out_path = os.path.join(tmpdir, "nota_dinas_isi.docx")
            filler = NotaDinasFiller(template_path="nota_dinas_template_fillable.docx")

            filler.fill_header(
                Kepada=kepada,
                Dari=dari,
                Nomor=nomor,
                Lampiran=lampiran,
                Tanggal=tanggal.strftime("%d/%m/%Y"),
                Perihal=perihal,
            )
            filler.fill_kegiatan(kegiatan)
            filler.fill_dasar_pelaksanaan(nomor=dasar_nomor)
            filler.fill_maksud_tujuan(maksud_tujuan)
            filler.fill_peserta(peserta_data)
            filler.fill_pelaksanaan(
                tanggal=pel_tanggal.strftime("%d/%m/%Y"),
                jam=pel_jam,
                tempat=pel_tempat,
            )
            filler.fill_kesimpulan(kesimpulan)
            filler.fill_tembusan([])
            filler.save(out_path)

            with open(out_path, "rb") as f:
                data = f.read()

            st.success("Nota Dinas berhasil di-generate!")
            st.download_button(
                label="Download Nota Dinas (.docx)",
                data=data,
                file_name="nota_dinas_isi.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            )
    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")
