import streamlit as st
import pandas as pd
import io
import os
import json
from google import genai
from google.genai import types
import openpyxl
from io import BytesIO # Import BytesIO for memory handling
import time
import hashlib
import re
import html  # Diperlukan untuk html.escape()


# --- 1. CONFIGURATION AND INITIALIZATION ---

st.set_page_config(
    page_title="AI Strategic Insight Generator",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("📊 AI Measurement Insight Midi")
st.markdown("Pilih Divisi dan Cluster untuk membandingkan program kerja *existing* dengan *insight* tren industri dari AI.")

# --- 2. GEMINI API KEY SETUP ---
# WARNING: Storing the API key directly in code is insecure.
api_key =st.secrets['MY_API_KEY'] 

try:
    client = genai.Client(api_key=api_key)
except Exception as e:
    st.error(f"Error initializing Gemini client: {e}")
    st.stop()


# --- 3. Logic : CALL GEMINI (Fungsi 1 - MODE TEKS) ---
# (Fungsi ini tidak berubah)
def get_gemini_strategic_insight(divisi: str, cluster: str, cluster_definition: str, model_name: str) -> str:
    """
    PANGGILAN API #1:
    Melakukan panggilan API ke Gemini dalam MODE TEKS untuk menghasilkan daftar insight strategis.
    
    Returns:
        (str): Hasil teks mentah (atau None jika error).
    """
    
    # (PROMPT DARI PERCAKAPAN SEBELUMNYA SUDAH BENAR)
    system_prompt = f"""
    Anda adalah AI yang bertugas membuat daftar program kerja (job programs) dan deskripsi yang relevan untuk Divisi '{divisi}' dengan Cluster '{cluster}' menurut insight anda sendiri yang dicari dari internet dan menentukan apakah
    itu termasuk dalam OKR tipe Objectives (Objectives (tujuan): Apa yang ingin dicapai? Objectives harus jelas, ringkas, dan inspiratif) atau Key Results(Key Results (hasil kunci): Bagaimana cara mencapai tujuan tersebut? Key results adalah indikator terukur yang menunjukkan seberapa dekat kita dengan pencapaian objective).

    DEFINISI CLUSTER SAAT INI (KONTEKS):
    '{cluster_definition}'

    INSTRUKSI SANGAT KETAT (HARUS DIIKUTI):
    1.  Output WAJIB dalam bahasa Indonesia.
    2.  Buat beberapa program kerja baru yang relevan dengan cluster ini, FOKUS pada tren industri ritel modern.
    3.  JANGAN PERNAH menulis paragraf pembuka, sapaan, atau penjelasan (Contoh: "Sebagai seorang ahli...", "Berikut adalah...", dsb.).
    4.  JANGAN PERNAH menulis kesimpulan atau ringkasan di bagian akhir.
    5.  Output Anda HARUS dan HANYA BOLEH berisi daftar berpoin (bullet points), dimulai dengan tanda `-` atau `*`.
    6.  Setiap poin WAJIB mengikuti format `Program:`, `Deskripsi:`, dan `OKR:` persis seperti contoh di bawah.
    7.  Di bagian PALING AKHIR, setelah semua daftar program, tambahkan bagian `Sumber:` dan berikan beberapa sitasi terpercaya (situs resmi seperti Amazon dll., jurnal pendidikan, laporan industri) dengan link yang bisa diakses. JANGAN gunakan blog pribadi.

    BENTUK OUTPUT (WAJIB DIIKUTI PERSIS):

    - Program: [Nama Program Kerja 1 Sesuai Tren]
      Deskripsi: [Deskripsi singkat untuk program 1 yang relevan dengan tren]
      OKR : [Tentukan apakah program ini termasuk Objectives atau Key Results. JAWAB HANYA DENGAN `Objectives` atau `Key Results`]

    - Program: [Nama Program Kerja 2 Sesuai Tren]
      Deskripsi: [Deskripsi singkat untuk program 2 yang relevan dengan tren]
      OKR : [Tentukan apakah program ini termasuk Objectives atau Key Results. JAWAB HANYA DENGAN `Objectives` atau `Key Results`]

    - Program: [Nama Program Kerja 3 Sesuai Tren]
      Deskripsi: [Deskripsi singkat untuk program 3 yang relevan dengan tren]
      OKR : [Tentukan apakah program ini termasuk Objectives atau Key Results. JAWAB HANYA DENGAN `Objectives` atau `Key Results`]

    Sumber:
    - [Nama Jurnal/Laporan/Situs Resmi] - Judul - https://contoh.url/
    - [Nama Jurnal/Laporan/Situs Resmi] - Judul - https://contoh.url/
    """
    
    config = types.GenerateContentConfig(
        system_instruction=system_prompt
    )
    
    user_prompt = f"Berikan insight AI untuk Divisi: {divisi}, Cluster: {cluster}"

    try:
        response = client.models.generate_content(
            model=model_name,
            contents=user_prompt,
            config=config 
        )

        result = getattr(response, "text", None)
        if not result:
            candidates = getattr(response, "candidates", None)
            if candidates:
                parts = []
                for c in candidates:
                    val = getattr(c, "content", None) or getattr(c, "output", None) or getattr(c, "text", None)
                    if val:
                        parts.append(str(val))
                    else:
                        parts.append(str(c))
                result = "\n".join(parts)
            else:
                result = str(response)

        if isinstance(result, str):
            return result.strip()
        else:
            return str(result).strip()

    except Exception as e:
        st.error(f"Error saat memanggil Gemini API (Call 1 - Text Mode): {e}")
        return None

# (Fungsi to_excel tidak berubah)
def to_excel(df):
    """
    Mengkonversi DataFrame menjadi file Excel di dalam memori (bytes).
    """
    output = BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Perbandingan_Strategi')
        for column in df:
            column_length = max(df[column].astype(str).map(len).max(), len(column)) + 2
            col_idx = df.columns.get_loc(column)
            writer.sheets['Perbandingan_Strategi'].column_dimensions[openpyxl.utils.get_column_letter(col_idx + 1)].width = column_length
            
    processed_data = output.getvalue()
    return processed_data


# --- 4. STREAMLIT APP LAYOUT ---

# (Sidebar tidak berubah)
with st.sidebar:
    st.header("Opsi Analisis")
    use_caching = st.checkbox(
        "Gunakan Cache Hasil Analisis", 
        value=True,
        help="Jika diaktifkan, hasil analisis sebelumnya akan disimpan. Matikan fitur ini jika Anda ingin AI menganalisis ulang."
    )
    st.markdown("---")
    st.info("Aplikasi ini memanggil 1 API Call per analisis. Fitur 'Gunakan Cache' sangat disarankan untuk menghemat kuota API.")

# (Logika Upload & Parsing Excel tidak berubah)
uploaded_file = st.file_uploader(
    "Upload your Excel (.xlsx) or CSV file",
    type=['xlsx', 'csv'],
    accept_multiple_files=False
)

if 'cluster_dict' not in st.session_state:
    st.session_state.cluster_dict = {}

if uploaded_file is not None:
    file_extension = uploaded_file.name.split('.')[-1]
    
    if file_extension == 'xlsx':
        try:
            SHEET_TO_DIVISION_MAP = {
                "information technology": "it", 
                "corporate legal & compliance": "corporate legal & compliance",
                "operation": "operation",
                "merchandising": "merchandising",
                "marketing": "marketing",
                "business controlling": "business controlling",
                "service quality": "service quality",
                "property development": "property development",
                "corporate audit": "corporate audit",
                "finance": "finance",
                "human capital": "human capital"
            }
            
            excel_data = uploaded_file.read()
            xls = pd.ExcelFile(BytesIO(excel_data))
            sheet_names = xls.sheet_names
            
            cluster_sheet_name = None
            for name in sheet_names:
                if name.strip().lower() == 'cluster':
                    cluster_sheet_name = name
                    break
            
            sheet_names_to_select = [name for name in sheet_names if name.strip().lower() != 'cluster']
            
            selected_sheet = st.selectbox(
                "Pilih Divisi (Sheet) yang Akan Diproses:",
                sheet_names_to_select,
                index=0,
                key='sheet_selector'
            )
            
            df = None
            
            if selected_sheet:
                df = pd.read_excel(BytesIO(excel_data), sheet_name=selected_sheet)
                
                if cluster_sheet_name:
                    try:
                        df_cluster_def = pd.read_excel(BytesIO(excel_data), sheet_name=cluster_sheet_name, header=None)
                        temp_cluster_dict = {}
                        start_index = -1
                        division_target = SHEET_TO_DIVISION_MAP.get(selected_sheet.strip().lower(), selected_sheet.strip().lower())
                        all_division_names_lower = [v.lower() for v in SHEET_TO_DIVISION_MAP.values()]

                        for index, row in df_cluster_def.iterrows():
                            col_a_raw = str(row.iloc[0]).strip() if len(row) > 0 else ""
                            if col_a_raw.lower() == division_target:
                                start_index = index
                                break
                        
                        if start_index != -1:
                            for index in range(start_index + 1, len(df_cluster_def)):
                                row = df_cluster_def.iloc[index]
                                col_a_raw = str(row.iloc[0]).strip() if len(row) > 0 and not pd.isna(row.iloc[0]) else ""
                                col_b_raw = str(row.iloc[1]).strip() if len(row) > 1 and not pd.isna(row.iloc[1]) else ""
                                col_a_lower = col_a_raw.lower()
                                is_empty_row = not col_a_raw and not col_b_raw
                                is_other_division_header = col_a_lower in all_division_names_lower and col_a_lower != division_target and (not col_b_raw or col_b_raw.lower() == 'desc')
                                if is_empty_row or is_other_division_header:
                                    break
                                if col_a_lower not in ['', 'cluster', 'nan'] and col_b_raw.lower() not in ['', 'desc', 'nan']:
                                    temp_cluster_dict[col_a_raw] = col_b_raw
                        st.session_state.cluster_dict = temp_cluster_dict
                    except Exception as e:
                        st.error(f"Gagal mem-parsing sheet '{cluster_sheet_name}'. Error: {e}")
                        st.session_state.cluster_dict = {}
                
                if 'Cluster' in df.columns:
                    all_clusters_unique = df['Cluster'].dropna().astype(str).str.strip().unique().tolist()
                    all_clusters_in_sheet = [c for c in all_clusters_unique if c.lower() not in ['cluster', 'nan']]
                    available_clusters = []
                    if selected_sheet.strip().lower() == "operation":
                        logistics_only_clusters = [
                            "Inventory & Stock Management",
                            "Supplier & Service Level",
                            "Warehouse & Project Execution",
                            "System Development"
                        ]
                        available_clusters = sorted(
                            [c for c in all_clusters_in_sheet if c not in logistics_only_clusters]
                        )
                    else:
                        available_clusters = sorted(all_clusters_in_sheet)
                    
                    if not available_clusters:
                        st.warning(f"Sheet '{selected_sheet}' tidak memiliki data valid di kolom 'Cluster'.")
                        selected_cluster = None 
                    else:
                        selected_cluster = st.selectbox(
                            "Pilih Cluster yang Akan Dianalisis:",
                            available_clusters,
                            index=0,
                            key=f'cluster_selector_{selected_sheet}' 
                        )
                    
                    if selected_cluster:
                        cluster_definition = st.session_state.cluster_dict.get(selected_cluster.strip(), "(Definisi tidak ditemukan di Sheet CLUSTER)")
                        
                        if "Definisi tidak ditemukan" not in cluster_definition:
                            st.info(f"**Definisi Cluster:** {cluster_definition}")
                        else:
                            st.error(f"**Definisi Cluster:** {cluster_definition}")

                        disable_button = "Definisi tidak ditemukan" in cluster_definition
                        
                        if st.button(f"🚀 Generate Insight untuk Cluster '{selected_cluster}'", use_container_width=True, disabled=disable_button):
                            
                            # --- 1. AMBIL DAN PROSES DATA KIRI (EXISTING) ---
                            existing_df = df[df['Cluster'].astype(str).str.strip() == selected_cluster][['Program Kerja', 'Deskripsi']]
                            existing_markdown_items = []
                            existing_data_list = [] 
                            
                            if not existing_df.empty:
                                for index, row in existing_df.iterrows():
                                    program = row['Program Kerja']
                                    deskripsi = row['Deskripsi']
                                    item_string = "" 
                                    if pd.notna(program) and str(program).strip():
                                        item_string += f"**Program:** {program}  \n" 
                                        if pd.notna(deskripsi) and str(deskripsi).strip():
                                            item_string += f"**Deskripsi:** {deskripsi}"
                                        if item_string:
                                            existing_markdown_items.append(item_string)
                                            existing_data_list.append({'program': program, 'deskripsi': deskripsi})
                            
                            # --- 2. AMBIL DAN PROSES DATA KANAN (AI) ---
                            definition_hash = hashlib.md5(cluster_definition.encode()).hexdigest()[:8]
                            cache_key = f"insight_text_{selected_sheet}_{selected_cluster}_{definition_hash}"
                            ai_text_response = None
                            
                            if use_caching and cache_key in st.session_state:
                                st.toast("Mengambil hasil dari cache...")
                                ai_text_response = st.session_state[cache_key]
                            else:
                                with st.spinner(f"Gemini menganalisis tren untuk '{selected_cluster}'..."):
                                    ai_text_response = get_gemini_strategic_insight(
                                        divisi=selected_sheet, cluster=selected_cluster,
                                        cluster_definition=cluster_definition, 
                                        model_name="gemini-2.5-flash"
                                    )
                                if ai_text_response:
                                    if use_caching: st.session_state[cache_key] = ai_text_response
                                else:
                                    st.error("Panggilan API 1 (Insight) gagal.")
                                    ai_text_response = None
                            
                            # --- 3. PARSING DATA AI (JIKA API SUKSES) ---
                            ai_markdown_items = []
                            ai_data_list = [] 
                            sources_part = ""
                            regex_failed = False

                            if ai_text_response:
                                if not isinstance(ai_text_response, str):
                                    ai_text_response = str(ai_text_response)

                                content_part = ai_text_response
                                if "\nSumber" in ai_text_response:
                                    parts = ai_text_response.split("\nSumber", 1)
                                    content_part = parts[0]
                                    sources_part = "\nSumber" + parts[1] 
                                elif "Sumber:" in ai_text_response:
                                    parts = ai_text_response.split("Sumber:", 1)
                                    content_part = parts[0]
                                    sources_part = "Sumber:" + parts[1]
                                
                                pattern = re.compile(
                                    r"[•*-]\s*Program:\s*(.*?)" +                         
                                    r"(?:\n\s+[•*-]?\s*Deskripsi:\s*)(.*?)" +             
                                    r"(?:\n\s+[•*-]?\s*OKR\s*:\s*)(.*?)" +                
                                    r"(?=\n\s*[•*-]\s*Program:|\Z)",                    
                                    re.DOTALL | re.IGNORECASE
                                )
                                matches = pattern.findall(content_part)
                                
                                if matches:
                                    for match in matches:
                                        program = match[0].strip()
                                        deskripsi = match[1].strip()
                                        okr = match[2].strip()
                                        item_string = f"**Program:** {program}  \n"
                                        item_string += f"**Deskripsi:** {deskripsi}  \n"
                                        item_string += f"**OKR :** {okr}"
                                        ai_markdown_items.append(item_string)
                                        ai_data_list.append({'program': program, 'deskripsi': deskripsi, 'okr': okr})
                                else:
                                    if ai_text_response and not sources_part:
                                        regex_failed = True 
                                
                                # --- 4. TAMPILKAN HASIL (DAN TOMBOL DOWNLOAD) ---
                                
                                st.subheader(f"Perbandingan Strategis: {selected_cluster}")

                                # (Logika Tombol Download tidak berubah)
                                all_rows_data = []
                                max_rows_for_df = max(len(existing_data_list), len(ai_data_list))
                                for i in range(max_rows_for_df):
                                    row_data = {}
                                    if i < len(existing_data_list):
                                        row_data['Program_Existing'] = existing_data_list[i]['program']
                                        row_data['Deskripsi_Existing'] = existing_data_list[i]['deskripsi']
                                    else:
                                        row_data['Program_Existing'] = None
                                        row_data['Deskripsi_Existing'] = None
                                    if i < len(ai_data_list):
                                        row_data['Program_AI'] = ai_data_list[i]['program']
                                        row_data['Deskripsi_AI'] = ai_data_list[i]['deskripsi']
                                        row_data['OKR_AI'] = ai_data_list[i]['okr']
                                    else:
                                        row_data['Program_AI'] = None
                                        row_data['Deskripsi_AI'] = None
                                        row_data['OKR_AI'] = None
                                    all_rows_data.append(row_data)

                                df_download = pd.DataFrame(all_rows_data, columns=[
                                    'Program_Existing', 'Deskripsi_Existing', 
                                    'Program_AI', 'Deskripsi_AI', 'OKR_AI'
                                ])
                                excel_bytes = to_excel(df_download)
                                
                                st.download_button(
                                    label="📥 Download Hasil ke Excel",
                                    data=excel_bytes,
                                    file_name=f"Analisis_{selected_sheet}_{selected_cluster}.xlsx",
                                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                    use_container_width=True
                                )

                                # 1. Tambahkan CSS untuk garis vertikal
                                st.markdown("""
                                <style>
                                /* Memilih kolom pertama (div) di dalam sebuah stHorizontalBlock */
                                div[data-testid="stHorizontalBlock"] > div:first-child {
                                    border-right: 1px solid rgba(255, 255, 255, 0.2); /* Garis putih transparan */
                                    padding-right: 24px; /* Sesuaikan dengan 'gap' Anda */
                                }
                                
                                /* Memilih kolom kedua (div) di dalam sebuah stHorizontalBlock */
                                div[data-testid="stHorizontalBlock"] > div:nth-child(2) {
                                    padding-left: 24px; /* Sesuaikan dengan 'gap' Anda */
                                }

                                /* CSS untuk fallback box jika regex gagal */
                                .ai-pre { 
                                    background-color: rgba(255,255,255,0.02);
                                    border: 1px solid rgba(255,255,255,0.06);
                                    border-radius: 8px;
                                    padding: 12px;
                                    overflow-x: auto;
                                    font-family: monospace;
                                } 
                                </style>
                                """, unsafe_allow_html=True) 

                                # 2. Cek apakah ada data untuk ditampilkan
                                max_rows = max(len(existing_markdown_items), len(ai_markdown_items))

                                if max_rows == 0 and not regex_failed:
                                    st.warning(f"Tidak ada Program Kerja (Existing) atau Insight AI (New) yang ditemukan untuk cluster '{selected_cluster}'.")
                                
                                else:
                                    # 3. Buat Judul Kolom
                                    header_col1, header_col2 = st.columns(2, gap="medium")
                                    with header_col1:
                                        st.markdown("#### Mapping dari Spreadsheet")
                                    with header_col2:
                                        st.markdown("#### Insight AI")

                                    # 4. Buat Kolom Konten
                                    col1, col2 = st.columns(2, gap="medium")
                                    
                                    # --- KOLOM KIRI (EXISTING) ---
                                    with col1:
                                        # Gabungkan kembali dengan '---'
                                        final_existing_markdown = "\n\n---\n\n".join(existing_markdown_items)
                                        if final_existing_markdown:
                                            st.markdown(final_existing_markdown)
                                        elif not regex_failed: 
                                            st.markdown("*(Tidak ada data existing)*") 

                                    # --- KOLOM KANAN (AI) ---
                                    with col2:
                                        if regex_failed:
                                            st.error("Gagal mem-parsing output AI, menampilkan teks mentah:")
                                            st.markdown(f"<div class='ai-pre'>{html.escape(ai_text_response)}</div>", unsafe_allow_html=True)
                                        else:
                                            # Gabungkan kembali dengan '---'
                                            # JANGAN tambahkan 'sources_part' di sini
                                            final_ai_markdown = "\n\n---\n\n".join(ai_markdown_items)
                                            
                                            if final_ai_markdown:
                                                st.markdown(final_ai_markdown)
                                            elif not existing_markdown_items:
                                                st.markdown("*(Tidak ada insight AI)*")
                                            else:
                                                st.markdown(" ") # Beri spasi agar sejajar

                                    if sources_part and not regex_failed:
                                        st.markdown("---") # Garis pemisah dari kolom
                                        st.markdown(sources_part.strip())

                            elif not ai_text_response and not existing_markdown_items:
                                # Handle jika API gagal total DAN tidak ada data existing
                                st.error("Gagal mendapatkan insight dari AI dan tidak ada data existing untuk ditampilkan.")

                else:
                    st.error(f"Sheet '{selected_sheet}' tidak memiliki kolom 'Cluster'. Mohon periksa file Anda.")

        except Exception as e:
            safe_error = html.escape(str(e))
            st.error(f"An error occurred while reading the Excel file: {safe_error}")
            
    elif file_extension == 'csv':
        st.warning("Mode CSV memiliki fungsionalitas terbatas.")
        with st.spinner("Loading CSV file..."):
            try:
                df = pd.read_csv(uploaded_file)
                st.dataframe(df.head())
                st.info("Harap gunakan format .xlsx untuk fungsionalitas penuh (pemilihan Divisi dan Cluster).")
            except Exception as e:
                st.error(f"An error occurred while reading the CSV file: {e}")

# --- FOOTER ---
st.markdown("---")
st.markdown(
    """
    <p style='text-align: center; color: grey;'>
        ⚡ Powered by Google Gemini
    </p>
    <p style='text-align: center; color: grey; font-size: 0.8em;'>
        PT. Midi Utama Indonesia Tbk.
    </p>
    """,
    unsafe_allow_html=True
)