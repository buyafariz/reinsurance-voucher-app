import streamlit as st
import pandas as pd
import os
import time
import st_aggrid


from datetime import datetime
from validator import validate_voucher
from vin_generator import generate_vin, create_cancel_row, get_log_path, generate_vin_from_drive, generate_vin_from_drive_log, create_negative_excel, dataframe_to_excel_bytes, upload_excel_bytes, get_log_filename, get_log_pml_filename, get_log_filename_outward, generate_vou_from_drive, generate_pml_from_drive, split_upload_with_log, get_last_seq_no, generate_pml_id
from drive_utils import upload_or_update_drive_file, get_period_drive_folders, get_or_create_folder, get_or_create_ceding_folders, get_drive_service, find_drive_file, acquire_drive_lock, release_drive_lock, upload_dataframe_to_drive, load_log_from_drive, upload_log_dataframe, load_voucher_excel_from_drive, calculate_due_date, get_exchange_rate, load_log_from_gsheet, update_gsheet, append_gsheet, create_log_gsheet, get_or_create_outward_folders, upload_dataframe_to_drive_outward, init_sheets_service, download_file_from_drive, update_pml_status_to_splitted
from lock_utils import acquire_lock, release_lock
from zoneinfo import ZoneInfo
from st_aggrid import AgGrid, GridOptionsBuilder, GridUpdateMode, JsCode
from io import BytesIO
from st_aggrid import JsCode
from datetime import date
from google.oauth2 import service_account

if "is_processing_split" not in st.session_state:
    st.session_state.is_processing_split = False

creds = service_account.Credentials.from_service_account_info(
    st.secrets["gcp_service_account"],
    scopes=[
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
)


def normalize_folder_name(name: str) -> str:
    return (
        name.upper()
        .replace("/", "-")
        .replace("&", "AND")
        .replace("(", "")
        .replace(")", "")
        .strip()
    )


def now_wib_naive():
    return datetime.now(ZoneInfo("Asia/Jakarta")).replace(tzinfo=None)


# ==========================
# ACCOUNTING FORMAT CONFIG
# ==========================
ACCOUNTING_COLS = [
    "sum insured",
    "sum at risk",
    "reins sum insured",
    "reins sum at risk",
    "reins total premium",
    "reins total comm",
    "reins tabarru",
    "reins ujrah",
    "reins nett premium",
]

def accounting_format(x):
    if pd.isna(x):
        return ""
    x = float(x)
    if x == 0:
        return "–"
    if x < 0:
        return f"({abs(x):,.2f})"
    return f"{x:,.2f}"

MONTH_ID = [
    "", "Januari", "Februari", "Maret", "April",
    "Mei", "Juni", "Juli", "Agustus",
    "September", "Oktober", "November", "Desember"]

# ==========================
# CONFIG
# ==========================
st.set_page_config(
    page_title="Retakaful Voucher Tools",
    layout="centered"
)

if "log_period" not in st.session_state:
    now = datetime.now(ZoneInfo("Asia/Jakarta"))
    st.session_state["log_period"] = {
        "year": now.year,
        "month": now.month
    }


ROOT_DRIVE_FOLDER_ID = st.secrets["drive_folder_id"]
CONFIG_FOLDER_ID = st.secrets["config_folder_id"]
RATE_FOLDER_ID = st.secrets["rate_folder_id"]


st.title("📄 Retakaful Voucher Tools")
st.write("")

tab_upload, tab_calc, tab_post, tab_update = st.tabs([
    "📤 Upload File",
    "🧮 Calculate",
    "📥 Create Voucher",
    "🔄 Update Voucher",
])


# ==========================
# SIMPAN VOUCHER
# ==========================

with tab_upload:
    st.subheader("📤 Upload File")
    
    # ===== ROW 1 =====
    row1_col1, row1_col2 = st.columns(2)

    with row1_col1:
        reins_type = st.selectbox(
            "Reinsurance Type",
            ["INWARD", "OUTWARD"],
            key="reins_type_upload",
            index=0, #INWARD
            disabled=True
        )

    # row1_col2 sengaja dikosongkan


    # ===== ROW 2 =====
    row2_col1, row2_col2 = st.columns(2)

    with row2_col1:
        department = st.selectbox(
            "Department",
            ["ADMIN", "CLAIM"],
            key="department_upload"
        )

    with row2_col2:
        biz_type = st.selectbox(
            "Biz Type",
            [
                "Kontribusi",
                "Claim",
                "Refund",
                "Alteration",
                "Retur",
                "Revise",
                "Batal",
                "Cancel"
            ],
            key="biz_type_upload"
        )

    uploaded_file = st.file_uploader(
        "Upload Voucher (.xlsx)",
        type=["xlsx"],
        key="upload_post_upload"
    )

    columns_template = [
        "No",
        "TL Detail ID",
        "Trans Category",
        "Policy Category",
        "Certificate No",
        "Insured Full Name",
        "Gender",
        "Main Pol No",
        "Main Policy",
        "Pol Holder No",
        "Policy Holder",
        "Birth Date",
        "Age At",
        "Issue Date",
        "Term Year",
        "Term Month",
        "Expired Date",
        "Medical",
        "Ced Risk Code",
        "Life Risk Name",
        "K.O.B Code",
        "Ced Product Code",
        "Ced Coverage Code",
        "Ced Product Desc",
        "Ccy Code",
        "Sum Insured",
        "Sum At Risk",
        "Reins Sum Insured",
        "Ced Retention",
        "Reins Sum At Risk",
        "Pay Period Type",
        "Ced EM Rate",
        "Ced ER Rate",
        "Reins Premium",
        "Reins EM Premium",
        "Reins ER Premium",
        "Reins Oth. Premium",
        "Reins Total Premium",
        "Reins Comm",
        "Reins EM Comm",
        "Reins ER Comm",
        "Reins Oth. Comm",
        "Reins Profit Share",
        "Reins Overriding",
        "Reins Broker Fee",
        "Reins Total Comm",
        "Reins Tabarru",
        "Reins Ujrah",
        "Reins Nett Premium",
        "Valuation Date",
        "Terminate Date",
        "TL Detail Remarks",
        "CBY",
        "CBM",
        "COB",
        "Voucher ID",
        "PML ID",
        "References No",
        "Elapse No",
        "Ref Voucher ID"
    ]

    columns_template_claim = [
        "BookYear",
        "BookMonth",
        "CedBookYear",
        "CedBookMonth",
        "Company Name",
        "Policy Holder No",
        "Policy Holder",
        "Certificate No",
        "Insured Name",
        "Birth Date",
        "Age",
        "Gender",
        "Sum Insured IDR",
        "Sum Reinsured IDR",
        "medicalcategory",
        "Product",
        "Coverage Code",
        "ClassOfBusiness",
        "PayPeriodType",
        "Issue Date",
        "Term Year",
        "Term Month",
        "End Date Policy",
        "Claim Date",
        "Claim Register Date",
        "Payment Date",
        "Currency",
        "ExchangeRate",
        "Amount of Claim IDR",
        "Reins Claim IDR",
        "Marein Share IDR",
        "Cause Of Claim",
        "Voucher ID",
        "References No"        
    ]


    if uploaded_file:
        # ==========================
        # READ FILE
        # ==========================
        df = pd.read_excel(uploaded_file)
        original_columns = df.columns.tolist()
        df.columns = df.columns.str.strip().str.lower()

        for col in ["certificate no", "pol holder no"]:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip()

        # ==========================
        # VALIDATION
        # ==========================
        errors = validate_voucher(df, st.session_state["biz_type_upload"], st.session_state["reins_type_upload"])

        if errors:
            st.error("❌ VALIDASI GAGAL")
            for err in errors:
                st.write(f"- {err}")
            st.stop()

        st.success("✅ Validasi berhasil")


        # ==========================
        # PREVIEW + FILTER (DINAMIS)
        # ==========================

        def get_non_empty_columns(df: pd.DataFrame):
            valid_cols = []
            for col in df.columns:
                series = df[col]
                non_na = series.dropna()

                if non_na.empty:
                    continue

                if series.dtype == "object":
                    if non_na.astype(str).str.strip().ne("").any():
                        valid_cols.append(col)
                else:
                    valid_cols.append(col)

            return valid_cols


        # ==========================
        # PREVIEW + FILTER
        # ==========================
        with st.expander("📊 Preview Data Voucher", expanded=True):
            if not df.empty:
                # 1. Batasi jumlah baris agar aplikasi tetap cepat
                MAX_PREVIEW = 1000
                total_rows = len(df)
                preview_df = df.head(MAX_PREVIEW).copy()

                st.caption(f"Menampilkan {len(preview_df):,} dari {total_rows:,} baris")

                # 2. SANITIZE & FORMATTING (Sama seperti cara Summary Financial)
                # Pastikan kolom accounting diformat dengan ribuan dan 2 desimal
                
                ACCOUNTING_COLS = [
                    "sum insured", "sum at risk", "reins sum insured", "reins sum at risk",
                    "reins premium", "reins em premium", "reins er premium", "reins total premium",
                    "reins total comm", "reins tabarru", "reins ujrah", "reins nett premium"
                ]

                # Buat dictionary formatter untuk kolom yang ada saja
                format_dict = {}
                for col in ACCOUNTING_COLS:
                    if col in preview_df.columns:
                        # Pastikan data adalah numerik sebelum diformat
                        preview_df[col] = pd.to_numeric(preview_df[col], errors='coerce').fillna(0)
                        format_dict[col] = "{:,.2f}"

                # 3. RENDER MENGGUNAKAN ST.DATAFRAME (Sama dengan Summary Financial)
                try:
                    st.dataframe(
                        preview_df.style.format(format_dict),
                        use_container_width=True,
                        height=450 # Memberikan scrollbar internal jika data banyak
                    )
                except Exception as e:
                    st.error(f"Gagal menampilkan preview: {e}")
                    st.dataframe(preview_df) # Fallback ke tabel mentah jika styling gagal
            else:
                st.info("Belum ada data untuk ditampilkan. Silakan upload file terlebih dahulu.")


        # ==========================
        # PERIOD & LOG
        # ==========================
        year = st.session_state["log_period"]["year"]
        month = st.session_state["log_period"]["month"]

        # ==========================
        # DRIVE FOLDER PER PERIODE (STEP 3)
        # ==========================
        drive_folders = get_period_drive_folders(
            year=year,
            month=month,
            root_folder_id=ROOT_DRIVE_FOLDER_ID
        )

        PERIOD_DRIVE_ID = drive_folders["period_id"]


        # ==========================
        # FORM INPUT
        # ==========================

        with st.expander("🧾 Informasi Voucher", expanded=True):

            col1, col2 = st.columns(2)

            with col1:

                pic = st.selectbox("PIC", ["Ardelia", "Buya", "Khansa", "Prabu"])

                account_with = st.selectbox(
                    "Account With",
                    [
                        "AIA FINANCIAL SYARIAH",
                        "AJS KITABISA (D/H AMANAH GITHA)",
                        "ALLIANZ LIFE SYARIAH",
                        "ALLIANZ LIFE SYARIAH (Health)",
                        "ALLIANZ LIFE SYARIAH (FlexiCare)",
                        "ALLIANZ LIFE SYARIAH (HSCP)",
                        "ALLIANZ LIFE SYARIAH (Individu DMTM)",
                        "ASTRA AVIVA LIFE",
                        "AVRIST ASSURANCE SYARIAH",
                        "AXA FINANCIAL INDONESIA SYARIAH",
                        "AXA MANDIRI FINANCIAL SERVICES SYARIAH",
                        "BNI LIFE SYARIAH",
                        "BRINGIN LIFE SYARIAH",
                        "BUMIPUTERA SYARIAH",
                        "CAPITAL LIFE SYARIAH",
                        "CENTRAL ASIA RAYA SYARIAH",
                        "FWD LIFE INDONESIA SYARIAH",
                        "GENERALI INDONESIA LIFE ASSURANCE SYARIAH",
                        "GREAT EASTERN LIFE SYARIAH",
                        "JASA MITRA ABADI SYARIAH",
                        "MANULIFE INDONESIA SYARIAH",
                        "PFI MEGA LIFE INSURANCE SYARIAH",
                        "PRUDENTIAL LIFE SYARIAH",
                        "PT ASURANSI JIWA SYARIAH BUMIPUTERA",
                        "REASURANSI INTERNATIONAL INDONESIA SYARIAH",
                        "RELIANCE SYARIAH",
                        "SINARMAS SYARIAH",
                        "SUN LIFE SYARIAH",
                        "SYARIAH AL-AMIN",
                        "TAKAFUL KELUARGA",
                        "GENERAL REINSURANCE AG (GEN RE) PLC, SINGAPORE",
                        "HANNOVER RETAKAFUL",
                        "MAREIN SYARIAH",
                        "MUNICH RE RETAKAFUL",
                        "SCOR SE LABUAN BRANCH",
                        "SWISS RE INTL. SE, SINGAPORE (SYARIAH)"
                    ]
                )

                cedant_company = st.selectbox(
                    "Cedant Company",
                    [
                        "AIA FINANCIAL SYARIAH",
                        "AJS KITABISA (D/H AMANAH GITHA)",
                        "ALLIANZ LIFE SYARIAH",
                        "ASTRA AVIVA LIFE",
                        "AVRIST ASSURANCE SYARIAH",
                        "AXA FINANCIAL INDONESIA SYARIAH",
                        "AXA MANDIRI FINANCIAL SERVICES SYARIAH",
                        "BNI LIFE SYARIAH",
                        "BRINGIN LIFE SYARIAH",
                        "BUMIPUTERA SYARIAH",
                        "CAPITAL LIFE SYARIAH",
                        "CENTRAL ASIA RAYA SYARIAH",
                        "FWD LIFE INDONESIA SYARIAH",
                        "GENERALI INDONESIA LIFE ASSURANCE SYARIAH",
                        "GREAT EASTERN LIFE SYARIAH",
                        "JASA MITRA ABADI SYARIAH",
                        "MANULIFE INDONESIA SYARIAH",
                        "PANIN DAICHI LIFE SYARIAH",
                        "PFI MEGA LIFE INSURANCE SYARIAH",
                        "PRUDENTIAL LIFE SYARIAH",
                        "PT ASURANSI JIWA SYARIAH BUMIPUTERA",
                        "REASURANSI INTERNATIONAL INDONESIA SYARIAH",
                        "RELIANCE SYARIAH",
                        "SINARMAS SYARIAH",
                        "SUN LIFE SYARIAH",
                        "SYARIAH AL-AMIN",
                        "TAKAFUL KELUARGA",
                        "GENERAL REINSURANCE AG (GEN RE) PLC, SINGAPORE",
                        "HANNOVER RETAKAFUL",
                        "MAREIN SYARIAH",
                        "MUNICH RE RETAKAFUL",
                        "SCOR SE LABUAN BRANCH",
                        "SWISS RE INTL. SE, SINGAPORE (SYARIAH)"
                    ]
                )

                curr = st.selectbox(
                    "Currency",
                    ["IDR", "USD"]
                )

            with col2:
                
                subject_email = st.text_area("Subject Email")

                email_date = st.date_input("Email Date",value=date.today())

                remarks = st.text_area("Remarks")


        # ==========================
        # FINANCIAL SUMMARY
        # ==========================
        st.subheader("💰 Summary Financial")

        if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
            summary_df = pd.DataFrame({
                "Keterangan": [
                    "Total Contribution",
                    "Commission",
                    "Tabarru",
                    "Ujrah",
                    "Nett Premium"
                ],
                "Nilai": [
                    df["reins total premium"].sum(),
                    df["reins total comm"].sum(),
                    df["reins tabarru"].sum(),
                    df["reins ujrah"].sum(),
                    df["reins nett premium"].sum(),
                ]
            })

        elif biz_type == "Claim":
            summary_df = pd.DataFrame({
                "Keterangan": [
                    "Amount of Claim IDR",
                    "Reins Claim IDR",
                    "Marein Share IDR"
                ],
                "Nilai": [
                    df["amount of claim idr"].sum(),
                    df["reins claim idr"].sum(),
                    df["marein share idr"].sum()
                    ]
            })
            

        st.dataframe(
            summary_df.style.format({"Nilai": "{:,.2f}"}),
            use_container_width=True
        )


        # ==========================
        # POST VOUCHER (LOCKED)
        # ==========================
        if st.button("💾 Simpan File"):
            start_time = time.time()

            if not remarks.strip() or not subject_email.strip():
                st.error("Product, Subject Email, dan Remarks wajib diisi")
                st.stop()

            #lock_path = log_path + ".lock"
            service = get_drive_service()

            with st.spinner("⏳ Menyimpan voucher, mohon tunggu..."):

                try:
                    service = get_drive_service()                    

                    drive_folders = get_period_drive_folders(
                        year=int(year),
                        month=int(month),
                        root_folder_id=ROOT_DRIVE_FOLDER_ID
                    )

                    PML_folders = get_period_drive_folders(
                        year=int(year),
                        month=int(month),
                        root_folder_id=ROOT_DRIVE_FOLDER_ID
                    )

                    PERIOD_DRIVE_ID = drive_folders["period_id"]

                    acquire_drive_lock(service, PERIOD_DRIVE_ID)

                    # reload log terbaru setelah lock
                    # if os.path.exists(log_path):
                    #     log_df = pd.read_excel(log_path)
                    # else:
                    #     log_df = pd.DataFrame()

                    pml_drive = get_or_create_folder(
                        service=service,
                        folder_name="Folder PML",
                        parent_id=PERIOD_DRIVE_ID
                    )

                    PML_DRIVE_ID = pml_drive

                    pml_id, seq_no, file_id = generate_pml_from_drive(
                        service=service,
                        period_folder_id=PML_DRIVE_ID,
                        year=int(year),
                        month=int(month),
                        find_drive_file=find_drive_file,
                        biz_type = biz_type
                    )

                    # Upload voucher (selalu CREATE)
                    log_pml_drive_id = find_drive_file(
                        service=service,
                        filename=get_log_pml_filename(int(year), int(month)),
                        # filename="log_produksi.xlsx",
                        parent_id=PML_DRIVE_ID,
                        mime_type="application/vnd.google-apps.spreadsheet"
                    )

                    rate_exchange = get_exchange_rate(
                        service=service,
                        config_folder_id=CONFIG_FOLDER_ID,
                        currency=curr,
                        month=month
                    )

                    if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
                        log_pml = {
                            "Seq No": seq_no,
                            "Department":department,
                            "Biz Type": biz_type,
                            "PML ID": pml_id,
                            "Account With": account_with,
                            "Cedant Company": cedant_company,
                            "PIC": pic,
                            "Curr":curr,
                            "Total Contribution": df["reins total premium"].sum(),
                            "Commission": df["reins total comm"].sum(),
                            "Overriding": df["reins overriding"].sum() if "reins overriding" in df.columns else 0,
                            "Total Commission": (df["reins total comm"].sum()) + (df["reins overriding"].sum() if "reins overriding" in df.columns else 0),
                            "Gross Premium Income": df["reins total premium"].sum() - ((df["reins total comm"].sum()) + (df["reins overriding"].sum() if "reins overriding" in df.columns else 0)),
                            "Tabarru": df["reins tabarru"].sum(),
                            "Ujrah": df["reins ujrah"].sum(),
                            "Claim": 0,
                            "Balance": df["reins total premium"].sum() - df["reins total comm"].sum() - (df["reins overriding"].sum() if "reins overriding" in df.columns else 0) - (df["claim"].sum() if "claim" in df.columns else 0),
                            "REMARKS": remarks,
                            "STATUS": "POSTED",
                            #"ENTRY_TYPE": entry_type,
                            "CREATED AT": now_wib_naive(),
                            "CREATED BY": pic,
                            "Subject Email": subject_email,
                            "Email Date": email_date,
                            "CANCELED AT": "-",
                            "CANCELED BY": "-",
                            "CANCEL OF VOUCHER": "-",
                            "CANCEL REASON":"-"
                        }

                    elif biz_type == "Claim":
                        log_pml = {
                            "Seq No": seq_no,
                            "Department":department,
                            "Biz Type": biz_type,
                            "PML ID": pml_id,
                            "Account With": account_with,
                            "Cedant Company": cedant_company,
                            "PIC": pic,
                            "Curr":curr,
                            "Total Contribution": 0,
                            "Commission": 0,
                            "Overriding": 0,
                            "Total Commission": 0,
                            "Gross Premium Income": 0,
                            "Tabarru": 0,
                            "Ujrah": 0,
                            "Claim": df["marein share idr"].sum(),
                            "Balance": 0 - (df["marein share idr"].sum() if "marein share idr" in df.columns else 0),
                            "REMARKS": remarks,
                            "STATUS": "POSTED",
                            #"ENTRY_TYPE": entry_type,
                            "CREATED AT": now_wib_naive(),
                            "CREATED BY": pic,
                            "Subject Email": subject_email,
                            "Email Date": email_date,
                            "CANCELED AT": "-",
                            "CANCELED BY": "-",
                            "CANCEL OF VOUCHER": "-",
                            "CANCEL REASON": "-"
                        }

                    # log_drive_id = find_drive_file(
                    #     service=service,
                    #     filename=get_log_filename(int(oby), int(obm)),
                    #     parent_id=PERIOD_DRIVE_ID,
                    #     mime_type="application/vnd.google-apps.spreadsheet"
                    # )

                    if not log_pml_drive_id:
                        log_pml_drive_id = create_log_gsheet(
                            service=service,
                            parent_id=PML_DRIVE_ID,
                            filename=get_log_pml_filename(int(year), int(month)),
                            columns=list(log_pml.keys())
                        )

                    sheets_service = init_sheets_service(creds)

                    append_gsheet(
                        service=sheets_service,
                        spreadsheet_id=log_pml_drive_id,
                        row_dict=log_pml
                    )

                    if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
                        upload_dataframe_to_drive(
                            service=service,
                            df=df,
                            template_columns=columns_template,
                            voucher_id=pml_id,
                            filename=f"{pml_id}.xlsx",
                            folder_id=PML_DRIVE_ID,
                            file_type = "PML"
                        )

                    elif biz_type == "Claim" :
                        upload_dataframe_to_drive(
                            service=service,
                            df=df,
                            template_columns=columns_template_claim,
                            voucher_id=pml_id,
                            filename=f"{pml_id}.xlsx",
                            folder_id=PML_DRIVE_ID,
                            file_type="PML"
                        )

                    end_time = time.time()
                    duration = end_time - start_time

                    st.success(f"✅ File berhasil diposting: {pml_id} ({int(duration)} seconds)")
                    st.code(pml_id)

                except RuntimeError as e:
                        st.error("⛔ Log sedang digunakan user lain. Silakan coba lagi.")
                        st.stop()

                finally:
                    release_drive_lock(service, PERIOD_DRIVE_ID)


# ==========================
# TAB CALCULATE
# ==========================
with tab_calc:

    st.subheader("📊 Calculate PML")

    # ==========================
    # INIT SERVICE
    # ==========================
    service = get_drive_service()
    sheets_service = init_sheets_service(creds)

    year = st.session_state["log_period"]["year"]
    month = st.session_state["log_period"]["month"]

    # ==========================
    # GET PERIOD FOLDER
    # ==========================
    drive_folders = get_period_drive_folders(
        year=int(year),
        month=int(month),
        root_folder_id=ROOT_DRIVE_FOLDER_ID
    )

    PERIOD_DRIVE_ID = drive_folders["period_id"]

    # ==========================
    # GET PML FOLDER
    # ==========================
    pml_drive = get_or_create_folder(
        service=service,
        folder_name="Folder PML",
        parent_id=PERIOD_DRIVE_ID
    )

    PML_DRIVE_ID = pml_drive

    # ==========================
    # GET LOG PML FILE
    # ==========================
    log_pml_drive_id = find_drive_file(
        service=service,
        filename=get_log_pml_filename(int(year), int(month)),
        parent_id=PML_DRIVE_ID,
        mime_type="application/vnd.google-apps.spreadsheet"
    )

    if not log_pml_drive_id:
        st.warning("⚠️ Log PML belum tersedia")
        st.stop()

    # ==========================
    # LOAD LOG DATA
    # ==========================
    log_df = load_log_from_gsheet(
        service=sheets_service,
        spreadsheet_id=log_pml_drive_id
    )

    if log_df.empty:
        st.warning("⚠️ Log PML kosong")
        st.stop()

    # ==========================
    # NORMALIZE COLUMN
    # ==========================
    log_df.columns = log_df.columns.str.strip()

    # ==========================
    # VALIDASI KOLOM
    # ==========================
    required_cols = ["PML ID", "STATUS"]

    missing_cols = [col for col in required_cols if col not in log_df.columns]

    if missing_cols:
        st.error(f"❌ Kolom tidak ditemukan: {missing_cols}")
        st.stop()

    # ==========================
    # FILTER STATUS POSTED
    # ==========================
    df_posted = log_df[log_df["STATUS"] == "POSTED"].copy()

    if df_posted.empty:
        st.info("Tidak ada data dengan status POSTED")
        st.stop()

    # ==========================
    # SEARCH (OPSIONAL)
    # ==========================
    search = st.text_input("🔍 Cari PML ID")

    if search:
        df_posted = df_posted[
            df_posted["PML ID"].astype(str).str.contains(search, case=False, na=False)
        ]

    st.write(f"Total PML POSTED: {len(df_posted)}")

    # ==========================
    # UI SELECT (SAMA DENGAN SPLIT)
    # ==========================
    # st.markdown("### 📋 Pilih Data PML untuk Di-Calculate")
    st.info("Centang pada kolom **'Pilih'** untuk menentukan baris yang akan diproses.")

    if not df_posted.empty:

        # Tambahkan checkbox column
        df_to_edit = df_posted.copy()
        df_to_edit.insert(0, "Pilih", False)

        # Data editor (CONSISTENT UI)
        edited_df = st.data_editor(
            df_to_edit,
            column_config={
                "Pilih": st.column_config.CheckboxColumn(
                    "Pilih",
                    help="Pilih baris ini untuk di-calculate",
                    default=False,
                ),
                "PML ID": st.column_config.Column(disabled=True),
                "STATUS": st.column_config.Column(disabled=True),
                "Product": st.column_config.Column(disabled=True),
                "Total Contribution": st.column_config.NumberColumn(
                    "Total Contribution",
                    format="#,##0",
                    disabled=True
                ),
            },
            disabled=["No", "PML ID", "STATUS", "Product", "Total Contribution"],
            hide_index=True,
            use_container_width=True,
        )

        # ==========================
        # AMBIL YANG DIPILIH
        # ==========================
        selected_rows = edited_df[edited_df["Pilih"] == True]

        # ==========================
        # VALIDASI
        # ==========================
        if len(selected_rows) > 1:
            st.warning("⚠️ Anda memilih lebih dari 1 baris. Harap pilih **satu baris saja** untuk proses calculate.")

        elif len(selected_rows) == 1:
            selected_pml_id = selected_rows.iloc[0]["PML ID"]
            st.success(f"✅ Baris terpilih: **{selected_pml_id}**")

        else:
            st.info("Silakan pilih satu baris untuk melanjutkan.")

    
        # ==========================
        # CEK RATE FILE
        # ==========================
        if not selected_rows.empty:
            selected_account = selected_rows.iloc[0]["Account With"]
        else:
            selected_account = None

        rate_file_id = find_drive_file(
            service=service,
            filename=f"{selected_account}.xlsx",
            parent_id=RATE_FOLDER_ID
        )

        has_rate = rate_file_id is not None

        # st.markdown("### ⚙️ Pilih Metode Calculate")
        # ==========================
        # LAYOUT 4 KOLOM
        # ==========================
        col1, col2, col3, col4 = st.columns([1, 1, 1.5, 1.5])

        with col3:
            ceding_clicked = st.button(
                "📥 Ceding Calculation",
                use_container_width=True,
                disabled = (selected_account is None),
                type="primary"
            )

        with col4:
            our_clicked = st.button(
                "🧮 Our Calculation",
                use_container_width=True,
                disabled= (selected_account is None or not has_rate),
                help="Rate belum tersedia"
            )

        # ==========================
        # INFO JIKA RATE BELUM ADA
        # ==========================
        if not has_rate:
            st.info("ℹ️ Rate belum tersedia → Our Calculate dinonaktifkan")


        # ==========================
        # POST VOUCHER (LOCKED)
        # ==========================
        if ceding_clicked:
            start_time = time.time()

            #lock_path = log_path + ".lock"
            service = get_drive_service()

            with st.spinner("⏳ Calculation sedang berjalan, mohon tunggu..."):

                try:
                    service = get_drive_service()                    

                    drive_folders = get_period_drive_folders(
                        year=int(year),
                        month=int(month),
                        root_folder_id=ROOT_DRIVE_FOLDER_ID
                    )

                    PERIOD_DRIVE_ID = drive_folders["period_id"]

                    pml_drive = get_or_create_folder(
                        service=service,
                        folder_name="Folder PML",
                        parent_id=PERIOD_DRIVE_ID
                    )

                    PML_DRIVE_ID = pml_drive

                    st.write("PML_DRIVE_ID:", PML_DRIVE_ID)

                    # acquire_drive_lock(service, PML_DRIVE_ID)

                    voucher, seq_no, file_id = generate_vin_from_drive(
                        service=service,
                        period_folder_id=PML_DRIVE_ID,
                        year=int(year),
                        month=int(month),
                        find_drive_file=find_drive_file,
                        biz_type=selected_rows.iloc[0]["Biz Type"]
                    )

                    st.write(f"File ID: {file_id}")

                    ceding_folder_name = normalize_folder_name(selected_rows.iloc[0]["Account With"])

                    ceding_drive = get_or_create_ceding_folders(
                        service=service,
                        period_folder_id=PERIOD_DRIVE_ID,
                        ceding_name=ceding_folder_name
                    )

                    CEDING_DRIVE_ID = ceding_drive["ceding_id"]


                    # Upload voucher (selalu CREATE)
                    log_pml_drive_id = find_drive_file(
                        service=service,
                        filename=get_log_pml_filename(int(year), int(month)),
                        # filename="log_produksi.xlsx",
                        parent_id=PML_DRIVE_ID,
                        mime_type="application/vnd.google-apps.spreadsheet"
                    )

                    st.write(f"Log PML drive id: {log_pml_drive_id}")

                    rate_exchange = get_exchange_rate(
                        service=service,
                        config_folder_id=CONFIG_FOLDER_ID,
                        currency=selected_rows.iloc[0]["Curr"],
                        month=month
                    )

                    due_date = calculate_due_date(
                        account_with=selected_rows.iloc[0]["Account With"],
                        year=year,
                        month=month,
                        service=service
                    )

                    st.write(f"PML ID: {selected_rows.iloc[0]["PML ID"]}")

                    # Cari file
                    pml_file_id = find_drive_file(
                        service=service,
                        filename= str(selected_rows.iloc[0]["PML ID"]).strip(),
                        parent_id=PML_DRIVE_ID
                    )

                    if not pml_file_id:
                        st.error("File PML tidak ditemukan")
                        st.stop()

                    # Load file PML
                    file_stream = download_file_from_drive(service, pml_file_id)
                    df = pd.read_excel(file_stream)

                    biz_type = selected_rows.iloc[0]["Biz Type"]


                    if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
                        log_entry = {
                            "Seq No": seq_no,
                            "Department":selected_rows.iloc[0]["Department"],
                            "Biz Type": selected_rows.iloc[0]["Biz Type"],
                            "Voucher No": voucher,
                            "Account With": selected_rows.iloc[0]["Account With"],
                            "Cedant Company": selected_rows.iloc[0]["Cedant Company"],
                            "PIC": selected_rows.iloc[0]["PIC"],
                            "Product": df["References No"],
                            "CBY": df["CBY"],
                            "CBM": df["CBM"],
                            "OBY": int(year),
                            "OBM": int(month),
                            "KOB": df["K.O.B Code"],
                            "COB": df["COB"],
                            "MOP": df["Pay Period Type"],
                            "Curr": df["Ccy Code"],
                            "Total Contribution": df["Reins Total Premium"].sum(),
                            "Commission": df["Reins Total Comm"].sum(),
                            "Overriding": df["Reins Overriding"].sum() if "Reins Overriding" in df.columns else 0,
                            "Total Commission": (df["Reins Total Comm"].sum()) + (df["Reins Overriding"].sum() if "Reins Overriding" in df.columns else 0),
                            "Gross Premium Income": df["Reins Total Premium"].sum() - ((df["Reins Total Comm"].sum()) + (df["Reins Overriding"].sum() if "Reins Overriding" in df.columns else 0)),
                            "Tabarru": df["Reins Tabarru"].sum(),
                            "Ujrah": df["Reins Ujrah"].sum(),
                            "Claim": 0,
                            "Balance": df["Reins Total Premium"].sum() - df["Reins Total Comm"].sum() - (df["Reins Overriding"].sum() if "Reins Overriding" in df.columns else 0) - (df["Claim"].sum() if "Claim" in df.columns else 0),
                            "Check Balance": "",
                            "Rate Exchange": rate_exchange,
                            "Kontribusi (IDR)": (df["Reins Total Premium"].sum())*rate_exchange,
                            "Commission (IDR)": (df["Reins Total Comm"].sum())*rate_exchange,
                            "Overiding (IDR)": (df["Reins Overriding"].sum() if "Reins Overriding" in df.columns else 0)*rate_exchange,
                            "Total Commission (IDR)": ((df["Reins Total Comm"].sum()) + (df["Reins Overriding"].sum() if "Reins Overriding" in df.columns else 0))*rate_exchange,
                            "Gross Premium Income (IDR)": (df["Reins Total Premium"].sum() - ((df["Reins Total Comm"].sum()) + (df["Reins Overriding"].sum() if "Reins Overriding" in df.columns else 0)))*rate_exchange,
                            "Tabarru (IDR)": (df["Reins Tabarru"].sum())*rate_exchange,
                            "Ujrah (IDR)": (df["Reins Ujrah"].sum())*rate_exchange,
                            "Claim (IDR)": 0,
                            "Balance (IDR)": (df["Reins Total Premium"].sum() - df["Reins Total Comm"].sum() - (df["Reins Overriding"].sum() if "Reins Overriding" in df.columns else 0) - (df["Claim"].sum() if "Claim" in df.columns else 0))*rate_exchange,
                            "Check Balance (IDR)":"",
                            "REMARKS": selected_rows.iloc[0]["Remarks"],
                            "STATUS": "POSTED",
                            "CREATED AT": now_wib_naive(),
                            "CREATED BY": pic,
                            "Due Date": due_date,
                            "Subject Email": selected_rows.iloc[0]["Subject Email"],
                            "Email Date": selected_rows.iloc[0]["Email Date"],
                            "CANCELED AT": "-",
                            "CANCELED BY": "-",
                            "CANCEL OF VOUCHER": "-",
                            "CANCEL REASON":"-"
                        }

                    elif biz_type == "Claim":
                        log_entry = {
                            "Seq No": seq_no,
                            "Department":selected_rows.iloc[0]["Department"],
                            "Biz Type": selected_rows.iloc[0]["Biz Type"],
                            "Voucher No": voucher,
                            "Account With": selected_rows.iloc[0]["Account With"],
                            "Cedant Company": selected_rows.iloc[0]["Cedant Company"],
                            "PIC": selected_rows.iloc[0]["PIC"],
                            "Product": df["References No"],
                            "CBY": df["CBY"],
                            "CBM": df["CBM"],
                            "OBY": int(year),
                            "OBM": int(month),
                            "KOB": df["K.O.B Code"],
                            "COB": df["COB"],
                            "MOP": df["Pay Period Type"],
                            "Curr": df["Ccy Code"],
                            "Total Contribution": 0,
                            "Commission": 0,
                            "Overriding": 0,
                            "Total Commission": 0,
                            "Gross Premium Income": 0,
                            "Tabarru": 0,
                            "Ujrah": 0,
                            "Claim": df["Marein Share IDR"].sum(),
                            "Balance": 0 - (df["Marein Share IDR"].sum() if "Marein Share IDR" in df.columns else 0),
                            "Check Balance": "",
                            "Rate Exchange": rate_exchange,
                            "Kontribusi (IDR)": 0,
                            "Commission (IDR)": 0,
                            "Overiding (IDR)": 0,
                            "Total Commission (IDR)": 0,
                            "Gross Premium Income (IDR)": 0,
                            "Tabarru (IDR)": 0,
                            "Ujrah (IDR)": 0,
                            "Claim (IDR)": (df["Marein Share IDR"].sum() if "Marein Share IDR" in df.columns else 0)*rate_exchange,
                            "Balance (IDR)": 0 - (df["Marein Share IDR"].sum() if "Marein Share IDR" in df.columns else 0)*rate_exchange,
                            "Check Balance (IDR)": "",
                            "REMARKS": selected_rows.iloc[0]["Remarks"],
                            "STATUS": "POSTED",
                            "CREATED AT": now_wib_naive(),
                            "CREATED BY": pic,
                            "Due Date": due_date,
                            "Subject Email": selected_rows.iloc[0]["Subject Email"],
                            "Email Date": selected_rows.iloc[0]["Email Date"],
                            "CANCELED AT": "-",
                            "CANCELED BY": "-",
                            "CANCEL OF VOUCHER": "-",
                            "CANCEL REASON": "-"
                        }

                    # Upload voucher (selalu CREATE)
                    log_drive_id = find_drive_file(
                        service=service,
                        filename=get_log_filename(int(year), int(month)),
                        # filename="log_produksi.xlsx",
                        parent_id=PERIOD_DRIVE_ID,
                        mime_type="application/vnd.google-apps.spreadsheet"
                    )

                    if not log_drive_id:
                        log_drive_id = create_log_gsheet(
                            service=service,
                            parent_id=PERIOD_DRIVE_ID,
                            filename=get_log_filename(int(year), int(month)),
                            columns=list(log_entry.keys())
                        )

                    sheets_service = init_sheets_service(creds)

                    append_gsheet(
                        service=sheets_service,
                        spreadsheet_id=log_drive_id,
                        row_dict=log_entry
                    )
            
                    upload_dataframe_to_drive(
                        service=service,
                        df=df,
                        template_columns=columns_template,
                        voucher_id=voucher,
                        filename=f"{voucher}.xlsx",
                        folder_id=CEDING_DRIVE_ID,
                        file_type="Voucher"
                    )

                    end_time = time.time()
                    duration = end_time - start_time

                    st.success(f"✅ Voucher berhasil diposting: {voucher} ({int(duration)} seconds)")
                    st.code(voucher)

                except RuntimeError as e:
                        st.error("⛔ Log sedang digunakan user lain. Silakan coba lagi.")
                        st.stop()

                finally:
                    release_drive_lock(service, PERIOD_DRIVE_ID)



# ==========================
# SIMPAN VOUCHER
# ==========================

with tab_post:
    st.subheader("📥 Create Voucher")
    
    columns_template = [
        "No",
        "TL Detail ID",
        "Trans Category",
        "Policy Category",
        "Certificate No",
        "Insured Full Name",
        "Gender",
        "Main Pol No",
        "Main Policy",
        "Pol Holder No",
        "Policy Holder",
        "Birth Date",
        "Age At Issue Date",
        "Term Year",
        "Term Month",
        "Expired Date",
        "Medical",
        "Ced Risk Code",
        "Life Risk Name",
        "K.O.B Code",
        "Ced Product Code",
        "Ced Coverage Code",
        "Ced Product Desc",
        "Ccy Code",
        "Sum Insured",
        "Sum At Risk",
        "Reins Sum Insured",
        "Ced Retention",
        "Reins Sum At Risk",
        "Pay Period Type",
        "Ced EM Rate",
        "Ced ER Rate",
        "Reins Premium",
        "Reins EM Premium",
        "Reins ER Premium",
        "Reins Oth. Premium",
        "Reins Total Premium",
        "Reins Comm",
        "Reins EM Comm",
        "Reins ER Comm",
        "Reins Oth. Comm",
        "Reins Profit Share",
        "Reins Overriding",
        "Reins Broker Fee",
        "Reins Total Comm",
        "Reins Tabarru",
        "Reins Ujrah",
        "Reins Nett Premium",
        "Valuation Date",
        "Terminate Date",
        "TL Detail Remarks",
        "CBY",
        "CBM",
        "COB",
        "Voucher ID",
        "PML ID",
        "References No",
        "Elapse No",
        "Ref Voucher ID"
    ]

    columns_template_claim = [
        "BookYear",
        "BookMonth",
        "CedBookYear",
        "CedBookMonth",
        "Company Name",
        "Policy Holder No",
        "Policy Holder",
        "Certificate No",
        "Insured Name",
        "Birth Date",
        "Age",
        "Gender",
        "Sum Insured IDR",
        "Sum Reinsured IDR",
        "medicalcategory",
        "Product",
        "Coverage Code",
        "ClassOfBusiness",
        "PayPeriodType",
        "Issue Date",
        "Term Year",
        "Term Month",
        "End Date Policy",
        "Claim Date",
        "Claim Register Date",
        "Payment Date",
        "Currency",
        "ExchangeRate",
        "Amount of Claim IDR",
        "Reins Claim IDR",
        "Marein Share IDR",
        "Cause Of Claim",
        "Voucher ID",
        "References No"        
    ]

    # ===== ROW 1 =====
    row1_col1, row1_col2 = st.columns(2)

    with row1_col1:
        reins_type = st.selectbox(
            "Reinsurance Type",
            ["INWARD", "OUTWARD"],
            key="reins_type"
        )

    # row1_col2 sengaja dikosongkan


    # ===== ROW 2 =====
    row2_col1, row2_col2 = st.columns(2)

    with row2_col1:
        department = st.selectbox(
            "Department",
            ["ADMIN", "CLAIM"],
            key="department"
        )

    with row2_col2:
        biz_type = st.selectbox(
            "Biz Type",
            [
                "Kontribusi",
                "Claim",
                "Refund",
                "Alteration",
                "Retur",
                "Revise",
                "Batal",
                "Cancel"
            ],
            key="biz_type"
        )

    uploaded_file = st.file_uploader(
        "Upload Voucher (.xlsx)",
        type=["xlsx"],
        key="upload_post"
    )

    if uploaded_file:
        if reins_type == "INWARD":
            # ==========================
            # READ FILE
            # ==========================
            df = pd.read_excel(uploaded_file)
            original_columns = df.columns.tolist()
            df.columns = df.columns.str.strip().str.lower()

            for col in ["certificate no", "pol holder no"]:
                if col in df.columns:
                    df[col] = df[col].astype(str).str.strip()

            # ==========================
            # VALIDATION
            # ==========================
            errors = validate_voucher(df, st.session_state["biz_type"], st.session_state["reins_type"])

            if errors:
                st.error("❌ VALIDASI GAGAL")
                for err in errors:
                    st.write(f"- {err}")
                st.stop()

            st.success("✅ Validasi berhasil")


            # ==========================
            # PREVIEW + FILTER (DINAMIS)
            # ==========================

            def get_non_empty_columns(df: pd.DataFrame):
                valid_cols = []
                for col in df.columns:
                    series = df[col]
                    non_na = series.dropna()

                    if non_na.empty:
                        continue

                    if series.dtype == "object":
                        if non_na.astype(str).str.strip().ne("").any():
                            valid_cols.append(col)
                    else:
                        valid_cols.append(col)

                return valid_cols


            # ==========================
            # PREVIEW + FILTER
            # ==========================
            with st.expander("📊 Preview Data Voucher", expanded=True):
                if not df.empty:
                    # 1. Batasi jumlah baris agar aplikasi tetap cepat
                    MAX_PREVIEW = 1000
                    total_rows = len(df)
                    preview_df = df.head(MAX_PREVIEW).copy()

                    st.caption(f"Menampilkan {len(preview_df):,} dari {total_rows:,} baris")

                    # 2. SANITIZE & FORMATTING (Sama seperti cara Summary Financial)
                    # Pastikan kolom accounting diformat dengan ribuan dan 2 desimal
                    
                    ACCOUNTING_COLS = [
                        "sum insured", "sum at risk", "reins sum insured", "reins sum at risk",
                        "reins premium", "reins em premium", "reins er premium", "reins total premium",
                        "reins total comm", "reins tabarru", "reins ujrah", "reins nett premium"
                    ]

                    # Buat dictionary formatter untuk kolom yang ada saja
                    format_dict = {}
                    for col in ACCOUNTING_COLS:
                        if col in preview_df.columns:
                            # Pastikan data adalah numerik sebelum diformat
                            preview_df[col] = pd.to_numeric(preview_df[col], errors='coerce').fillna(0)
                            format_dict[col] = "{:,.2f}"

                    # 3. RENDER MENGGUNAKAN ST.DATAFRAME (Sama dengan Summary Financial)
                    try:
                        st.dataframe(
                            preview_df.style.format(format_dict),
                            use_container_width=True,
                            height=450 # Memberikan scrollbar internal jika data banyak
                        )
                    except Exception as e:
                        st.error(f"Gagal menampilkan preview: {e}")
                        st.dataframe(preview_df) # Fallback ke tabel mentah jika styling gagal
                else:
                    st.info("Belum ada data untuk ditampilkan. Silakan upload file terlebih dahulu.")


            # ==========================
            # PERIOD & LOG
            # ==========================
            year = st.session_state["log_period"]["year"]
            month = st.session_state["log_period"]["month"]

            # ==========================
            # DRIVE FOLDER PER PERIODE (STEP 3)
            # ==========================
            drive_folders = get_period_drive_folders(
                year=year,
                month=month,
                root_folder_id=ROOT_DRIVE_FOLDER_ID
            )

            PERIOD_DRIVE_ID = drive_folders["period_id"]


            # ==========================
            # FORM INPUT
            # ==========================


            with st.expander("🧾 Informasi Voucher", expanded=True):

                col1, col2 = st.columns(2)

                with col1:

                    account_with = st.selectbox(
                        "Account With",
                        [
                            "AIA FINANCIAL SYARIAH",
                            "AJS KITABISA (D/H AMANAH GITHA)",
                            "ALLIANZ LIFE SYARIAH",
                            "ALLIANZ LIFE SYARIAH (Health)",
                            "ALLIANZ LIFE SYARIAH (FlexiCare)",
                            "ALLIANZ LIFE SYARIAH (HSCP)",
                            "ALLIANZ LIFE SYARIAH (Individu DMTM)",
                            "ASTRA AVIVA LIFE",
                            "AVRIST ASSURANCE SYARIAH",
                            "AXA FINANCIAL INDONESIA SYARIAH",
                            "AXA MANDIRI FINANCIAL SERVICES SYARIAH",
                            "BNI LIFE SYARIAH",
                            "BRINGIN LIFE SYARIAH",
                            "BUMIPUTERA SYARIAH",
                            "CAPITAL LIFE SYARIAH",
                            "CENTRAL ASIA RAYA SYARIAH",
                            "FWD LIFE INDONESIA SYARIAH",
                            "GENERALI INDONESIA LIFE ASSURANCE SYARIAH",
                            "GREAT EASTERN LIFE SYARIAH",
                            "JASA MITRA ABADI SYARIAH",
                            "MANULIFE INDONESIA SYARIAH",
                            "PFI MEGA LIFE INSURANCE SYARIAH",
                            "PRUDENTIAL LIFE SYARIAH",
                            "PT ASURANSI JIWA SYARIAH BUMIPUTERA",
                            "REASURANSI INTERNATIONAL INDONESIA SYARIAH",
                            "RELIANCE SYARIAH",
                            "SINARMAS SYARIAH",
                            "SUN LIFE SYARIAH",
                            "SYARIAH AL-AMIN",
                            "TAKAFUL KELUARGA",
                            "GENERAL REINSURANCE AG (GEN RE) PLC, SINGAPORE",
                            "HANNOVER RETAKAFUL",
                            "MAREIN SYARIAH",
                            "MUNICH RE RETAKAFUL",
                            "SCOR SE LABUAN BRANCH",
                            "SWISS RE INTL. SE, SINGAPORE (SYARIAH)"
                        ]
                    )

                    cedant_company = st.selectbox(
                        "Cedant Company",
                        [
                            "AIA FINANCIAL SYARIAH",
                            "AJS KITABISA (D/H AMANAH GITHA)",
                            "ALLIANZ LIFE SYARIAH",
                            "ASTRA AVIVA LIFE",
                            "AVRIST ASSURANCE SYARIAH",
                            "AXA FINANCIAL INDONESIA SYARIAH",
                            "AXA MANDIRI FINANCIAL SERVICES SYARIAH",
                            "BNI LIFE SYARIAH",
                            "BRINGIN LIFE SYARIAH",
                            "BUMIPUTERA SYARIAH",
                            "CAPITAL LIFE SYARIAH",
                            "CENTRAL ASIA RAYA SYARIAH",
                            "FWD LIFE INDONESIA SYARIAH",
                            "GENERALI INDONESIA LIFE ASSURANCE SYARIAH",
                            "GREAT EASTERN LIFE SYARIAH",
                            "JASA MITRA ABADI SYARIAH",
                            "MANULIFE INDONESIA SYARIAH",
                            "PANIN DAICHI LIFE SYARIAH",
                            "PFI MEGA LIFE INSURANCE SYARIAH",
                            "PRUDENTIAL LIFE SYARIAH",
                            "PT ASURANSI JIWA SYARIAH BUMIPUTERA",
                            "REASURANSI INTERNATIONAL INDONESIA SYARIAH",
                            "RELIANCE SYARIAH",
                            "SINARMAS SYARIAH",
                            "SUN LIFE SYARIAH",
                            "SYARIAH AL-AMIN",
                            "TAKAFUL KELUARGA",
                            "GENERAL REINSURANCE AG (GEN RE) PLC, SINGAPORE",
                            "HANNOVER RETAKAFUL",
                            "MAREIN SYARIAH",
                            "MUNICH RE RETAKAFUL",
                            "SCOR SE LABUAN BRANCH",
                            "SWISS RE INTL. SE, SINGAPORE (SYARIAH)"
                        ]
                    )


                    pic = st.selectbox("PIC", ["Ardelia", "Buya", "Khansa", "Prabu"])
                    product = st.text_input("Product")

                with col2:
                    years = list(range(2010, year + 1))
                    months = list(range(1, 13))

                    cby = st.selectbox("Ceding Book Year (CBY)", years, index=years.index(year))
                    cbm = st.selectbox("Ceding Book Month (CBM)", months)#, index=months.index(month))
                    #oby = st.selectbox("Our Book Year (OBY)", years, index=years.index(year))
                    #obm = st.selectbox("Our Book Month (OBM)", months)#, index=months.index(month))
                    oby = st.text_input("Our Book Year (OBY)", value=year, disabled=True)
                    obm = st.text_input("Our Book Month (OBM)", value=month, disabled=True)

                kob = st.selectbox(
                    "Kind of Business (KOB)",
                    [
                        "TTY",
                        "FAC"
                    ]
                )

                cob = st.selectbox(
                    "Class of Business (COB)",
                    [
                        "CREDIT GROUP",
                        "HEALTH GROUP",
                        "HEALTH INDIVIDUAL",
                        "LIFE GROUP",
                        "LIFE INDIVIDUAL",
                        "P.A GROUP",
                        "P.A INDIVIDUAL",
                    ]
                )

                mop = st.selectbox(
                    "Mode of Payment (MOP)",
                    ["Monthly", "Quarterly", "Half Yearly", "Yearly", "Single Premium"]
                )

                curr = st.selectbox(
                    "Currency",
                    ["IDR", "USD"]
                )

                subject_email = st.text_area("Subject Email")

                email_date = st.date_input("Email Date",value=date.today())

                remarks = st.text_area("Remarks")


            # ==========================
            # FINANCIAL SUMMARY
            # ==========================
            st.subheader("💰 Ringkasan Finansial")

            if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
                summary_df = pd.DataFrame({
                    "Keterangan": [
                        "Total Contribution",
                        "Commission",
                        "Tabarru",
                        "Ujrah",
                        "Nett Premium"
                    ],
                    "Nilai": [
                        df["reins total premium"].sum(),
                        df["reins total comm"].sum(),
                        df["reins tabarru"].sum(),
                        df["reins ujrah"].sum(),
                        df["reins nett premium"].sum(),
                    ]
                })

            elif biz_type == "Claim":
                summary_df = pd.DataFrame({
                    "Keterangan": [
                        "Amount of Claim IDR",
                        "Reins Claim IDR",
                        "Marein Share IDR"
                    ],
                    "Nilai": [
                        df["amount of claim idr"].sum(),
                        df["reins claim idr"].sum(),
                        df["marein share idr"].sum()
                        ]
                })
                

            st.dataframe(
                summary_df.style.format({"Nilai": "{:,.2f}"}),
                use_container_width=True
            )


            # ==========================
            # POST VOUCHER (LOCKED)
            # ==========================
            if st.button("💾 Simpan Voucher"):
                start_time = time.time()

                if not product.strip() or not remarks.strip() or not subject_email.strip():
                    st.error("Product, Subject Email, dan Remarks wajib diisi")
                    st.stop()

                #lock_path = log_path + ".lock"
                service = get_drive_service()

                with st.spinner("⏳ Menyimpan voucher, mohon tunggu..."):

                    try:
                        service = get_drive_service()                    

                        drive_folders = get_period_drive_folders(
                            year=int(oby),
                            month=int(obm),
                            root_folder_id=ROOT_DRIVE_FOLDER_ID
                        )

                        PERIOD_DRIVE_ID = drive_folders["period_id"]

                        acquire_drive_lock(service, PERIOD_DRIVE_ID)

                        # reload log terbaru setelah lock
                        # if os.path.exists(log_path):
                        #     log_df = pd.read_excel(log_path)
                        # else:
                        #     log_df = pd.DataFrame()

                        voucher, seq_no, file_id = generate_vin_from_drive(
                            service=service,
                            period_folder_id=PERIOD_DRIVE_ID,
                            year=int(oby),
                            month=int(obm),
                            find_drive_file=find_drive_file,
                            biz_type=biz_type
                        )

                        st.write(file_id)

                        ceding_folder_name = normalize_folder_name(account_with)

                        ceding_drive = get_or_create_ceding_folders(
                            service=service,
                            period_folder_id=PERIOD_DRIVE_ID,
                            ceding_name=ceding_folder_name
                        )

                        CEDING_DRIVE_ID = ceding_drive["ceding_id"]


                        # Upload voucher (selalu CREATE)
                        log_drive_id = find_drive_file(
                            service=service,
                            filename=get_log_filename(int(oby), int(obm)),
                            # filename="log_produksi.xlsx",
                            parent_id=PERIOD_DRIVE_ID,
                            mime_type="application/vnd.google-apps.spreadsheet"
                        )

                        rate_exchange = get_exchange_rate(
                            service=service,
                            config_folder_id=CONFIG_FOLDER_ID,
                            currency=curr,
                            month=month
                        )

                        due_date = calculate_due_date(
                            account_with=account_with,
                            year=year,
                            month=month,
                            service=service
                        )


                        if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
                            log_entry = {
                                "Seq No": seq_no,
                                "Department":department,
                                "Biz Type": biz_type,
                                "Voucher No": voucher,
                                "Account With": account_with,
                                "Cedant Company": cedant_company,
                                "PIC": pic,
                                "Product": product,
                                "CBY": cby,
                                "CBM": cbm,
                                "OBY": oby,
                                "OBM": obm,
                                "KOB": kob,
                                "COB": cob,
                                "MOP": mop,
                                "Curr":curr,
                                "Total Contribution": df["reins total premium"].sum(),
                                "Commission": df["reins total comm"].sum(),
                                "Overriding": df["reins overriding"].sum() if "reins overriding" in df.columns else 0,
                                "Total Commission": (df["reins total comm"].sum()) + (df["reins overriding"].sum() if "reins overriding" in df.columns else 0),
                                "Gross Premium Income": df["reins total premium"].sum() - ((df["reins total comm"].sum()) + (df["reins overriding"].sum() if "reins overriding" in df.columns else 0)),
                                "Tabarru": df["reins tabarru"].sum(),
                                "Ujrah": df["reins ujrah"].sum(),
                                "Claim": 0,
                                "Balance": df["reins total premium"].sum() - df["reins total comm"].sum() - (df["reins overriding"].sum() if "reins overriding" in df.columns else 0) - (df["claim"].sum() if "claim" in df.columns else 0),
                                "Check Balance": "",
                                "Rate Exchange": rate_exchange,
                                "Kontribusi (IDR)": (df["reins total premium"].sum())*rate_exchange,
                                "Commission (IDR)": (df["reins total comm"].sum())*rate_exchange,
                                "Overiding (IDR)": (df["reins overriding"].sum() if "reins overriding" in df.columns else 0)*rate_exchange,
                                "Total Commission (IDR)": ((df["reins total comm"].sum()) + (df["reins overriding"].sum() if "reins overriding" in df.columns else 0))*rate_exchange,
                                "Gross Premium Income (IDR)": (df["reins total premium"].sum() - ((df["reins total comm"].sum()) + (df["reins overriding"].sum() if "reins overriding" in df.columns else 0)))*rate_exchange,
                                "Tabarru (IDR)": (df["reins tabarru"].sum())*rate_exchange,
                                "Ujrah (IDR)": (df["reins ujrah"].sum())*rate_exchange,
                                "Claim (IDR)": 0,
                                "Balance (IDR)": (df["reins total premium"].sum() - df["reins total comm"].sum() - (df["reins overriding"].sum() if "reins overriding" in df.columns else 0) - (df["claim"].sum() if "claim" in df.columns else 0))*rate_exchange,
                                "Check Balance (IDR)":"",
                                "REMARKS": remarks,
                                "STATUS": "POSTED",
                                #"ENTRY_TYPE": entry_type,
                                "CREATED AT": now_wib_naive(),
                                "CREATED BY": pic,
                                "Due Date": due_date,
                                "Subject Email": subject_email,
                                "Email Date": email_date,
                                "CANCELED AT": "-",
                                "CANCELED BY": "-",
                                "CANCEL OF VOUCHER": "-",
                                "CANCEL REASON":"-"
                            }

                        elif biz_type == "Claim":
                            log_entry = {
                                "Seq No": seq_no,
                                "Department":department,
                                "Biz Type": biz_type,
                                "Voucher No": voucher,
                                "Account With": account_with,
                                "Cedant Company": cedant_company,
                                "PIC": pic,
                                "Product": product,
                                "CBY": cby,
                                "CBM": cbm,
                                "OBY": oby,
                                "OBM": obm,
                                "KOB": kob,
                                "COB": cob,
                                "MOP": mop,
                                "Curr":curr,
                                "Total Contribution": 0,
                                "Commission": 0,
                                "Overriding": 0,
                                "Total Commission": 0,
                                "Gross Premium Income": 0,
                                "Tabarru": 0,
                                "Ujrah": 0,
                                "Claim": df["marein share idr"].sum(),
                                "Balance": 0 - (df["marein share idr"].sum() if "marein share idr" in df.columns else 0),
                                "Check Balance": "",
                                "Rate Exchange": rate_exchange,
                                "Kontribusi (IDR)": 0,
                                "Commission (IDR)": 0,
                                "Overiding (IDR)": 0,
                                "Total Commission (IDR)": 0,
                                "Gross Premium Income (IDR)": 0,
                                "Tabarru (IDR)": 0,
                                "Ujrah (IDR)": 0,
                                "Claim (IDR)": (df["marein share idr"].sum() if "marein share idr" in df.columns else 0)*rate_exchange,
                                "Balance (IDR)": 0 - (df["marein share idr"].sum() if "marein share idr" in df.columns else 0)*rate_exchange,
                                "Check Balance (IDR)": "",
                                "REMARKS": remarks,
                                "STATUS": "POSTED",
                                #"ENTRY_TYPE": entry_type,
                                "CREATED AT": now_wib_naive(),
                                "CREATED BY": pic,
                                "Due Date": due_date,
                                "Subject Email": subject_email,
                                "Email Date": email_date,
                                "CANCELED AT": "-",
                                "CANCELED BY": "-",
                                "CANCEL OF VOUCHER": "-",
                                "CANCEL REASON": "-"
                            }

                        # log_drive_id = find_drive_file(
                        #     service=service,
                        #     filename=get_log_filename(int(oby), int(obm)),
                        #     parent_id=PERIOD_DRIVE_ID,
                        #     mime_type="application/vnd.google-apps.spreadsheet"
                        # )

                        if not log_drive_id:
                            log_drive_id = create_log_gsheet(
                                service=service,
                                parent_id=PERIOD_DRIVE_ID,
                                filename=get_log_filename(int(oby), int(obm)),
                                columns=list(log_entry.keys())
                            )

                        sheets_service = init_sheets_service(creds)

                        append_gsheet(
                            service=sheets_service,
                            spreadsheet_id=log_drive_id,
                            row_dict=log_entry
                        )
                
                        upload_dataframe_to_drive(
                            service=service,
                            df=df,
                            template_columns=columns_template,
                            voucher_id=voucher,
                            filename=f"{voucher}.xlsx",
                            folder_id=CEDING_DRIVE_ID,
                            file_type="Voucher"
                        )

                        end_time = time.time()
                        duration = end_time - start_time

                        st.success(f"✅ Voucher berhasil diposting: {voucher} ({int(duration)} seconds)")
                        st.code(voucher)

                    except RuntimeError as e:
                            st.error("⛔ Log sedang digunakan user lain. Silakan coba lagi.")
                            st.stop()

                    finally:
                        release_drive_lock(service, PERIOD_DRIVE_ID)


        elif reins_type == "OUTWARD":
            # ==========================
            # READ FILE
            # ==========================
            df = pd.read_excel(uploaded_file)
            original_columns = df.columns.tolist()
            df.columns = df.columns.str.strip().str.lower()

            for col in ["certificate no", "pol holder no"]:
                if col in df.columns:
                    df[col] = df[col].astype(str).str.strip()

            # ==========================
            # VALIDATION
            # ==========================
            errors = validate_voucher(df, st.session_state["biz_type"], st.session_state["reins_type"])

            if errors:
                st.error("❌ VALIDASI GAGAL")
                for err in errors:
                    st.write(f"- {err}")
                st.stop()

            st.success("✅ Validasi berhasil")


            # ==========================
            # PREVIEW + FILTER (DINAMIS)
            # ==========================

            def get_non_empty_columns(df: pd.DataFrame):
                valid_cols = []
                for col in df.columns:
                    series = df[col]
                    non_na = series.dropna()

                    if non_na.empty:
                        continue

                    if series.dtype == "object":
                        if non_na.astype(str).str.strip().ne("").any():
                            valid_cols.append(col)
                    else:
                        valid_cols.append(col)

                return valid_cols


            # ==========================
            # PREVIEW + FILTER
            # ==========================
            with st.expander("📊 Preview Data Voucher", expanded=True):

                filtered_df = df.copy()

                MAX_PREVIEW_ROWS = 2000

                display_df = filtered_df.copy()
                total_rows = len(display_df)

                # # ==========================
                # # LIMIT PREVIEW ROWS
                # # ==========================

                if total_rows > MAX_PREVIEW_ROWS:
                    st.warning(
                        f"⚠️ Data sangat besar ({total_rows:,} baris). "
                        f"Hanya menampilkan {MAX_PREVIEW_ROWS:,} baris pertama untuk preview."
                    )
                    preview_df = display_df.head(MAX_PREVIEW_ROWS)
                else:
                    preview_df = display_df

                st.caption(f"Menampilkan {len(preview_df):,} dari {total_rows:,} baris")

                # ==========================
                # SANITIZE DATA (ANTI-CRASH)
                # ==========================

                preview_df = preview_df.copy()

                for col in preview_df.columns:
                    # Convert datetime (including timezone) to string
                    if pd.api.types.is_datetime64_any_dtype(preview_df[col]):
                        preview_df[col] = preview_df[col].astype(str)

                    # Convert Period
                    elif "period" in str(preview_df[col].dtype):
                        preview_df[col] = preview_df[col].astype(str)

                    # Convert object that might contain mixed types
                    elif preview_df[col].dtype == "object":
                        preview_df[col] = preview_df[col].astype(str)

                preview_df = preview_df.fillna("")

                # ==========================
                # GRID BUILDER
                # ==========================

                gb = GridOptionsBuilder.from_dataframe(preview_df)

                gb.configure_default_column(
                    filter=True,
                    sortable=True,
                    resizable=True,
                    minWidth=120,
                    flex=0
                )

                # ==========================
                # ACCOUNTING FORMATTER (INTERNATIONAL)
                # ==========================

                accounting_formatter = JsCode("""
                function(params) {
                    if (params.value == null || params.value === '') return '';

                    let value = Number(params.value);

                    let formatted = Math.abs(value).toLocaleString('en-US', {
                        minimumFractionDigits: 2,
                        maximumFractionDigits: 2
                    });

                    if (value < 0) {
                        return '(' + formatted + ')';
                    }

                    return formatted;
                }
                """)

                ACCOUNTING_COLS = [
                    "sum insured",
                    "sum at risk",
                    "reins sum insured",
                    "reins sum at risk",
                    "retro sum insured",
                    "retro sum at risk",
                    "retro premium",
                    "retro em premium",
                    "retro er premium",
                    "retro oth premium",
                    "retro total premium",
                    "retro comm",
                    "retro em comm",
                    "retro er comm",
                    "retro oth comm",
                    "retro profit share",
                    "retro overriding",
                    "retro total comm",
                    "retro tabarru",
                    "retro ujrah",
                    "retro nett premium"
                ]

                for col in ACCOUNTING_COLS:
                    if col in preview_df.columns:
                        gb.configure_column(
                            col,
                            type=["numericColumn"],
                            valueFormatter=accounting_formatter,
                            cellStyle={"textAlign": "right"}
                        )

                # ==========================
                # GRID OPTIONS
                # ==========================

                gb.configure_pagination(
                    paginationAutoPageSize=False,
                    paginationPageSize=50
                )

                gb.configure_grid_options(
                    headerHeight=42,
                    rowHeight=36,
                    domLayout="normal",
                    suppressHorizontalScroll=False,
                    onFirstDataRendered="""
                    function(params) {
                        const allColumnIds = [];
                        params.columnApi.getAllColumns().forEach(function(col) {
                            allColumnIds.push(col.getId());
                        });
                        params.columnApi.autoSizeColumns(allColumnIds, false);
                    }
                    """
                )

                grid_options = gb.build()

                # ==========================
                # CUSTOM CSS
                # ==========================

                custom_css = {

                    # ---------- WRAPPER ----------
                    ".ag-root-wrapper": {
                        "background-color": "var(--secondary-background-color)",
                        # "border": "1px solid var(--primary-color)",
                        "border-radius": "12px",
                    },

                    ".ag-center-cols-viewport": {
                        "background-color": "var(--secondary-background-color)",
                    },

                    ".ag-body-viewport": {
                        "background-color": "var(--secondary-background-color)",
                    },

                    ".ag-center-cols-container": {
                        "background-color": "var(--secondary-background-color)",
                    },

                    # ---------- HEADER ----------
                    ".ag-header": {
                        "background-color": "var(--background-color)",
                        "color": "var(--text-color)",
                        "font-weight": "600",
                        "font-size": "13px",
                        # "border-bottom": "2px solid var(--primary-color)"
                    },

                    ".ag-header-cell": {
                        "padding-top": "8px",
                        "padding-bottom": "8px",
                        "border-right": "1px solid var(--secondary-background-color)"
                    },

                    ".ag-header-cell-label": {
                        "display": "flex",
                        "align-items": "center",
                        "justify-content": "center",
                        "width": "100%"
                    },

                    ".ag-header-cell-text": {
                        "flex-grow": "1",
                        "text-align": "center",
                        "text-transform": "capitalize"
                    },

                    # ---------- BODY ----------
                    ".ag-row": {
                        "background-color": "var(--secondary-background-color)",
                        "color": "var(--text-color)",
                        "border-bottom": "1px solid var(--background-color)"
                    },

                    ".ag-row-hover": {
                        "background-color": "rgba(128,128,128,0.15)",
                    },

                    ".ag-cell": {
                        "border-color": "var(--background-color)",
                        "border-right": "1px solid var(--background-color)",
                        "border-bottom": "1px solid var(--background-color)"
                    },

                    # ---------- PAGINATION ----------
                    ".ag-paging-panel": {
                        "background-color": "var(--secondary-background-color)",
                        "color": "var(--text-color)",
                    },

                    # ---------- ICON STYLE ----------
                    ".ag-icon": {
                        "font-size": "11px",
                        "opacity": "0.8"
                    }

                }


                # ==========================
                # RENDER GRID
                # ==========================

                AgGrid(
                    preview_df,
                    gridOptions=grid_options,
                    height=600,
                    theme="dark",
                    custom_css=custom_css,
                    allow_unsafe_jscode=True
                )


            # ==========================
            # PERIOD & LOG
            # ==========================
            year = st.session_state["log_period"]["year"]
            month = st.session_state["log_period"]["month"]


            # ==========================
            # DRIVE FOLDER PER PERIODE (STEP 3)
            # ==========================
            drive_folders = get_period_drive_folders(
                year=year,
                month=month,
                root_folder_id=ROOT_DRIVE_FOLDER_ID
            )

            PERIOD_DRIVE_ID = drive_folders["period_id"]


            # ==========================
            # FORM INPUT
            # ==========================

            with st.expander("🧾 Informasi Voucher", expanded=True):

                col1, col2 = st.columns(2)

                with col1:
                    retro_type = st.selectbox(
                        "Retro Type",
                        [
                            "Sp Program",
                            "Sp Arrangement",
                            "Panel"
                        ]
                    )

                    inward_vin = st.text_input("Inward Vin Ref", value=str(df.loc[0, "inw vouc id"]))

                    account_with = st.selectbox(
                        "Account With",
                        [
                            "GENERAL REINSURANCE AG (GEN RE) PLC, SINGAPORE",
                            "HANNOVER RETAKAFUL",
                            "MAREIN SYARIAH",
                            "MUNICH RE RETAKAFUL",
                            "SCOR SE LABUAN BRANCH",
                            "SWISS RE INTL. SE, SINGAPORE (SYARIAH)",
                            "REASURANSI INTERNATIONAL INDONESIA SYARIAH",
                            "REASURANSI NASIONAL INDONESIA SYARIAH"
                        ]

                    )
                    

                    cedant_company = st.selectbox(
                        "Cedant Company",
                        [
                            "AIA FINANCIAL SYARIAH",
                            "AJS KITABISA (D/H AMANAH GITHA)",
                            "ALLIANZ LIFE SYARIAH",
                            "ASTRA AVIVA LIFE",
                            "AVRIST ASSURANCE SYARIAH",
                            "AXA FINANCIAL INDONESIA SYARIAH",
                            "AXA MANDIRI FINANCIAL SERVICES SYARIAH",
                            "BNI LIFE SYARIAH",
                            "BRINGIN LIFE SYARIAH",
                            "BUMIPUTERA SYARIAH",
                            "CAPITAL LIFE SYARIAH",
                            "CENTRAL ASIA RAYA SYARIAH",
                            "FWD LIFE INDONESIA SYARIAH",
                            "GENERALI INDONESIA LIFE ASSURANCE SYARIAH",
                            "GREAT EASTERN LIFE SYARIAH",
                            "JASA MITRA ABADI SYARIAH",
                            "MANULIFE INDONESIA SYARIAH",
                            "PANIN DAICHI LIFE SYARIAH",
                            "PFI MEGA LIFE INSURANCE SYARIAH",
                            "PRUDENTIAL LIFE SYARIAH",
                            "PT ASURANSI JIWA SYARIAH BUMIPUTERA",
                            "REASURANSI INTERNATIONAL INDONESIA SYARIAH",
                            "RELIANCE SYARIAH",
                            "SINARMAS SYARIAH",
                            "SUN LIFE SYARIAH",
                            "SYARIAH AL-AMIN",
                            "TAKAFUL KELUARGA",
                            "GENERAL REINSURANCE AG (GEN RE) PLC, SINGAPORE",
                            "HANNOVER RETAKAFUL",
                            "MAREIN SYARIAH",
                            "MUNICH RE RETAKAFUL",
                            "SCOR SE LABUAN BRANCH",
                            "SWISS RE INTL. SE, SINGAPORE (SYARIAH)"
                        ]
                    )


                    pic = st.selectbox("PIC", ["Ardelia", "Buya", "Khansa", "Prabu"])
                    product = st.text_input("Product")

                with col2:
                    years = list(range(2015, year + 1))
                    months = list(range(1, 13))

                    cby = st.selectbox("Ceding Book Year (CBY)", years, index=years.index(year))
                    cbm = st.selectbox("Ceding Book Month (CBM)", months)#, index=months.index(month))
                    #oby = st.selectbox("Our Book Year (OBY)", years, index=years.index(year))
                    #obm = st.selectbox("Our Book Month (OBM)", months)#, index=months.index(month))
                    oby = st.text_input("Our Book Year (OBY)", value=year, disabled=True)
                    obm = st.text_input("Our Book Month (OBM)", value=month, disabled=True)

                kob = st.selectbox(
                    "Kind of Business (KOB)",
                    [
                        "TTY",
                        "FAC"
                    ]
                )

                cob = st.selectbox(
                    "Class of Business (COB)",
                    [
                        "CREDIT GROUP",
                        "HEALTH GROUP",
                        "HEALTH INDIVIDUAL",
                        "LIFE GROUP",
                        "LIFE INDIVIDUAL",
                        "P.A GROUP",
                        "P.A INDIVIDUAL",
                    ]
                )

                mop = st.selectbox(
                    "Mode of Payment (MOP)",
                    ["Monthly", "Quarterly", "Half Yearly", "Yearly", "Single Premium"]
                )

                curr = st.selectbox(
                    "Currency",
                    ["IDR", "USD"]
                )

                remarks = st.text_area("Remarks")


            # ==========================
            # FINANCIAL SUMMARY
            # ==========================
            st.subheader("💰 Ringkasan Finansial")

            if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
                summary_df = pd.DataFrame({
                    "Keterangan": [
                        "Total Contribution",
                        "Commission",
                        "Overriding",
                        "Tabarru",
                        "Ujrah",
                        "Nett Premium"
                    ],
                    "Nilai": [
                        df["retro total premium"].sum(),
                        df["retro total comm"].sum(),
                        df["retro overriding"].sum(),
                        df["retro tabarru"].sum(),
                        df["retro ujrah"].sum(),
                        df["retro nett premium"].sum(),
                    ]
                })

            elif biz_type == "Claim":
                summary_df = pd.DataFrame({
                    "Keterangan": [
                        "Reins Claim",
                        "Retro share"
                    ],
                    "Nilai": [
                        df["reins claim"].sum(),
                        df["your share"].sum()
                        ]
                })
                

            st.dataframe(
                summary_df.style.format({"Nilai": "{:,.2f}"}),
                use_container_width=True
            )


            # ==========================
            # POST VOUCHER (LOCKED)
            # ==========================
            if st.button("💾 Simpan Voucher"):
                start_time = time.time()

                if not inward_vin.strip() or not product.strip() or not remarks.strip():
                    st.error("Inward Vin Ref, Product, dan Remarks wajib diisi")
                    st.stop()

                #lock_path = log_path + ".lock"
                service = get_drive_service()

                with st.spinner("⏳ Menyimpan voucher, mohon tunggu..."):

                    try:
                        service = get_drive_service()                    

                        drive_folders = get_period_drive_folders(
                            year=int(oby),
                            month=int(obm),
                            root_folder_id=ROOT_DRIVE_FOLDER_ID
                        )

                        PERIOD_DRIVE_ID = drive_folders["period_id"]

                        
                        outward_drive_folder = get_or_create_outward_folders(
                            service=service,
                            period_folder_id=PERIOD_DRIVE_ID,
                        )

                        OUTWARD_DRIVE_ID = outward_drive_folder["outward_id"]

                        acquire_drive_lock(service, OUTWARD_DRIVE_ID)

                        # reload log terbaru setelah lock
                        # if os.path.exists(log_path):
                        #     log_df = pd.read_excel(log_path)
                        # else:
                        #     log_df = pd.DataFrame()

                        voucher, seq_no = generate_vou_from_drive(
                            service=service,
                            outward_folder_id=OUTWARD_DRIVE_ID,
                            year=int(oby),
                            month=int(obm),
                            find_drive_file=find_drive_file,
                            biz_type=biz_type
                        )

                        st.success("Voucher berhasil dibuat!")
                            
                        ceding_folder_name = normalize_folder_name(account_with)

                        ceding_drive = get_or_create_ceding_folders(
                            service=service,
                            period_folder_id=OUTWARD_DRIVE_ID,
                            ceding_name=ceding_folder_name
                        )

                        CEDING_DRIVE_ID = ceding_drive["ceding_id"]


                        # Upload voucher (selalu CREATE)
                        log_drive_id = find_drive_file(
                            service=service,
                            filename=get_log_filename_outward(int(oby), int(obm)),
                            # filename="log_produksi.xlsx",
                            parent_id=OUTWARD_DRIVE_ID,
                            mime_type="application/vnd.google-apps.spreadsheet"
                        )

                        rate_exchange = get_exchange_rate(
                            service=service,
                            config_folder_id=CONFIG_FOLDER_ID,
                            currency=curr,
                            month=month
                        )

                        due_date = calculate_due_date(
                            account_with=account_with,
                            year=year,
                            month=month,
                            service=service
                        )


                        if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
                            log_entry = {
                                "Seq No": seq_no,
                                "Department":department,
                                "Biz Type": biz_type,
                                "Retro Type": retro_type,
                                "Inward VIN Ref": inward_vin,
                                "Voucher No": voucher,
                                "Account With": account_with,
                                "Cedant Company": cedant_company,
                                "PIC": pic,
                                "Product": product,
                                "CBY": cby,
                                "CBM": cbm,
                                "OBY": oby,
                                "OBM": obm,
                                "KOB": kob,
                                "COB": cob,
                                "MOP": mop,
                                "Curr":curr,
                                "Total Contribution": df["retro total premium"].sum(),
                                "Commission": df["retro total comm"].sum(),
                                "reins overriding": df["retro overriding"].sum() if "retro overriding" in df.columns else 0,
                                "Total Commission": (df["retro total comm"].sum()) + (df["retro overriding"].sum() if "retro overriding" in df.columns else 0),
                                "Gross Premium Income": df["retro total premium"].sum() - ((df["retro total comm"].sum()) + (df["retro overriding"].sum() if "retro overriding" in df.columns else 0)),
                                "Tabarru": df["retro tabarru"].sum(),
                                "Ujrah": df["retro ujrah"].sum(),
                                "Ujrah Spc": 0,
                                "Claim": 0,
                                "Balance": df["retro total premium"].sum() - df["retro total comm"].sum() - (df["claim"].sum() if "claim" in df.columns else 0),
                                "Check Balance": "",
                                "Rate Exchange": rate_exchange,
                                "Kontribusi (IDR)": (df["retro total premium"].sum())*rate_exchange,
                                "Commission (IDR)": (df["retro total comm"].sum())*rate_exchange,
                                "Overiding (IDR)": (df["retro overriding"].sum() if "retro overriding" in df.columns else 0)*rate_exchange,
                                "Total Commission (IDR)": ((df["retro total comm"].sum()) + (df["retro overriding"].sum() if "retro overriding" in df.columns else 0))*rate_exchange,
                                "Gross Premium Income (IDR)": (df["retro total premium"].sum() - ((df["retro total comm"].sum()) + (df["retro overriding"].sum() if "retro overriding" in df.columns else 0)))*rate_exchange,
                                "Tabarru (IDR)": (df["retro tabarru"].sum())*rate_exchange,
                                "Ujrah (IDR)": (df["retro ujrah"].sum())*rate_exchange,
                                "Ujrah Spc (IDR)": 0,
                                "Claim (IDR)": 0,
                                "Balance (IDR)": (df["retro total premium"].sum() - df["retro total comm"].sum() - (df["claim"].sum() if "claim" in df.columns else 0))*rate_exchange,
                                "Check Balance (IDR)":"",
                                "REMARKS": remarks,
                                "STATUS": "POSTED",
                                #"ENTRY_TYPE": entry_type,
                                "CREATED AT": now_wib_naive(),
                                "CREATED BY": pic,
                                "Due Date": due_date,
                                "CANCELED AT": "-",
                                "CANCELED BY": "-",
                                "CANCEL OF VOUCHER": "-",
                                "CANCEL REASON":"-"
                            }

                        elif biz_type == "Claim":
                            log_entry = {
                                "Seq No": seq_no,
                                "Department":department,
                                "Biz Type": biz_type,
                                "Retro Type": retro_type,
                                "Inward VIN Ref": inward_vin,
                                "Voucher No": voucher,
                                "Account With": account_with,
                                "Cedant Company": cedant_company,
                                "PIC": pic,
                                "Product": product,
                                "CBY": cby,
                                "CBM": cbm,
                                "OBY": oby,
                                "OBM": obm,
                                "KOB": kob,
                                "COB": cob,
                                "MOP": mop,
                                "Curr":curr,
                                "Total Contribution": 0,
                                "Commission": 0,
                                "Overriding": 0,
                                "Total Commission": 0,
                                "Gross Premium Income": 0,
                                "Tabarru": 0,
                                "Ujrah": 0,
                                "Ujrah Spc": "",
                                "Claim": df["your share"].sum(),
                                "Balance": 0 - (df["your share"].sum() if "your share" in df.columns else 0),
                                "Check Balance": "",
                                "Rate Exchange": rate_exchange,
                                "Kontribusi (IDR)": 0,
                                "Commission (IDR)": 0,
                                "Overiding (IDR)": 0,
                                "Total Commission (IDR)": 0,
                                "Gross Premium Income (IDR)": 0,
                                "Tabarru (IDR)": 0,
                                "Ujrah (IDR)": 0,
                                "Ujrah Spc (IDR)": "",
                                "Claim (IDR)": (df["your share"].sum() if "your share" in df.columns else 0)*rate_exchange,
                                "Balance (IDR)": 0 - (df["your share"].sum() if "your share" in df.columns else 0)*rate_exchange,
                                "Check Balance (IDR)": "",
                                "REMARKS": remarks,
                                "STATUS": "POSTED",
                                #"ENTRY_TYPE": entry_type,
                                "CREATED AT": now_wib_naive(),
                                "CREATED BY": pic,
                                "Due Date": due_date,
                                "CANCELED AT": "-",
                                "CANCELED BY": "-",
                                "CANCEL OF VOUCHER": "-",
                                "CANCEL REASON": "-"
                            }

                        # log_drive_id = find_drive_file(
                        #     service=service,
                        #     filename=get_log_filename(int(oby), int(obm)),
                        #     parent_id=PERIOD_DRIVE_ID,
                        #     mime_type="application/vnd.google-apps.spreadsheet"
                        # )

                        if not log_drive_id:
                            log_drive_id = create_log_gsheet(
                                service=service,
                                parent_id=OUTWARD_DRIVE_ID,
                                filename=get_log_filename_outward(int(oby), int(obm)),
                                columns=list(log_entry.keys())
                            )

                        sheets_service = init_sheets_service(creds)

                        append_gsheet(
                            service=sheets_service,
                            spreadsheet_id=log_drive_id,
                            row_dict=log_entry
                        )

                        upload_dataframe_to_drive_outward(
                            service=service,
                            df=df,
                            original_columns=original_columns,
                            voucher_id=voucher,
                            filename=f"{voucher}.xlsx",
                            folder_id=CEDING_DRIVE_ID,
                            biz_type=biz_type
                        )

                        end_time = time.time()
                        duration = end_time - start_time

                        st.success(f"✅ Voucher berhasil diposting: {voucher} ({int(duration)} seconds)")
                        st.code(voucher)

                    except RuntimeError as e:
                            st.error("⛔ Log sedang digunakan user lain. Silakan coba lagi.")
                            st.stop()

                    finally:
                        release_drive_lock(service, OUTWARD_DRIVE_ID)


with tab_update:
    st.subheader("🔄 Update  Voucher")
    PROD_PERIOD_ID = None
    NOW_PERIOD_ID = None

    service = get_drive_service()

    # ==============================
    # PILIH PERIODE PRODUKSI
    # ==============================
    action_type = st.radio(
        "Pilih Opsi",
        ["Split Voucher", 
         "Delete Voucher", "Cancel Voucher"],
        key="action_type"
    )

    year = st.session_state["log_period"]["year"]
    month = st.session_state["log_period"]["month"]

    prod_year = int(year)
    prod_month = int(month)

    years = list(range(2026, datetime.now().year + 1))
    months = list(range(1, 13))

    # --- 0. INISIALISASI SERVICE ---
    # Pastikan fungsi ini didefinisikan di drive_utils.py atau di atas
    drive_service = get_drive_service()
    sheets_service = init_sheets_service(creds)

    if action_type == "Split Voucher":
        # 1. SETUP PARAMETER AWAL (Di luar try agar finally bisa mengaksesnya)
        PERIOD_DRIVE_ID = None
        df_posted = pd.DataFrame() # Default kosong agar tidak NameError
        
        # Ambil Folder ID berdasarkan Tahun/Bulan
        drive_folders = get_period_drive_folders(
            year=int(year),
            month=int(month),
            root_folder_id=ROOT_DRIVE_FOLDER_ID
        )
        PERIOD_DRIVE_ID = drive_folders.get("period_id")

        if not PERIOD_DRIVE_ID:
            st.error("Folder periode tidak ditemukan di Drive.")
            st.stop()

        try:
            # 2. LOCKING (Gunakan Drive Service)
            acquire_drive_lock(drive_service, PERIOD_DRIVE_ID)

            # 3. MENCARI & MEMBACA DATA
            # Cari Folder PML
            pml_drive_id = get_or_create_folder(
                service=drive_service,
                folder_name="Folder PML",
                parent_id=PERIOD_DRIVE_ID
            )

            # Cari File Log Spreadsheet
            log_pml_drive_id = find_drive_file(
                service=drive_service,
                filename=get_log_pml_filename(int(year), int(month)),
                parent_id=pml_drive_id,
                mime_type="application/vnd.google-apps.spreadsheet"
            )

            if log_pml_drive_id:
                # PENTING: Baca isi pakai SHEETS SERVICE
                # Pastikan load_log_from_gsheet menggunakan range "'Nama Sheet'!A:Z" (dengan kutip satu)
                df_log = load_log_from_gsheet(sheets_service, log_pml_drive_id)
                
                if not df_log.empty and 'STATUS' in df_log.columns:
                    # Filter hanya yang POSTED
                    df_posted = df_log[df_log['STATUS'] == 'POSTED'].copy()
                else:
                    st.warning("Data Log kosong atau kolom STATUS tidak ditemukan.")
            else:
                st.error("File Log Spreadsheet tidak ditemukan di Folder PML.")

        except Exception as e:
            st.error(f"Terjadi kesalahan saat memproses data: {e}")
        
        finally:
            # 4. RELEASE LOCK (Selalu dijalankan meskipun error di atas)
            if PERIOD_DRIVE_ID:
                release_drive_lock(drive_service, PERIOD_DRIVE_ID)

        # --- 5. RENDER UI DENGAN PEMILIHAN (CHECKBOX) ---
        st.markdown("### 📋 Pilih Data PML untuk Di-Split")
        st.info("Centang pada kolom **'Pilih'** untuk menentukan baris yang akan diproses.")

        if not df_posted.empty:
            # 1. Tambahkan kolom Checkbox (default False)
            df_to_edit = df_posted.copy()
            df_to_edit.insert(0, "Pilih", False)

            # 2. Konfigurasi Tampilan Kolom (Formatting)
            # Pastikan kolom angka tetap rapi
            format_dict = {"Total Contribution": "{:,.0f}"}

            # 3. Gunakan st.data_editor agar bisa dicentang
            edited_df = st.data_editor(
                df_to_edit,
                column_config={
                    "Pilih": st.column_config.CheckboxColumn(
                        "Pilih",
                        help="Pilih baris ini untuk di-split",
                        default=False,
                    ),
                    # Kunci kolom lain agar tidak bisa diedit oleh user
                    "PML ID": st.column_config.Column(disabled=True),
                    "STATUS": st.column_config.Column(disabled=True),
                    "Product": st.column_config.Column(disabled=True),
                    "Total Contribution": st.column_config.NumberColumn(
                        "Total Contribution", format="#,##0", disabled=True
                    ),
                },
                disabled=["No", "PML ID", "STATUS", "Product", "Total Contribution"],
                hide_index=True,
                use_container_width=True,
            )

            # 4. Filter Baris yang Dipilih
            selected_rows = edited_df[edited_df["Pilih"] == True]

            # 5. Logika Validasi Pilihan
            if len(selected_rows) > 1:
                st.warning("⚠️ Anda memilih lebih dari 1 baris. Harap pilih **satu baris saja** untuk proses split.")
            
            elif len(selected_rows) == 1:
                selected_pml_id = selected_rows.iloc[0]["PML ID"]
                st.success(f"✅ Baris terpilih: **{selected_pml_id}**")

                # 1. Mengambil data asli dari file PML yang sudah di-upload sebelumnya
                pml_drive = get_or_create_folder(
                    service=service,
                    folder_name="Folder PML",
                    parent_id=PERIOD_DRIVE_ID
                )

                PML_DRIVE_ID = pml_drive


                pml_file_id = find_drive_file(
                    service=service,
                    filename=selected_pml_id,
                    parent_id=PML_DRIVE_ID
                )

                service = get_drive_service()

                # Folder
                pml_drive = get_or_create_folder(
                    service=service,
                    folder_name="Folder PML",
                    parent_id=PERIOD_DRIVE_ID
                )

                PML_DRIVE_ID = pml_drive

                # Cari file
                pml_file_id = find_drive_file(
                    service=service,
                    filename=selected_pml_id,
                    parent_id=PML_DRIVE_ID
                )

                if not pml_file_id:
                    st.error("File PML tidak ditemukan")
                    st.stop()


                # ==========================
                # LOAD FILE
                # ==========================
                file_stream = download_file_from_drive(service, pml_file_id)
                df = pd.read_excel(file_stream)

                st.write("Preview Data:", df.head())

                # ==========================
                # SELECT MULTI COLUMN
                # ==========================
                selected_columns = st.multiselect(
                    "Pilih kolom untuk split (bisa lebih dari 1)",
                    df.columns.tolist()
                )

                if not selected_columns:
                    st.warning("Pilih minimal 1 kolom untuk split")
                    st.stop()

                # ==========================
                # PROSES SPLIT
                # ==========================
                if st.button(f"Proses Split untuk {selected_pml_id}", type="primary"):

                    # 🔥 PREVENT DOUBLE RUN
                    if st.session_state.is_processing_split:
                        st.warning("⏳ Proses masih berjalan...")
                        st.stop()

                    st.session_state.is_processing_split = True

                    with st.spinner(f"⏳ Split {selected_pml_id} sedang diproses..."):

                        progress_bar = st.progress(0)
                        status_text = st.empty()

                        acquire_drive_lock(service, PERIOD_DRIVE_ID)

                        try:
                            sheets_service = init_sheets_service(creds)

                            log_pml_drive_id = find_drive_file(
                                service=service,
                                filename=get_log_pml_filename(int(year), int(month)),
                                parent_id=PML_DRIVE_ID,
                                mime_type="application/vnd.google-apps.spreadsheet"
                            )

                            base_info = {
                                "department": selected_rows.iloc[0]["Department"],
                                "account_with": selected_rows.iloc[0]["Account With"],
                                "cedant_company": selected_rows.iloc[0]["Cedant Company"],
                                "pic": selected_rows.iloc[0]["PIC"],
                                "curr": selected_rows.iloc[0]["Curr"],
                                "subject_email": selected_rows.iloc[0]["Subject Email"],
                                "email_date": selected_rows.iloc[0]["Email Date"],
                                "source_pml": selected_rows.iloc[0]["PML ID"]
                            }

                            if selected_rows.iloc[0]["Biz Type"] in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
                                results = split_upload_with_log(
                                    service=service,
                                    sheets_service=sheets_service,
                                    df=df,
                                    split_columns=selected_columns,
                                    period_drive_id=PERIOD_DRIVE_ID,
                                    pml_folder_id=PML_DRIVE_ID,
                                    log_pml_drive_id=log_pml_drive_id,
                                    year=int(year),
                                    month=int(month),
                                    biz_type=selected_rows.iloc[0]["Biz Type"],
                                    base_info=base_info,
                                    columns_template=columns_template,
                                    progress_bar=progress_bar,
                                    status_text=status_text
                                )

                            elif selected_rows.iloc[0]["Biz Type"] == "Claim":
                                    results = split_upload_with_log(
                                    service=service,
                                    sheets_service=sheets_service,
                                    df=df,
                                    split_columns=selected_columns,
                                    period_drive_id=PERIOD_DRIVE_ID,
                                    pml_folder_id=PML_DRIVE_ID,
                                    log_pml_drive_id=log_pml_drive_id,
                                    year=int(year),
                                    month=int(month),
                                    biz_type=selected_rows.iloc[0]["Biz Type"],
                                    base_info=base_info,
                                    columns_template=columns_template_claim,
                                    progress_bar=progress_bar,
                                    status_text=status_text
                                )

                            # 🔥 UPDATE STATUS
                            update_pml_status_to_splitted(
                                service=sheets_service,
                                spreadsheet_id=log_pml_drive_id,
                                pml_id=selected_pml_id
                            )

                            progress_bar.progress(1.0)
                            status_text.text("✅ Selesai!")

                            st.success("✅ Split selesai & status diupdate!")

                            for r in results:
                                st.write(f"📄 {r['pml_id']} → {r['rows']} rows ({r['split_value']})")

                        finally:
                            release_drive_lock(service, PERIOD_DRIVE_ID)
                            st.session_state.is_processing_split = False

            else:
                st.write("Silakan pilih baris terlebih dahulu.")

        else:
            st.info("Tidak ada data dengan status 'POSTED'.")

    if action_type == "Delete Voucher":
        prod_year = st.selectbox(
            "Tahun Produksi",
            [year],
            key="prod_year",
            disabled=True
        )

        prod_month = st.selectbox(
            "Bulan Produksi",
            [month],
            key="prod_month",
            disabled=True
        )

    # elif action_type == 'Cancel Voucher':

    #     year = st.session_state["log_period"]["year"]
    #     month = st.session_state["log_period"]["month"]

    #     years = list(range(2026, datetime.now().year + 1))
    #     months = list(range(1, 13))

    #     prod_year = st.selectbox(
    #         "Tahun Produksi",
    #         years,
    #         key="prod_year",
    #         index=years.index(year)
    #     )

    #     # kalau pilih tahun berjalan → exclude bulan berjalan
    #     if prod_year == year:
    #         allowed_months = [m for m in months if m < month]
    #     else:
    #         allowed_months = months

    #     prod_month = st.selectbox(
    #         "Bulan Produksi",
    #         allowed_months,
    #         key="prod_month"
    #     )

    # prod_folders = get_period_drive_folders(
    #     year=prod_year,
    #     month=prod_month,
    #     root_folder_id=ROOT_DRIVE_FOLDER_ID
    # )

    # PROD_PERIOD_ID = prod_folders["period_id"]

    # log_drive_id = find_drive_file(
    #     service=service,
    #     filename=get_log_filename(int(prod_year), int(prod_month)),
    #     parent_id=PROD_PERIOD_ID,
    #     mime_type="application/vnd.google-apps.spreadsheet"
    # )

    # if not log_drive_id:
    #     st.info("Log belum tersedia")
    #     st.stop()

    # prod_log_df = load_log_from_gsheet(
    #     service=service,
    #     spreadsheet_id=log_drive_id
    # )

    # if "STATUS" not in prod_log_df.columns:
    #     st.error("Tidak ada voucher")
    #     st.stop()

    # # Filter hanya yang POSTED
    # posted_df = prod_log_df[prod_log_df["STATUS"] == "POSTED"]

    # if posted_df.empty:
    #     st.info("Tidak ada voucher POSTED")
    #     st.stop()

    # # =========================
    # # 1️⃣ PILIH CEDING
    # # =========================

    # ceding_list = sorted(posted_df["Account With"].dropna().unique())

    # selected_ceding = st.selectbox(
    #     "Pilih Ceding",
    #     ceding_list,
    #     key="update_ceding"
    # )

    # # =========================
    # # 2️⃣ FILTER BERDASARKAN CEDING
    # # =========================

    # ceding_df = posted_df[
    #     posted_df["Account With"] == selected_ceding
    # ]

    # # =========================
    # # 3️⃣ PILIH VOUCHER - PRODUCT
    # # =========================

    # voucher_options = [
    #     f"{row['Voucher No']} - {row['Product']}"
    #     for _, row in ceding_df.iterrows()
    # ]

    # selected_voucher_display = st.selectbox(
    #     "Pilih Voucher",
    #     voucher_options,
    #     key="update_voucher"
    # )

    # # Ambil voucher no asli
    # selected_voucher = selected_voucher_display.split(" - ")[0]

    # # =========================
    # # PIC
    # # =========================

    # pic = st.selectbox(
    #     "PIC",
    #     ["Ardelia", "Buya", "Khansa", "Prabu"],
    #     key="update_pic"
    # )

    # cancel_reason = st.text_area("Reason (WAJIB)")


    # # ==============================
    # # PROSES
    # # ==============================

    # button_label = None

    # if action_type == "Delete Voucher":
    #     button_label = "❌ Delete Voucher"
    # elif action_type == "Cancel Voucher":
    #     button_label = "🔁 Cancel Voucher"

    # if button_label is None:
    #     st.error(f"Action type tidak valid: {action_type}")
    #     st.stop()

    # if st.button(button_label, key="process_update"):

    #     if not cancel_reason.strip():
    #         st.error("Reason wajib diisi")
    #         st.stop()

    #     with st.spinner("⏳ Update voucher, mohon tunggu..."):
            
    #         try:
    #             acquire_drive_lock(service, PROD_PERIOD_ID)
                

    #             original_row = prod_log_df[
    #                 prod_log_df["Voucher No"] == selected_voucher
    #             ].iloc[0]

    #             # ==============================
    #             # DELETE VOUCHER
    #             # ==============================

    #             if action_type == "Delete Voucher":
                    
    #                 # Delete log record 
    #                 prod_log_df = prod_log_df[
    #                     prod_log_df["Voucher No"] != selected_voucher
    #                 ]

    #                 log_drive_id = find_drive_file(
    #                     service=service,
    #                     filename=get_log_filename(int(prod_year), int(prod_month)),
    #                     parent_id=PROD_PERIOD_ID,
    #                     mime_type="application/vnd.google-apps.spreadsheet"
    #                 )

    #                 update_gsheet(
    #                     service=service,
    #                     spreadsheet_id=log_drive_id,
    #                     df=prod_log_df
    #                 )

    #                 # Delete voucher file
    #                 ceding_folder_name = normalize_folder_name(original_row["Account With"])

    #                 ceding_drive = get_or_create_ceding_folders(
    #                     service=service,
    #                     period_folder_id=PROD_PERIOD_ID,
    #                     ceding_name=ceding_folder_name
    #                 )

    #                 CEDING_DRIVE_ID = ceding_drive["ceding_id"]

    #                 voucher_filename = f"{selected_voucher}.xlsx"

    #                 voucher_file_id = find_drive_file(
    #                     service=service,
    #                     filename=voucher_filename,
    #                     parent_id=CEDING_DRIVE_ID
    #                 )

    #                 files_in_folder = service.files().list(
    #                     q=f"'{CEDING_DRIVE_ID}' in parents and trashed=false",
    #                     fields="files(id,name,mimeType)",
    #                     supportsAllDrives=True,
    #                     includeItemsFromAllDrives=True,
    #                 ).execute()

    #                 if voucher_file_id:
    #                     service.files().delete(
    #                         fileId=voucher_file_id,
    #                         supportsAllDrives=True
    #                     ).execute()

    #                 st.success("Voucher & record berhasil dihapus")

    #             # ==============================
    #             # CANCEL LINTAS PERIODE
    #             # ==============================

    #             elif action_type == "Cancel Voucher":

    #                 # =============================
    #                 # 1️⃣ UPDATE LOG PERIODE LAMA
    #                 # =============================

    #                 mask = prod_log_df["Voucher No"].astype(str).str.strip() == str(selected_voucher).strip()

    #                 now_wib = datetime.now(ZoneInfo("Asia/Jakarta")).strftime("%Y-%m-%d %H:%M:%S")

    #                 prod_log_df.loc[mask, ["STATUS", "CANCELED AT", "CANCELED BY"]] = [
    #                     "CANCELED",
    #                     now_wib,
    #                     str(pic)
    #                 ]

    #                 log_drive_id = find_drive_file(
    #                     service=service,
    #                     filename=get_log_filename(int(prod_year), int(prod_month)),
    #                     parent_id=PROD_PERIOD_ID,
    #                     mime_type="application/vnd.google-apps.spreadsheet"
    #                 )

    #                 try:
    #                     update_gsheet(
    #                         service=service,
    #                         spreadsheet_id=log_drive_id,
    #                         df=prod_log_df
    #                     )
    #                 except Exception as e:
    #                     st.error(e)

    #                 # =============================
    #                 # 2️⃣ LOAD LOG BULAN SEKARANG
    #                 # =============================

    #                 now_year = st.session_state["log_period"]["year"]
    #                 now_month = st.session_state["log_period"]["month"]

    #                 now_folders = get_period_drive_folders(
    #                     year=now_year,
    #                     month=now_month,
    #                     root_folder_id=ROOT_DRIVE_FOLDER_ID
    #                 )

    #                 NOW_PERIOD_ID = now_folders["period_id"]

    #                 #acquire_drive_lock(service, NOW_PERIOD_ID)

    #                 current_log_drive_id = find_drive_file(
    #                     service=service,
    #                     filename=get_log_filename(int(now_year), int(now_month)),
    #                     parent_id=NOW_PERIOD_ID,
    #                     mime_type="application/vnd.google-apps.spreadsheet"
    #                 )

    #                 if not current_log_drive_id:
    #                     current_log_drive_id = create_log_gsheet(
    #                         service=service,
    #                         parent_id=NOW_PERIOD_ID,
    #                         filename=get_log_filename(int(now_year), int(now_month)),
    #                         columns=list(prod_log_df.columns)
    #                     )
                    

    #                 current_log_df = load_log_from_gsheet(
    #                     service=service,
    #                     spreadsheet_id=current_log_drive_id,
    #                 )

    #                 # =============================
    #                 # 3️⃣ GENERATE NOMOR BARU
    #                 # =============================
    #                 row = prod_log_df.loc[mask]

    #                 cancel_voucher, cancel_seq = generate_vin_from_drive_log(
    #                     log_df=current_log_df,
    #                     year=int(now_year),
    #                     month=int(now_month),
    #                     biz_type = row["Biz Type"].iloc[0]
    #                 )
                    
    #                 cancel_row = create_cancel_row(
    #                     original_row=original_row,
    #                     new_voucher=cancel_voucher,
    #                     seq_no=cancel_seq,
    #                     year = int(now_year),
    #                     month = int(now_month),
    #                     user=pic,
    #                     reason=cancel_reason
    #                 )

    #                 current_log_df = pd.concat(
    #                     [current_log_df, pd.DataFrame([cancel_row])],
    #                     ignore_index=True
    #                 )

    #                 update_gsheet(
    #                     service=service,
    #                     spreadsheet_id=current_log_drive_id,
    #                     df=current_log_df
    #                 )

    #                 st.success("Log berhasil diupdate")

    #                 # =============================
    #                 # 4️⃣ BUAT FILE REVERSAL
    #                 # =============================

    #                 # cari folder ceding bulan sekarang
    #                 ceding_folder_name = normalize_folder_name(original_row["Account With"])

    #                 ceding_drive = get_or_create_ceding_folders(
    #                     service=service,
    #                     period_folder_id=NOW_PERIOD_ID,
    #                     ceding_name=ceding_folder_name
    #                 )

    #                 CEDING_DRIVE_ID = ceding_drive["ceding_id"]

    #                 old_ceding_folder_name = normalize_folder_name(
    #                     original_row["Account With"]
    #                 )

    #                 old_ceding_drive = get_or_create_ceding_folders(
    #                     service=service,
    #                     period_folder_id=PROD_PERIOD_ID,   # ⬅️ periode produksi lama
    #                     ceding_name=old_ceding_folder_name
    #                 )

    #                 OLD_CEDING_DRIVE_ID = old_ceding_drive["ceding_id"]

    #                 # load file lama
    #                 original_file_df = load_voucher_excel_from_drive(
    #                     service=service,
    #                     voucher_no=selected_voucher,
    #                     ceding_folder_id=OLD_CEDING_DRIVE_ID
    #                 )

    #                 reversal_df = create_negative_excel(original_file_df, row["Voucher No"].iloc[0], cancel_voucher)

    #                 file_bytes = dataframe_to_excel_bytes(reversal_df)

    #                 upload_excel_bytes(
    #                     service=service,
    #                     file_bytes=file_bytes,
    #                     filename=f"{cancel_voucher}.xlsx",
    #                     parent_id=CEDING_DRIVE_ID
    #                 )

    #                 st.success(f"✅ Reversal dibuat: {cancel_voucher}")

    #         except RuntimeError:
    #             st.error("⛔ Log sedang digunakan user lain") 
            
    #         finally: 
    #             if PROD_PERIOD_ID:
    #                 release_drive_lock(service, PROD_PERIOD_ID)
    #             if NOW_PERIOD_ID:
    #                 release_drive_lock(service, NOW_PERIOD_ID)
    #             st.rerun()

