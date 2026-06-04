import streamlit as st
import pandas as pd
import os
import time
import st_aggrid


from datetime import datetime
from validator import validate_voucher, validate_calculate
from vin_generator import generate_vin, create_cancel_row, get_log_path, generate_vin_from_drive, generate_vou_from_drive, generate_vin_from_drive_log, create_negative_excel, dataframe_to_excel_bytes, upload_excel_bytes, get_log_filename, get_log_pml_filename, get_log_filename_outward, generate_vou_from_drive, generate_pml_from_drive, generate_pml_outward_from_drive, split_upload_with_log, split_upload_with_log_outward, get_last_seq_no, generate_pml_id
from drive_utils import upload_or_update_drive_file, get_period_drive_folders, get_or_create_folder, get_or_create_ceding_folders, get_drive_service, find_drive_file, acquire_drive_lock, release_drive_lock, upload_dataframe_to_drive, load_log_from_drive, upload_log_dataframe, load_voucher_excel_from_drive, calculate_due_date, get_exchange_rate, load_log_from_gsheet, update_gsheet, append_gsheet, create_log_gsheet, get_or_create_outward_folders, upload_dataframe_to_drive_outward, init_sheets_service, download_file_from_drive, download_file_csv_from_drive, update_pml_status_to_splitted, update_pml_status_to_calculated, create_review_spreadsheet, get_pml_metadata
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

tab_upload, tab_split, tab_calc = st.tabs([
    "📤 Upload File",
    "🧩 Split File",
    "🧮 Calculate",
    #"🔄 Update Voucher",
    # "📥 Create Voucher",
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
            index=0 #INWARD
            #disabled=True
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

    if reins_type == "INWARD":

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
            "Smoker",
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
            "MedicalCategory",
            "Product",
            "Coverage Code",
            "ClassOfBusiness",
            "PayPeriodType",
            "KindOfBusiness",
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

        LOG_COLUMNS = [
            "Seq No", "Department", "Biz Type", "Voucher No",
            "Account With", "Cedant Company", "PIC",
            "Product", "CBY", "CBM", "OBY", "OBM",
            "KOB", "COB", "MOP", "Curr",
            "Total Contribution", "Commission", "Overriding",
            "Total Commission", "Gross Premium Income",
            "Tabarru", "Ujrah", "Claim", "Balance", "Check Balance",
            "Rate Exchange",
            "Kontribusi (IDR)", "Commission (IDR)", "Overiding (IDR)",
            "Total Commission (IDR)", "Gross Premium Income (IDR)",
            "Tabarru (IDR)", "Ujrah (IDR)", "Claim (IDR)",
            "Balance (IDR)", "Check Balance (IDR)",
            "REMARKS", "PML ID", "STATUS", "CREATED AT", "CREATED BY",
            "Due Date", "Subject Email", "Email Date",
            "CANCELED AT", "CANCELED BY", "CANCEL OF VOUCHER", "CANCEL REASON"
        ]


        if uploaded_file:
            # ==========================
            # READ FILE
            # ==========================
            df = pd.read_excel(uploaded_file)
            original_columns = df.columns.tolist()
            df.columns = df.columns.str.strip().str.lower()

            for col in ["certificate no", "main pol no", "pol holder no"]:
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
                        "sum insured", "sum at risk", "reins sum insured", "ced retention", "reins sum at risk",
                        "reins premium", "reins em premium", "reins er premium", "reins oth. premium", "reins total premium",
                        "reins comm", "reins em comm", "reins er comm", "reins oth. comm", "reins profit share", "reins overriding", "reins broker fee", 
                        "reins total comm", "reins tabarru", "reins ujrah", "reins nett premium",
                        "sum insured idr", "sum reinsured idr", "amount of claim idr", "reins claim idr", "marein share idr"
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
                            "SCOR RE LABUAN BRANCH",
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
                            "SCOR RE LABUAN BRANCH",
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
                    st.error("Subject Email dan Remarks wajib diisi")
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
                            if biz_type == "Kontribusi":
                                df["trans category"] = "PREMIUM"

                            elif biz_type == "Refund":
                                df["trans category"] = "TERMINATE"
                                df["policy category"] = "T"

                            log_pml = {
                                "Seq No": seq_no,
                                "Department":department,
                                "Biz Type": biz_type,
                                "PML ID": pml_id,
                                "Account With": account_with,
                                "Cedant Company": cedant_company,
                                "PIC": pic,
                                "Product": df["references no"][0],
                                "CBY": df["cby"][0],
                                "CBM": df["cbm"][0],
                                "Curr":curr,
                                "Total Contribution": df["reins total premium"].sum(),
                                "Commission": df["reins comm"].sum(),
                                "Overriding": df["reins overriding"].sum() if "reins overriding" in df.columns else 0,
                                "Total Commission": (df["reins comm"].sum()) + (df["reins overriding"].sum() if "reins overriding" in df.columns else 0),
                                "Gross Premium Income": df["reins total premium"].sum() - ((df["reins total comm"].sum()) + (df["reins overriding"].sum() if "reins overriding" in df.columns else 0)),
                                "Tabarru": df["reins tabarru"].sum(),
                                "Ujrah": df["reins ujrah"].sum(),
                                "Claim": 0,
                                "Balance": df["reins total premium"].sum() - df["reins comm"].sum() - (df["reins overriding"].sum() if "reins overriding" in df.columns else 0) - (df["claim"].sum() if "claim" in df.columns else 0),
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
                                "Product": df["references no"][0],
                                "CBY": df["cedbookyear"][0],
                                "CBM": df["cedbookmonth"][0],
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

    elif reins_type == "OUTWARD":

        columns_template_outward = [
            "Out PL Detail ID",
            "Retro Type",
            "Acc With Name",
            "Policy Category",
            "KOB Code",
            "Ref Offer Insured Risk ID",
            "Main Pol No",
            "Main Policy",
            "Pol Holder No",
            "Policy Holder",
            "Certificate No",
            "Insured Full Name",
            "Birth Date",
            "Gender",
            "Issue Date",
            "Age At",
            "Term Year",
            "Term Month",
            "Expired Date",
            "RI Period From",
            "RI Period Until",
            "Smoker",
            "Medical",
            "Ced Product Code",
            "Ced Coverage Code",
            "Ced Risk Code",
            "Life Risk Detail",
            "PA Class Category",
            "Ccy Code",
            "Sum Insured",
            "Sum At Risk",
            "Reins Sum Insured",
            "Reins Sum At Risk",
            "Marein Sum Insured",
            "Marein Sum At Risk",
            "Own Retention",
            "Excess OR",
            "Check_1",
            "Retro Sum Insured",
            "Retro Sum At Risk",
            "Premium Ccy",
            "Exchange Rate",
            "Out Tty Rate",
            "Inw Tty Rate",
            "EM Rate",
            "ER Rate",
            "Retro Premium",
            "Retro EM Premium",
            "Retro ER Premium",
            "Retro Oth Premium",
            "Retro Total Premium",
            "Retro Comm",
            "Retro EM Comm",
            "Retro ER Comm",
            "Retro Oth Comm",
            "Retro Profit Share",
            "Retro Total Comm",
            "Retro Tabarru",
            "Retro Ujrah",
            "Check_2",
            "Retro Overriding",
            "Retro Sliding Scale",
            "Retro Inw Brokerage",
            "Reins Nett Premium",
            "Retro Nett Premium",
            "Check_3",
            "Is Accum Policy",
            "Is Calculated",
            "PL Detail ID",
            "Inw Vouc ID",
            "Out Vouc ID",
            "Inw Tty Product Code",
            "Inw Book Year",
            "Inw Book Month",
            "Ced Book Year",
            "Ced Book Month",
            "Inw Pay Period Type",
            "Out Pay Period Type",
            "COB",
            "Valuation Date",
            "Term Condition Remark",
            "App Date",
            "Input Date",
            "Input Username",
            "Modif Date",
            "Modif Username",
            "References No"
        ]

        columns_template_claim_outward = [
            "No",
            "Retro Type",
            "Cedant Name",
            "DLA Out Voucher ID",
            "Main Pol No",
            "Main Policy",
            "Pol Holder No",
            "Policy Holder",
            "Certificate No",
            "Insured Name",
            "Birth Date",
            "Age",
            "Gender",
            "Ced Product Code",
            "Ced Coverage Code",
            "Ced Risk Code",
            "COB Detail",
            "Issue Date",
            "Term Year",
            "Term Month",
            "KOB Code",
            "Smoker",
            "Medical",
            "Claim Date",
            "Cause Of Claim",
            "Inw Book Year",
            "Inw Book Month",
            "Ced Book Year",
            "Ced Book Month",
            "Method of Payment",
            "Curr",
            "Reins Claim",
            "Your Share",
            "Reinsurer Name",
            "Voucher ID",
            "Out Voucher ID",
            "Voucher Desc"
        ]

        LOG_COLUMNS_OUTWARD = [
            "Seq No",
            "Department",
            "Biz Type",
            "Retro Type",
            "Inward VIN Ref",
            "Voucher No",
            "Account With",
            "Cedant Company",
            "PIC",
            "Product",
            "CBY",
            "CBM",
            "OBY",
            "OBM",
            "KOB",
            "COB",
            "MOP",
            "Curr",
            "Total Contribution",
            "Commission",
            "Overiding",
            "Total Commission",
            "Gross Premium Income",
            "Tabarru",
            "Ujrah",
            "Ujrah Spc",
            "Claim",
            "Balance",
            "Check Balance",
            "Rate Exchange",
            "Kontribusi (IDR)",
            "Commission (IDR)",
            "Overiding (IDR)",
            "Total Commission (IDR)",
            "Gross Premium Income (IDR)",
            "Tabarru (IDR)",
            "Ujrah (IDR)",
            "Ujrah Spc (IDR)",
            "Claim (IDR)",
            "Balance (IDR)",
            "Check Balance (IDR)",
            "REMARKS",
            "PML ID",
            "STATUS",
            "CREATED AT",
            "CREATED BY",
            "Due Date",
            "CANCELED AT",
            "CANCELED BY",
            "CANCEL OF VOUCHER",
            "CANCEL REASON"
        ]


        if uploaded_file:
            # ==========================
            # READ FILE
            # ==========================
            df = pd.read_excel(uploaded_file)
            original_columns = df.columns.tolist()
            df.columns = df.columns.str.strip().str.lower()

            for col in ["certificate no", "main pol no", "pol holder no"]:
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
                        "marein sum insured", "marein sum at risk", "own retention", "excess or",
                        "retro sum insured", "retro sum at risk",
                        "retro premium", "retro em premium", "retro er premium", "retro oth premium", "retro total premium",
                        "retro comm", "retro em comm", "retro er comm", "retro oth comm", "retro profit share", "retro total comm", 
                        "retro tabarru", "retro ujrah", "retro overriding", "retro sliding scale", "retro inw brokerage", "reins nett premium", "retro nett premium",
                        "reins claim", "your share"
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
                            "REASURANSI INTERNATIONAL INDONESIA SYARIAH",
                            "REASURANSI NASIONAL INDONESIA SYARIAH",
                            "HANNOVER RETAKAFUL",
                            "MAREIN SYARIAH",
                            "MUNICH RE RETAKAFUL",
                            "SCOR RE LABUAN BRANCH",
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
                            "SCOR RE LABUAN BRANCH",
                            "SWISS RE INTL. SE, SINGAPORE (SYARIAH)"
                        ]
                    )

                    curr = st.selectbox(
                        "Currency",
                        ["IDR", "USD"]
                    )

                with col2:
                    
                    subject_email = st.text_area("Subject Email", disabled=True)

                    email_date = st.date_input("Email Date",value=date.today(), disabled=True)

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
                        "Retro Share"
                    ],
                    "Nilai": [
                        df["reins claim"].sum(),
                        df["your share"].sum()
                        ]
                })
                

            #st.write(df.columns.tolist())


            st.dataframe(
                summary_df.style.format({"Nilai": "{:,.2f}"}),
                use_container_width=True
            )


            # ==========================
            # POST VOUCHER (LOCKED)
            # ==========================
            if st.button("💾 Simpan File"):
                start_time = time.time()

                if biz_type == "Claim":
                    df["cedant name"] = cedant_company
                    df["reinsurer name"] = account_with


                if not remarks.strip():
                    st.error("Remarks wajib diisi")
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
                            folder_name="Folder PML (Outward)",
                            parent_id=PERIOD_DRIVE_ID
                        )

                        PML_DRIVE_ID = pml_drive

                        pml_id, seq_no, file_id = generate_pml_outward_from_drive(
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
                            filename=f"{get_log_pml_filename(int(year), int(month))} (Outward)",
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
                                "Product": df["references no"][0],
                                "CBY": df["ced book year"][0],
                                "CBM": df["ced book month"][0],
                                "Curr":curr,
                                "Total Contribution": df["retro total premium"].sum(),
                                "Commission": df["retro total comm"].sum(),
                                "Overriding": df["retro overriding"].sum() if "retro overriding" in df.columns else 0,
                                "Total Commission": (df["retro total comm"].sum()) + (df["retro overriding"].sum() if "retro overriding" in df.columns else 0),
                                "Gross Premium Income": df["retro total premium"].sum() - ((df["retro total comm"].sum()) + (df["retro overriding"].sum() if "retro overriding" in df.columns else 0)),
                                "Tabarru": df["retro tabarru"].sum(),
                                "Ujrah": df["retro ujrah"].sum(),
                                "Claim": 0,
                                "Balance": df["retro total premium"].sum() - df["retro total comm"].sum() - (df["retro overriding"].sum() if "retro overriding" in df.columns else 0) - (df["claim"].sum() if "claim" in df.columns else 0),
                                "REMARKS": remarks,
                                "STATUS": "POSTED",
                                #"ENTRY_TYPE": entry_type,
                                "CREATED AT": now_wib_naive(),
                                "CREATED BY": pic,
                                "Subject Email": "-",
                                "Email Date": "-",
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
                                "Product": df["voucher desc"][0],
                                "CBY": df["ced book year"][0],
                                "CBM": df["ced book month"][0],
                                "Curr":curr,
                                "Total Contribution": 0,
                                "Commission": 0,
                                "Overriding": 0,
                                "Total Commission": 0,
                                "Gross Premium Income": 0,
                                "Tabarru": 0,
                                "Ujrah": 0,
                                "Claim": df["your share"].sum(),
                                "Balance": 0 - (df["your share"].sum() if "your share" in df.columns else 0),
                                "REMARKS": remarks,
                                "STATUS": "POSTED",
                                #"ENTRY_TYPE": entry_type,
                                "CREATED AT": now_wib_naive(),
                                "CREATED BY": pic,
                                "Subject Email": "-",
                                "Email Date": "-",
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
                                filename=f"{get_log_pml_filename(int(year), int(month))} (Outward)",
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
                                template_columns=columns_template_outward,
                                voucher_id=pml_id,
                                filename=f"{pml_id}.xlsx",
                                folder_id=PML_DRIVE_ID,
                                file_type = "PML"
                            )

                        elif biz_type == "Claim" :
                            df["dla out voucher id"] = pml_id
                            upload_dataframe_to_drive(
                                service=service,
                                df=df,
                                template_columns=columns_template_claim_outward,
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
# TAB SPLIT
# ==========================
with tab_split:
    st.subheader("🧩 Split File")

    reins_type = st.selectbox(
        "Reinsurance Type",
        ["INWARD", "OUTWARD"],
        key="reins_type_split",
        index=0) #INWARD
        #disabled=True

    if reins_type == "INWARD":

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
            "Smoker",
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
            "MedicalCategory",
            "Product",
            "Coverage Code",
            "ClassOfBusiness",
            "PayPeriodType",
            "KindOfBusiness",
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

        LOG_COLUMNS = [
            "Seq No", "Department", "Biz Type", "Voucher No",
            "Account With", "Cedant Company", "PIC",
            "Product", "CBY", "CBM", "OBY", "OBM",
            "KOB", "COB", "MOP", "Curr",
            "Total Contribution", "Commission", "Overriding",
            "Total Commission", "Gross Premium Income",
            "Tabarru", "Ujrah", "Claim", "Balance", "Check Balance",
            "Rate Exchange",
            "Kontribusi (IDR)", "Commission (IDR)", "Overiding (IDR)",
            "Total Commission (IDR)", "Gross Premium Income (IDR)",
            "Tabarru (IDR)", "Ujrah (IDR)", "Claim (IDR)",
            "Balance (IDR)", "Check Balance (IDR)",
            "REMARKS", "PML ID", "STATUS", "CREATED AT", "CREATED BY",
            "Due Date", "Subject Email", "Email Date",
            "CANCELED AT", "CANCELED BY", "CANCEL OF VOUCHER", "CANCEL REASON"
        ]

        drive_service = get_drive_service()
        sheets_service = init_sheets_service(creds)

        service = get_drive_service()

        year = st.session_state["log_period"]["year"]
        month = st.session_state["log_period"]["month"]

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
            # acquire_drive_lock(drive_service, PERIOD_DRIVE_ID)

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
        
        # finally:
        #     # 4. RELEASE LOCK (Selalu dijalankan meskipun error di atas)
        #     if PERIOD_DRIVE_ID:
        #         release_drive_lock(drive_service, PERIOD_DRIVE_ID)

        if not df_posted.empty:

            # ==========================
            # SNAPSHOT LOG KE SESSION STATE
            # ==========================
            log_snapshot_key = f"log_snapshot_split_inward_{year}_{month}"

            if log_snapshot_key not in st.session_state:
                st.session_state[log_snapshot_key] = df_posted.copy()

            df_working = st.session_state[log_snapshot_key]

            # ==========================
            # REFRESH LOG PML
            # ==========================
            col_ref1, col_ref2 = st.columns([2, 5])

            with col_ref1:
                if st.button("🔄 Refresh Log PML", key="btn_refresh_pml_split_inward"):
                    if log_snapshot_key in st.session_state:
                        del st.session_state[log_snapshot_key]
                    st.cache_data.clear()
                    st.rerun()

            with col_ref2:
                current_posted_count  = len(df_posted)
                snapshot_posted_count = len(df_working)

                if current_posted_count != snapshot_posted_count:
                    st.warning(
                        f"⚠️ Log PML telah diperbarui oleh user lain "
                        f"({snapshot_posted_count} → {current_posted_count} baris POSTED). "
                        f"Klik Refresh jika ingin memuat data terbaru."
                    )

            # Gunakan snapshot
            df_posted = df_working.copy()

            # Terapkan filter
            df_filtered = df_posted.copy()

            # --- 5. RENDER UI DENGAN PEMILIHAN (CHECKBOX) ---
            st.markdown("### 📋 Pilih Data PML untuk Di-Split")
            st.info("Centang pada kolom **'Pilih'** untuk menentukan baris yang akan diproses.")

            # Tambahkan checkbox column
            df_to_edit = df_filtered.copy()
            df_to_edit.insert(0, "Pilih", False)

            # Data editor (CONSISTENT UI)
            cols_numeric = ["Total Contribution", "Commission", "Overriding", "Total Commission", "Gross Premium Income", "Tabarru", "Ujrah", "Claim", "Balance"]

            for col in cols_numeric:

                def clean_number(x):
                    x = str(x)

                    # kalau ada dua separator → anggap titik ribuan, koma desimal
                    if "." in x and "," in x:
                        x = x.replace(".", "").replace(",", ".")
                    else:
                        x = x.replace(",", "")

                    return pd.to_numeric(x, errors="coerce")

                df_to_edit[col] = df_to_edit[col].apply(clean_number)

            edited_df = st.data_editor(
                df_to_edit,
                column_config={
                    "Pilih": st.column_config.CheckboxColumn(
                        "Pilih",
                        help="Pilih baris ini untuk di-split",
                        default=False,
                    ),
                    "PML ID":  st.column_config.Column(disabled=True),
                    "STATUS":  st.column_config.Column(disabled=True),
                    "Product": st.column_config.Column("Product", disabled=True),
                    "CBY":     st.column_config.Column("CBY", disabled=True),
                    "CBM":     st.column_config.Column("CBM", disabled=True),
                    "Total Contribution":   st.column_config.NumberColumn("Total Contribution",   format="%,.0f"),
                    "Commission":           st.column_config.NumberColumn("Commission",           format="%,.0f"),
                    "Overriding":           st.column_config.NumberColumn("Overriding",           format="%,.0f"),
                    "Total Commission":     st.column_config.NumberColumn("Total Commission",     format="%,.0f"),
                    "Gross Premium Income": st.column_config.NumberColumn("Gross Premium Income", format="%,.0f"),
                    "Tabarru": st.column_config.NumberColumn("Tabarru", format="%,.0f"),
                    "Ujrah":   st.column_config.NumberColumn("Ujrah",   format="%,.0f"),
                    "Claim":   st.column_config.NumberColumn("Claim",   format="%,.0f"),
                    "Balance": st.column_config.NumberColumn("Balance", format="%,.0f"),
                },
                disabled=[
                    "No", "PML ID", "STATUS", "Product", "CBY", "CBM",
                    "Total Contribution", "Commission", "Overriding",
                    "Total Commission", "Gross Premium Income",
                    "Tabarru", "Ujrah", "Claim", "Balance"
                ],
                hide_index=True,
                use_container_width=True,
                key="data_editor_split_inward"
            )

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
                    filename=f"{selected_pml_id}.xlsx",
                    parent_id=PML_DRIVE_ID,
                    mime_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
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
                    filename=f"{selected_pml_id}.xlsx",
                    parent_id=PML_DRIVE_ID,
                    mime_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                if not pml_file_id:
                    st.error("File PML tidak ditemukan")
                    st.stop()


                # ==========================
                # LOAD FILE
                # ==========================
                file_stream = download_file_from_drive(service, pml_file_id)
                df = pd.read_excel(file_stream)

                ACCOUNTING_COLS = [
                    "Sum Insured", "Sum At Risk", "Reins Sum Insured", "Reins Sum At Risk", "Ced Retention",
                    "Reins Premium", "Reins EM Premium", "Reins ER Premium", "Reins Oth. premium", "Reins Total Premium",
                    "Reins Comm", "Reins EM Comm", "Reins ER Comm", "Reins Oth. Comm", "Reins Profit Share", "Reins Total Comm", 
                    "Reins Tabarru", "Reins Ujrah", "Reins Overriding", "Reins Sliding Scale", "Reins Inw Brokerage", "Reins Nett Premium",
                    "Sum Insured IDR", "Sum Reinsured IDR", "Amount of Claim IDR", "Reins Claim IDR", "Marein Share IDR"
                ]

                # Buat dictionary formatter untuk kolom yang ada saja
                format_dict = {}

                for col in ACCOUNTING_COLS:
                    if col in df.columns:
                        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                        format_dict[col] = "{:,.2f}"

                # =========================
                # SAFE DATAFRAME
                # =========================
                safe_df = df.copy()

                for col in safe_df.columns:

                    # Kolom accounting tetap numerik
                    if col in ACCOUNTING_COLS:
                        continue

                    # Selain accounting -> string
                    safe_df[col] = safe_df[col].astype(str).replace("nan", "")

                # =========================
                # RENDER
                # =========================
                try:
                    st.dataframe(
                        safe_df.style.format(format_dict),
                        use_container_width=True
                    )

                except Exception:
                    # fallback final
                    st.dataframe(
                        safe_df.astype(str),
                        use_container_width=True
                    )

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

                        try:
                            acquire_drive_lock(service, PERIOD_DRIVE_ID)
                            
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

                        except RuntimeError:
                            st.error("⛔ Log sedang digunakan user lain. Silakan coba lagi.")

                        finally:
                            release_drive_lock(service, PERIOD_DRIVE_ID)
                            st.session_state.is_processing_split = False

            else:
                st.write("Silakan pilih baris terlebih dahulu.")

        else:
            st.info("Tidak ada data dengan status 'POSTED'.")

    elif reins_type == "OUTWARD":

        columns_template_outward = [
            "Out PL Detail ID",
            "Retro Type",
            "Acc With Name",
            "Policy Category",
            "KOB Code",
            "Ref Offer Insured Risk ID",
            "Main Pol No",
            "Main Policy",
            "Pol Holder No",
            "Policy Holder",
            "Certificate No",
            "Insured Full Name",
            "Birth Date",
            "Gender",
            "Issue Date",
            "Age At",
            "Term Year",
            "Term Month",
            "Expired Date",
            "RI Period From",
            "RI Period Until",
            "Smoker",
            "Medical",
            "Ced Product Code",
            "Ced Coverage Code",
            "Ced Risk Code",
            "Life Risk Detail",
            "PA Class Category",
            "Ccy Code",
            "Sum Insured",
            "Sum At Risk",
            "Reins Sum Insured",
            "Reins Sum At Risk",
            "Marein Sum Insured",
            "Marein Sum At Risk",
            "Own Retention",
            "Excess OR",
            "Check_1",
            "Retro Sum Insured",
            "Retro Sum At Risk",
            "Premium Ccy",
            "Exchange Rate",
            "Out Tty Rate",
            "Inw Tty Rate",
            "EM Rate",
            "ER Rate",
            "Retro Premium",
            "Retro EM Premium",
            "Retro ER Premium",
            "Retro Oth Premium",
            "Retro Total Premium",
            "Retro Comm",
            "Retro EM Comm",
            "Retro ER Comm",
            "Retro Oth Comm",
            "Retro Profit Share",
            "Retro Total Comm",
            "Retro Tabarru",
            "Retro Ujrah",
            "Check_2",
            "Retro Overriding",
            "Retro Sliding Scale",
            "Retro Inw Brokerage",
            "Reins Nett Premium",
            "Retro Nett Premium",
            "Check_3",
            "Is Accum Policy",
            "Is Calculated",
            "PL Detail ID",
            "Inw Vouc ID",
            "Out Vouc ID",
            "Inw Tty Product Code",
            "Inw Book Year",
            "Inw Book Month",
            "Ced Book Year",
            "Ced Book Month",
            "Inw Pay Period Type",
            "Out Pay Period Type",
            "COB",
            "Valuation Date",
            "Term Condition Remark",
            "App Date",
            "Input Date",
            "Input Username",
            "Modif Date",
            "Modif Username",
            "References No"
        ]

        columns_template_claim_outward = [
            "No",
            "Retro Type",
            "Cedant Name",
            "DLA Out Voucher ID",
            "Main Pol No",
            "Main Policy",
            "Pol Holder No",
            "Policy Holder",
            "Certificate No",
            "Insured Name",
            "Birth Date",
            "Age",
            "Gender",
            "Ced Product Code",
            "Ced Coverage Code",
            "Ced Risk Code",
            "COB Detail",
            "Issue Date",
            "Term Year",
            "Term Month",
            "KOB Code",
            "Smoker",
            "Medical",
            "Claim Date",
            "Cause Of Claim",
            "Inw Book Year",
            "Inw Book Month",
            "Ced Book Year",
            "Ced Book Month",
            "Method of Payment",
            "Curr",
            "Reins Claim",
            "Your Share",
            "Reinsurer Name",
            "Voucher ID",
            "Out Voucher ID",
            "Voucher Desc"
        ]

        LOG_COLUMNS_OUTWARD = [
            "Seq No",
            "Department",
            "Biz Type",
            "Retro Type",
            "Inward VIN Ref",
            "Voucher No",
            "Account With",
            "Cedant Company",
            "PIC",
            "Product",
            "CBY",
            "CBM",
            "OBY",
            "OBM",
            "KOB",
            "COB",
            "MOP",
            "Curr",
            "Total Contribution",
            "Commission",
            "Overiding",
            "Total Commission",
            "Gross Premium Income",
            "Tabarru",
            "Ujrah",
            "Ujrah Spc",
            "Claim",
            "Balance",
            "Check Balance",
            "Rate Exchange",
            "Kontribusi (IDR)",
            "Commission (IDR)",
            "Overiding (IDR)",
            "Total Commission (IDR)",
            "Gross Premium Income (IDR)",
            "Tabarru (IDR)",
            "Ujrah (IDR)",
            "Ujrah Spc (IDR)",
            "Claim (IDR)",
            "Balance (IDR)",
            "Check Balance (IDR)",
            "REMARKS",
            "PML ID",
            "STATUS",
            "CREATED AT",
            "CREATED BY",
            "Due Date",
            "CANCELED AT",
            "CANCELED BY",
            "CANCEL OF VOUCHER",
            "CANCEL REASON"
        ]

        drive_service = get_drive_service()
        sheets_service = init_sheets_service(creds)

        service = get_drive_service()

        year = st.session_state["log_period"]["year"]
        month = st.session_state["log_period"]["month"]

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
                folder_name="Folder PML (Outward)",
                parent_id=PERIOD_DRIVE_ID
            )

            # Cari File Log Spreadsheet
            log_pml_drive_id = find_drive_file(
                service=drive_service,
                filename=f"{get_log_pml_filename(int(year), int(month))} (Outward)",
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

        if not df_posted.empty:
            # ==========================
            # SNAPSHOT LOG KE SESSION STATE
            # ==========================
            log_snapshot_key = f"log_snapshot_split_outward_{year}_{month}"

            if log_snapshot_key not in st.session_state:
                st.session_state[log_snapshot_key] = df_posted.copy()

            df_working = st.session_state[log_snapshot_key]

            # ==========================
            # REFRESH LOG PML
            # ==========================
            col_ref1, col_ref2 = st.columns([2, 5])

            with col_ref1:
                if st.button("🔄 Refresh Log PML", key="btn_refresh_pml_split_outward"):
                    if log_snapshot_key in st.session_state:
                        del st.session_state[log_snapshot_key]
                    st.cache_data.clear()
                    st.rerun()

            with col_ref2:
                current_posted_count  = len(df_posted)
                snapshot_posted_count = len(df_working)

                if current_posted_count != snapshot_posted_count:
                    st.warning(
                        f"⚠️ Log PML telah diperbarui oleh user lain "
                        f"({snapshot_posted_count} → {current_posted_count} baris POSTED). "
                        f"Klik Refresh jika ingin memuat data terbaru."
                    )

            # Gunakan snapshot
            df_posted = df_working.copy()

            # Terapkan filter
            df_filtered = df_posted.copy()

            # --- 5. RENDER UI DENGAN PEMILIHAN (CHECKBOX) ---
            st.markdown("### 📋 Pilih Data PML untuk Di-Split")
            st.info("Centang pada kolom **'Pilih'** untuk menentukan baris yang akan diproses.")

            # Tambahkan checkbox column
            df_to_edit = df_filtered.copy()
            df_to_edit.insert(0, "Pilih", False)

            # Data editor (CONSISTENT UI)
            cols_numeric = ["Total Contribution", "Commission", "Overriding", "Total Commission", "Gross Premium Income", "Tabarru", "Ujrah", "Claim", "Balance"]

            for col in cols_numeric:

                def clean_number(x):
                    x = str(x)

                    # kalau ada dua separator → anggap titik ribuan, koma desimal
                    if "." in x and "," in x:
                        x = x.replace(".", "").replace(",", ".")
                    else:
                        x = x.replace(",", "")

                    return pd.to_numeric(x, errors="coerce")

                df_to_edit[col] = df_to_edit[col].apply(clean_number)

            edited_df = st.data_editor(
                df_to_edit,
                column_config={
                    "Pilih": st.column_config.CheckboxColumn(
                        "Pilih",
                        help="Pilih baris ini untuk di-split",
                        default=False,
                    ),
                    "PML ID":  st.column_config.Column(disabled=True),
                    "STATUS":  st.column_config.Column(disabled=True),
                    "Product": st.column_config.Column("Product", disabled=True),
                    "CBY":     st.column_config.Column("CBY", disabled=True),
                    "CBM":     st.column_config.Column("CBM", disabled=True),
                    "Total Contribution":   st.column_config.NumberColumn("Total Contribution",   format="%,.0f"),
                    "Commission":           st.column_config.NumberColumn("Commission",           format="%,.0f"),
                    "Overriding":           st.column_config.NumberColumn("Overriding",           format="%,.0f"),
                    "Total Commission":     st.column_config.NumberColumn("Total Commission",     format="%,.0f"),
                    "Gross Premium Income": st.column_config.NumberColumn("Gross Premium Income", format="%,.0f"),
                    "Tabarru": st.column_config.NumberColumn("Tabarru", format="%,.0f"),
                    "Ujrah":   st.column_config.NumberColumn("Ujrah",   format="%,.0f"),
                    "Claim":   st.column_config.NumberColumn("Claim",   format="%,.0f"),
                    "Balance": st.column_config.NumberColumn("Balance", format="%,.0f"),
                },
                disabled=[
                    "No", "PML ID", "STATUS", "Product", "CBY", "CBM",
                    "Total Contribution", "Commission", "Overriding",
                    "Total Commission", "Gross Premium Income",
                    "Tabarru", "Ujrah", "Claim", "Balance"
                ],
                hide_index=True,
                use_container_width=True,
                key="data_editor_split_outward"
            )

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
                    folder_name="Folder PML (Outward)",
                    parent_id=PERIOD_DRIVE_ID
                )

                PML_DRIVE_ID = pml_drive


                pml_file_id = find_drive_file(
                    service=service,
                    filename=f"{selected_pml_id}.xlsx",
                    parent_id=PML_DRIVE_ID,
                    mime_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                service = get_drive_service()

                # Folder
                pml_drive = get_or_create_folder(
                    service=service,
                    folder_name="Folder PML (Outward)",
                    parent_id=PERIOD_DRIVE_ID
                )

                PML_DRIVE_ID = pml_drive

                # Cari file
                pml_file_id = find_drive_file(
                    service=service,
                    filename=f"{selected_pml_id}.xlsx",
                    parent_id=PML_DRIVE_ID,
                    mime_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

                if not pml_file_id:
                    st.error("File PML tidak ditemukan")
                    st.stop()


                # ==========================
                # LOAD FILE
                # ==========================
                file_stream = download_file_from_drive(service, pml_file_id)
                df = pd.read_excel(file_stream)

                ACCOUNTING_COLS = [
                    "Sum Insured", "Sum At Risk", "Reins Sum Insured", "Reins Sum At Risk",
                    "Marein Sum Insured", "Marein Sum At Risk", "Own Retention", "Excess OR",
                    "Retro Sum Insured", "Retro Sum At Risk",
                    "Retro Premium", "Retro EM Premium", "Retro ER Premium", "Retro Oth premium", "Retro Total Premium",
                    "Retro Comm", "Retro EM Comm", "Retro ER Comm", "Retro Oth Comm", "Retro Profit Share", "Retro Total Comm", 
                    "Retro Tabarru", "Retro Ujrah", "Retro Overriding", "Retro Sliding Scale", "Retro Inw Brokerage", "Reins Nett Premium", "Retro Nett Premium",
                    "Reins Claim", "Your Share"
                ]

                # Buat dictionary formatter untuk kolom yang ada saja
                format_dict = {}

                for col in ACCOUNTING_COLS:
                    if col in df.columns:
                        df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
                        format_dict[col] = "{:,.2f}"

                # =========================
                # SAFE DATAFRAME
                # =========================
                safe_df = df.copy()

                for col in safe_df.columns:

                    # Kolom accounting tetap numerik
                    if col in ACCOUNTING_COLS:
                        continue

                    # Selain accounting -> string
                    safe_df[col] = safe_df[col].astype(str).replace("nan", "")

                # =========================
                # RENDER
                # =========================
                try:
                    st.dataframe(
                        safe_df.style.format(format_dict),
                        use_container_width=True
                    )

                except Exception:
                    # fallback final
                    st.dataframe(
                        safe_df.astype(str),
                        use_container_width=True
                    )


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
                        # st.stop()

                    st.session_state.is_processing_split = True

                    with st.spinner(f"⏳ Split {selected_pml_id} sedang diproses..."):

                        progress_bar = st.progress(0)
                        status_text = st.empty()

                        acquire_drive_lock(service, PERIOD_DRIVE_ID)

                        try:
                            sheets_service = init_sheets_service(creds)

                            log_pml_drive_id = find_drive_file(
                                service=service,
                                filename=f"{get_log_pml_filename(int(year), int(month))} (Outward)",
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
                                results = split_upload_with_log_outward(
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
                                    columns_template=columns_template_outward,
                                    progress_bar=progress_bar,
                                    status_text=status_text
                                )

                            elif selected_rows.iloc[0]["Biz Type"] == "Claim":
                                    results = split_upload_with_log_outward(
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
                                    columns_template=columns_template_claim_outward,
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



# ==========================
# TAB CALCULATE
# ==========================
with tab_calc:

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
        "Smoker",
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
        "MedicalCategory",
        "Product",
        "Coverage Code",
        "ClassOfBusiness",
        "PayPeriodType",
        "KindOfBusiness",
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

    LOG_COLUMNS = [
        "Seq No", "Department", "Biz Type", "Voucher No",
        "Account With", "Cedant Company", "PIC",
        "Product", "CBY", "CBM", "OBY", "OBM",
        "KOB", "COB", "MOP", "Curr",
        "Total Contribution", "Commission", "Overriding",
        "Total Commission", "Gross Premium Income",
        "Tabarru", "Ujrah", "Claim", "Balance", "Check Balance",
        "Rate Exchange",
        "Kontribusi (IDR)", "Commission (IDR)", "Overiding (IDR)",
        "Total Commission (IDR)", "Gross Premium Income (IDR)",
        "Tabarru (IDR)", "Ujrah (IDR)", "Claim (IDR)",
        "Balance (IDR)", "Check Balance (IDR)",
        "REMARKS", "PML ID", "STATUS", "CREATED AT", "CREATED BY",
        "Due Date", "Subject Email", "Email Date",
        "CANCELED AT", "CANCELED BY", "CANCEL OF VOUCHER", "CANCEL REASON"
    ]


    columns_template_outward = [
        "Out PL Detail ID",
        "Retro Type",
        "Acc With Name",
        "Policy Category",
        "KOB Code",
        "Ref Offer Insured Risk ID",
        "Main Pol No",
        "Main Policy",
        "Pol Holder No",
        "Policy Holder",
        "Certificate No",
        "Insured Full Name",
        "Birth Date",
        "Gender",
        "Issue Date",
        "Age At",
        "Term Year",
        "Term Month",
        "Expired Date",
        "RI Period From",
        "RI Period Until",
        "Smoker",
        "Medical",
        "Ced Product Code",
        "Ced Coverage Code",
        "Ced Risk Code",
        "Life Risk Detail",
        "PA Class Category",
        "Ccy Code",
        "Sum Insured",
        "Sum At Risk",
        "Reins Sum Insured",
        "Reins Sum At Risk",
        "Marein Sum Insured",
        "Marein Sum At Risk",
        "Own Retention",
        "Excess OR",
        "Check_1",
        "Retro Sum Insured",
        "Retro Sum At Risk",
        "Premium Ccy",
        "Exchange Rate",
        "Out Tty Rate",
        "Inw Tty Rate",
        "EM Rate",
        "ER Rate",
        "Retro Premium",
        "Retro EM Premium",
        "Retro ER Premium",
        "Retro Oth Premium",
        "Retro Total Premium",
        "Retro Comm",
        "Retro EM Comm",
        "Retro ER Comm",
        "Retro Oth Comm",
        "Retro Profit Share",
        "Retro Total Comm",
        "Retro Tabarru",
        "Retro Ujrah",
        "Check_2",
        "Retro Overriding",
        "Retro Sliding Scale",
        "Retro Inw Brokerage",
        "Reins Nett Premium",
        "Retro Nett Premium",
        "Check_3",
        "Is Accum Policy",
        "Is Calculated",
        "PL Detail ID",
        "Inw Vouc ID",
        "Out Vouc ID",
        "Inw Tty Product Code",
        "Inw Book Year",
        "Inw Book Month",
        "Ced Book Year",
        "Ced Book Month",
        "Inw Pay Period Type",
        "Out Pay Period Type",
        "COB",
        "Valuation Date",
        "Term Condition Remark",
        "App Date",
        "Input Date",
        "Input Username",
        "Modif Date",
        "Modif Username",
        "References No"
    ]

    columns_template_claim_outward = [
        "No",
        "Retro Type",
        "Cedant Name",
        "DLA Out Voucher ID",
        "Main Pol No",
        "Main Policy",
        "Pol Holder No",
        "Policy Holder",
        "Certificate No",
        "Insured Name",
        "Birth Date",
        "Age",
        "Gender",
        "Ced Product Code",
        "Ced Coverage Code",
        "Ced Risk Code",
        "COB Detail",
        "Issue Date",
        "Term Year",
        "Term Month",
        "KOB Code",
        "Smoker",
        "Medical",
        "Claim Date",
        "Cause Of Claim",
        "Inw Book Year",
        "Inw Book Month",
        "Ced Book Year",
        "Ced Book Month",
        "Method of Payment",
        "Curr",
        "Reins Claim",
        "Your Share",
        "Reinsurer Name",
        "Voucher ID",
        "Out Voucher ID",
        "Voucher Desc"
    ]

    LOG_COLUMNS_OUTWARD = [
        "Seq No",
        "Department",
        "Biz Type",
        "Retro Type",
        "Inward VIN Ref",
        "Voucher No",
        "Account With",
        "Cedant Company",
        "PIC",
        "Product",
        "CBY",
        "CBM",
        "OBY",
        "OBM",
        "KOB",
        "COB",
        "MOP",
        "Curr",
        "Total Contribution",
        "Commission",
        "Overiding",
        "Total Commission",
        "Gross Premium Income",
        "Tabarru",
        "Ujrah",
        "Ujrah Spc",
        "Claim",
        "Balance",
        "Check Balance",
        "Rate Exchange",
        "Kontribusi (IDR)",
        "Commission (IDR)",
        "Overiding (IDR)",
        "Total Commission (IDR)",
        "Gross Premium Income (IDR)",
        "Tabarru (IDR)",
        "Ujrah (IDR)",
        "Ujrah Spc (IDR)",
        "Claim (IDR)",
        "Balance (IDR)",
        "Check Balance (IDR)",
        "REMARKS",
        "PML ID",
        "STATUS",
        "CREATED AT",
        "CREATED BY",
        "Due Date",
        "CANCELED AT",
        "CANCELED BY",
        "CANCEL OF VOUCHER",
        "CANCEL REASON"
    ]


    st.subheader("📊 Calculate PML")

    reins_type = st.selectbox(
        "Reinsurance Type",
        ["INWARD", "OUTWARD"],
        key="reins_type_calculate",
        index=0) #INWARD
        #disabled=True


    if reins_type == "INWARD":

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
        # SNAPSHOT LOG KE SESSION STATE
        # Proteksi dari upload user lain saat sedang memilih
        # ==========================
        log_snapshot_key = f"log_snapshot_inward_{year}_{month}"

        if log_snapshot_key not in st.session_state:
            st.session_state[log_snapshot_key] = log_df.copy()

        # Gunakan snapshot, bukan log_df langsung
        df_working = st.session_state[log_snapshot_key]

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
        df_posted = df_working[df_working["STATUS"] == "POSTED"].copy()

        if df_posted.empty:
            st.info("Tidak ada data dengan status POSTED")
            st.stop()

        # ==========================
        # REFRESH LOG PML
        # ==========================
        if st.button("🔄 Refresh Log PML", key="btn_refresh_pml"):
            if log_snapshot_key in st.session_state:
                del st.session_state[log_snapshot_key]
            st.cache_data.clear()
            st.rerun()

        # Warning di bawah tombol
        current_posted_count  = len(log_df[log_df["STATUS"] == "POSTED"])
        snapshot_posted_count = len(df_working[df_working["STATUS"] == "POSTED"])

        if current_posted_count != snapshot_posted_count:
            st.warning(
                f"⚠️ Log PML telah diperbarui oleh user lain "
                f"({snapshot_posted_count} → {current_posted_count} baris POSTED). "
                f"Klik Refresh jika ingin memuat data terbaru."
            )

        # ==========================
        # SEARCH (OPSIONAL)
        # ==========================
        search = st.text_input("🔍 Cari PML ID")

        if search:
            df_posted = df_posted[
                df_posted["PML ID"].astype(str).str.contains(search, case=False, na=False)
            ]

        # ==========================
        # FILTER
        # ==========================
        with st.expander("🔽 Filter Data", expanded=False):

            col_f1, col_f2, col_f3, col_f4 = st.columns(4)

            with col_f1:
                # Filter Cedant Company
                cedant_options = ["(Semua)"] + sorted(df_posted["Account With"].dropna().unique().tolist())
                filter_cedant = st.selectbox("Account With", cedant_options, key="filter_cedant")

            with col_f2:
                # ✅ Ganti selectbox → text_input untuk pencarian bebas
                filter_product = st.text_input(
                    "🔍 Product",
                    key="filter_product"
                )

            with col_f3:
                # Filter CBY
                cby_options = ["(Semua)"] + sorted(df_posted["CBY"].dropna().unique().tolist())
                filter_cby = st.selectbox("CBY", cby_options, key="filter_cby")

            with col_f4:
                # Filter CBM
                cbm_options = ["(Semua)"] + sorted(df_posted["CBM"].dropna().unique().tolist())
                filter_cbm = st.selectbox("CBM", cbm_options, key="filter_cbm")

        # Terapkan filter
        df_filtered = df_posted.copy()

        if filter_cedant != "(Semua)":
            df_filtered = df_filtered[df_filtered["Account With"] == filter_cedant]

        # ✅ Hanya filter jika ada keyword yang diketik (kosong = tampil semua)
        if filter_product.strip():  # <-- ini yang perlu diperbaiki
            df_filtered = df_filtered[
                df_filtered["Product"].astype(str).str.contains(filter_product.strip(), case=False, na=False)
            ]

        if filter_cby != "(Semua)":
            df_filtered = df_filtered[df_filtered["CBY"] == filter_cby]

        if filter_cbm != "(Semua)":
            df_filtered = df_filtered[df_filtered["CBM"] == filter_cbm]

        st.write(f"Total PML POSTED: {len(df_filtered)} {'(difilter)' if len(df_filtered) != len(df_posted) else ''}")

        # ==========================
        # UI SELECT
        # ==========================
        st.info("Centang pada kolom **'Pilih'** untuk menentukan baris yang akan diproses.")

        if not df_filtered.empty:

            # ==========================
            # SELECT ALL BUTTON
            # ==========================
            col_btn1, col_btn2, col_btn3 = st.columns([2.5, 3, 10])

            with col_btn1:
                select_all = st.button("✅ Select All", key="btn_select_all")

            with col_btn2:
                deselect_all = st.button("⬜ Deselect All", key="btn_deselect_all")

            # State untuk select all
            if "pilih_state" not in st.session_state:
                st.session_state["pilih_state"] = {}

            if select_all:
                for idx in df_filtered.index:
                    st.session_state["pilih_state"][idx] = True

            if deselect_all:
                for idx in df_filtered.index:
                    st.session_state["pilih_state"][idx] = False

            # Tambahkan checkbox column berdasarkan state
            df_to_edit = df_filtered.copy()
            df_to_edit.insert(
                0, "Pilih",
                df_to_edit.index.map(lambda i: st.session_state["pilih_state"].get(i, False))
            )

            # ==========================
            # CLEAN NUMERIC
            # ==========================
            cols_numeric = ["Total Contribution", "Gross Premium Income", "Tabarru", "Ujrah", "Claim", "Balance"]

            for col in cols_numeric:
                def clean_number(x):
                    x = str(x)
                    if "." in x and "," in x:
                        x = x.replace(".", "").replace(",", ".")
                    else:
                        x = x.replace(",", "")
                    return pd.to_numeric(x, errors="coerce")

                df_to_edit[col] = df_to_edit[col].apply(clean_number)

            # ==========================
            # DATA EDITOR
            # ==========================
            edited_df = st.data_editor(
                df_to_edit,
                column_config={
                    "Pilih": st.column_config.CheckboxColumn(
                        "Pilih",
                        help="Pilih baris ini untuk di-calculate",
                        default=False,
                    ),
                    "PML ID":  st.column_config.Column(disabled=True),
                    "STATUS":  st.column_config.Column(disabled=True),
                    "Product": st.column_config.Column(
                        "Product", disabled=True,
                        help="Baris pertama kolom References No pada file PML"
                    ),
                    "CBY": st.column_config.Column(
                        "CBY", disabled=True,
                        help="Baris pertama kolom CBY pada file PML"
                    ),
                    "CBM": st.column_config.Column(
                        "CBM", disabled=True,
                        help="Baris pertama kolom CBM pada file PML"
                    ),
                    "Total Contribution": st.column_config.NumberColumn(
                        "Total Contribution", format="%,.0f"
                    ),
                    "Gross Premium Income": st.column_config.NumberColumn(
                        "Gross Premium Income", format="%,.0f"
                    ),
                    "Tabarru": st.column_config.NumberColumn(
                        "Tabarru", format="%,.0f"
                    ),
                    "Ujrah": st.column_config.NumberColumn(
                        "Ujrah", format="%,.0f"
                    ),
                    "Claim": st.column_config.NumberColumn(
                        "Claim", format="%,.0f"
                    ),
                    "Balance": st.column_config.NumberColumn(
                        "Balance", format="%,.0f"
                    ),
                },
                disabled=[
                    "No", "PML ID", "STATUS", "Product", "CBY", "CBM",
                    "Total Contribution", "Gross Premium Income",
                    "Tabarru", "Ujrah", "Claim", "Balance"
                ],
                hide_index=True,
                use_container_width=True,
                key="data_editor_calculate"
            )

            # ==========================
            # SYNC STATE DARI EDITED DF
            # ==========================
            for idx, row in edited_df.iterrows():
                st.session_state["pilih_state"][idx] = row["Pilih"]

            # ==========================
            # AMBIL YANG DIPILIH
            # ==========================
            selected_rows = edited_df[edited_df["Pilih"] == True]

            if selected_rows.empty:
                st.info("Silakan pilih minimal satu baris untuk melanjutkan.")
                selected_rows = None
            else:
                st.success(f"✅ {len(selected_rows)} baris terpilih")

        
            # ==========================
            # CEK RATE FILE
            # ==========================
            if selected_rows is None or selected_rows.empty:
                selected_account = None

            else:
                # Ambil unique account
                unique_accounts = selected_rows["Account With"].dropna().unique()

                if len(unique_accounts) > 1:
                    st.error("❌ Account With harus sama untuk proses calculate")
                    selected_account = None
                else:
                    selected_account = unique_accounts[0]

            # ==========================
            # CEK RATE FILE DI DRIVE
            # ==========================
            rate_file_id = None

            if selected_account:
                rate_file_id = find_drive_file(
                    service=service,
                    filename=f"{selected_account}.xlsx",
                    parent_id=RATE_FOLDER_ID
                )

            has_rate = rate_file_id is not None

            # ==========================
            # LAYOUT BUTTON
            # ==========================
            col1, col2, col3, col4 = st.columns([1, 1, 1.5, 1.5])

            with col3:
                ceding_clicked = st.button(
                    "📥 Ceding Calculation",
                    use_container_width=True,
                    disabled=(selected_account is None),
                    type="primary"
                )

            with col4:
                our_clicked = st.button(
                    "🧮 Our Calculation",
                    use_container_width=True,
                    disabled=(selected_account is None or not has_rate),
                    help="Rate belum tersedia" if not has_rate else ""
                )

            # ==========================
            # INFO RATE
            # ==========================
            if selected_account is None:
                st.info("ℹ️ Pilih data terlebih dahulu (pastikan Account With sama)")

            elif not has_rate:
                st.info("ℹ️ Rate belum tersedia → Our Calculation dinonaktifkan")

            else:
                st.success(f"✅ Rate ditemukan untuk account: {selected_account}")


            # ==========================
            # POST VOUCHER (MULTI - LOCKED)
            # ==========================
            if ceding_clicked:
                with st.spinner("🔍 Validating PML..."):
                    start_time = time.time()

                    # ==========================
                    # INIT
                    # ==========================
                    validation_errors = []
                    validated_data = []

                    service = get_drive_service()

                    drive_folders = get_period_drive_folders(
                        year=int(year),
                        month=int(month),
                        root_folder_id=ROOT_DRIVE_FOLDER_ID
                    )

                    PERIOD_DRIVE_ID = drive_folders["period_id"]

                    # ==========================
                    # PREPARE GLOBAL
                    # ==========================
                    pml_drive = get_or_create_folder(
                        service=service,
                        folder_name="Folder PML",
                        parent_id=PERIOD_DRIVE_ID
                    )

                    PML_DRIVE_ID = pml_drive

                    ceding_folder_name = normalize_folder_name(selected_account)

                    ceding_drive = get_or_create_ceding_folders(
                        service=service,
                        period_folder_id=PERIOD_DRIVE_ID,
                        ceding_name=ceding_folder_name
                    )

                    CEDING_DRIVE_ID = ceding_drive["ceding_id"]

                    rate_exchange = get_exchange_rate(
                        service=service,
                        config_folder_id=CONFIG_FOLDER_ID,
                        currency=selected_rows.iloc[0]["Curr"],
                        month=month
                    )

                    due_date = calculate_due_date(
                        account_with=selected_account,
                        year=year,
                        month=month,
                        service=service
                    )

                    # ==========================
                    # LOAD / CREATE LOG
                    # ==========================
                    log_drive_id = find_drive_file(
                        service=service,
                        filename=get_log_filename(int(year), int(month)),
                        parent_id=PERIOD_DRIVE_ID,
                        mime_type="application/vnd.google-apps.spreadsheet"
                    )

                    if not log_drive_id:

                        log_drive_id = create_log_gsheet(
                            service=service,
                            parent_id=PERIOD_DRIVE_ID,
                            filename=get_log_filename(int(year), int(month)),
                            columns=LOG_COLUMNS
                        )

                    # ==========================
                    # VALIDATION STAGE
                    # ==========================
                    for _, row in selected_rows.iterrows():

                        try:

                            # ==========================
                            # GET PML FILE
                            # ==========================
                            pml_file_id = find_drive_file(
                                service=service,
                                filename=f"{row['PML ID']}.xlsx",
                                parent_id=PML_DRIVE_ID,
                                mime_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )

                            if not pml_file_id:

                                validation_errors.append(
                                    f"{row['PML ID']} → file tidak ditemukan"
                                )

                                continue

                            file_stream = download_file_from_drive(
                                service,
                                pml_file_id
                            )

                            df = pd.read_excel(file_stream)

                            # ==========================
                            # VALIDATE
                            # ==========================
                            errors = validate_calculate(
                                df,
                                row["Biz Type"],
                                reins_type
                            )

                            if errors:

                                validation_errors.append(
                                    f"{row['PML ID']} → {', '.join(errors)} (Kolom Tidak Unik)"
                                )

                                continue

                            # ==========================
                            # SAVE VALIDATED DATA
                            # ==========================
                            validated_data.append({
                                "row": row,
                                "df": df
                            })

                        except Exception as e:

                            validation_errors.append(
                                f"{row['PML ID']} → {e}"
                            )

                # ==========================
                # STOP IF VALIDATION FAILED
                # ==========================
                if validation_errors:

                    st.error("❌ Validation gagal")

                    for err in validation_errors:
                        st.write(f"- {err}")

                    st.stop()

                # ==========================
                # POSTING STAGE
                # ==========================
                with st.spinner("⏳ Calculation sedang berjalan, mohon tunggu..."):

                    try:

                        # 🔒 LOCK SEKALI SAJA
                        acquire_drive_lock(service, PERIOD_DRIVE_ID)

                        sheets_service = init_sheets_service(creds)

                        success_count = 0

                        # ==========================
                        # LOOP POSTING
                        # ==========================
                        for item in validated_data:

                            row = item["row"]
                            df = item["df"]

                            try:

                                biz_type = row["Biz Type"]

                                # ==========================
                                # GENERATE VOUCHER
                                # ==========================
                                voucher, seq_no, _ = generate_vin_from_drive(
                                    service=service,
                                    period_folder_id=PERIOD_DRIVE_ID,
                                    year=int(year),
                                    month=int(month),
                                    find_drive_file=find_drive_file,
                                    biz_type=biz_type
                                )

                                # ==========================
                                # BUILD LOG ENTRY
                                # ==========================
                                if biz_type in [
                                    "Kontribusi",
                                    "Refund",
                                    "Alteration",
                                    "Retur",
                                    "Revise",
                                    "Batal",
                                    "Cancel"
                                ]:

                                    total_contribution = df["Reins Total Premium"].sum()

                                    commission = df["Reins Comm"].sum()

                                    overriding = (
                                        df["Reins Overriding"].sum()
                                        if "Reins Overriding" in df.columns
                                        else 0
                                    )

                                    total_commission = commission + overriding

                                    claim_amount = (
                                        df["Claim"].sum()
                                        if "Claim" in df.columns
                                        else 0
                                    )

                                    balance = (
                                        total_contribution
                                        - total_commission
                                        - claim_amount
                                    )

                                    log_entry = {
                                        "Seq No": seq_no,
                                        "Department": row["Department"],
                                        "Biz Type": row["Biz Type"],
                                        "Voucher No": voucher,
                                        "Account With": row["Account With"],
                                        "Cedant Company": row["Cedant Company"],
                                        "PIC": row["PIC"],
                                        "Product": df["References No"].iloc[0],
                                        "CBY": df["CBY"].iloc[0],
                                        "CBM": df["CBM"].iloc[0],
                                        "OBY": int(year),
                                        "OBM": int(month),
                                        "KOB": df["K.O.B Code"].iloc[0],
                                        "COB": df["COB"].iloc[0],
                                        "MOP": df["Pay Period Type"].iloc[0],
                                        "Curr": df["Ccy Code"].iloc[0],

                                        "Total Contribution": total_contribution,
                                        "Commission": commission,
                                        "Overriding": overriding,
                                        "Total Commission": total_commission,
                                        "Gross Premium Income": total_contribution - total_commission,
                                        "Tabarru": df["Reins Tabarru"].sum(),
                                        "Ujrah": df["Reins Ujrah"].sum(),
                                        "Claim": 0,
                                        "Balance": balance,
                                        "Check Balance": "",

                                        "Rate Exchange": rate_exchange,

                                        "Kontribusi (IDR)": total_contribution * rate_exchange,
                                        "Commission (IDR)": commission * rate_exchange,
                                        "Overiding (IDR)": overriding * rate_exchange,
                                        "Total Commission (IDR)": total_commission * rate_exchange,
                                        "Gross Premium Income (IDR)": (
                                            total_contribution - total_commission
                                        ) * rate_exchange,

                                        "Tabarru (IDR)": (
                                            df["Reins Tabarru"].sum()
                                            * rate_exchange
                                        ),

                                        "Ujrah (IDR)": (
                                            df["Reins Ujrah"].sum()
                                            * rate_exchange
                                        ),

                                        "Claim (IDR)": 0,

                                        "Balance (IDR)": (
                                            balance * rate_exchange
                                        ),

                                        "Check Balance (IDR)": "",

                                        "REMARKS": "-",
                                        "PML ID": row["PML ID"],
                                        "STATUS": "POSTED",
                                        "CREATED AT": now_wib_naive(),
                                        "CREATED BY": row["PIC"],
                                        "Due Date": due_date,
                                        "Subject Email": row["Subject Email"],
                                        "Email Date": row["Email Date"],
                                        "CANCELED AT": "-",
                                        "CANCELED BY": "-",
                                        "CANCEL OF VOUCHER": "-",
                                        "CANCEL REASON": "-"
                                    }

                                elif biz_type == "Claim":

                                    claim_amount = (
                                        df["Marein Share IDR"].sum()
                                        if "Marein Share IDR" in df.columns
                                        else 0
                                    )

                                    balance = -claim_amount

                                    log_entry = {
                                        "Seq No": seq_no,
                                        "Department": row["Department"],
                                        "Biz Type": row["Biz Type"],
                                        "Voucher No": voucher,
                                        "Account With": row["Account With"],
                                        "Cedant Company": row["Cedant Company"],
                                        "PIC": row["PIC"],
                                        "Product": df["References No"].iloc[0],
                                        "CBY": df["CedBookYear"].iloc[0],
                                        "CBM": df["CedBookMonth"].iloc[0],
                                        "OBY": int(year),
                                        "OBM": int(month),
                                        "KOB": df["KindOfBusiness"].iloc[0],
                                        "COB": df["ClassOfBusiness"].iloc[0],
                                        "MOP": df["PayPeriodType"].iloc[0],
                                        "Curr": df["Currency"].iloc[0],

                                        "Total Contribution": 0,
                                        "Commission": 0,
                                        "Overriding": 0,
                                        "Total Commission": 0,
                                        "Gross Premium Income": 0,
                                        "Tabarru": 0,
                                        "Ujrah": 0,
                                        "Claim": claim_amount,
                                        "Balance": balance,
                                        "Check Balance": "",

                                        "Rate Exchange": rate_exchange,

                                        "Kontribusi (IDR)": 0,
                                        "Commission (IDR)": 0,
                                        "Overiding (IDR)": 0,
                                        "Total Commission (IDR)": 0,
                                        "Gross Premium Income (IDR)": 0,
                                        "Tabarru (IDR)": 0,
                                        "Ujrah (IDR)": 0,

                                        "Claim (IDR)": (
                                            claim_amount * rate_exchange
                                        ),

                                        "Balance (IDR)": (
                                            balance * rate_exchange
                                        ),

                                        "Check Balance (IDR)": "",

                                        "REMARKS": "-",
                                        "PML ID": row["PML ID"],
                                        "STATUS": "POSTED",
                                        "CREATED AT": now_wib_naive(),
                                        "CREATED BY": row["PIC"],
                                        "Due Date": due_date,
                                        "Subject Email": row["Subject Email"],
                                        "Email Date": row["Email Date"],
                                        "CANCELED AT": "-",
                                        "CANCELED BY": "-",
                                        "CANCEL OF VOUCHER": "-",
                                        "CANCEL REASON": "-"
                                    }

                                # ==========================
                                # APPEND LOG
                                # ==========================
                                append_gsheet(
                                    service=sheets_service,
                                    spreadsheet_id=log_drive_id,
                                    row_dict=log_entry
                                )

                                # ==========================
                                # UPLOAD VOUCHER FILE
                                # ==========================
                                upload_dataframe_to_drive(
                                    service=service,
                                    df=df,
                                    template_columns=(
                                        columns_template
                                        if biz_type != "Claim"
                                        else columns_template_claim
                                    ),
                                    voucher_id=voucher,
                                    filename=f"{voucher}.xlsx",
                                    folder_id=CEDING_DRIVE_ID,
                                    file_type="Voucher"
                                )

                                # ==========================
                                # UPDATE STATUS
                                # ==========================
                                update_pml_status_to_calculated(
                                    service=sheets_service,
                                    spreadsheet_id=log_pml_drive_id,
                                    pml_id=[str(row["PML ID"])]
                                )

                                success_count += 1

                            except Exception as e:

                                st.error(
                                    f"❌ Error posting PML {row['PML ID']}: {e}"
                                )

                        # ==========================
                        # DONE
                        # ==========================
                        if success_count > 0:

                            # Hapus snapshot agar next load ambil data fresh
                            if log_snapshot_key in st.session_state:
                                del st.session_state[log_snapshot_key]

                            # Reset pilih state
                            st.session_state["pilih_state"] = {}

                            end_time = time.time()
                            duration = int(end_time - start_time)

                            st.success(
                                f"✅ {success_count} voucher berhasil diposting "
                                f"({duration} detik)"
                            )


                    except RuntimeError:

                        st.error(
                            "⛔ Log sedang digunakan user lain. "
                            "Silakan coba lagi."
                        )

                    finally:

                        release_drive_lock(service, PERIOD_DRIVE_ID)

            elif our_clicked:
                with st.spinner("🔍 Validating PML..."):
                    start_time = time.time()

                    # ==========================
                    # INIT
                    # ==========================
                    validation_errors = []
                    validated_data = []

                    service = get_drive_service()

                    drive_folders = get_period_drive_folders(
                        year=int(year),
                        month=int(month),
                        root_folder_id=ROOT_DRIVE_FOLDER_ID
                    )

                    PERIOD_DRIVE_ID = drive_folders["period_id"]

                    # ==========================
                    # PREPARE GLOBAL
                    # ==========================
                    pml_drive = get_or_create_folder(
                        service=service,
                        folder_name="Folder PML",
                        parent_id=PERIOD_DRIVE_ID
                    )

                    PML_DRIVE_ID = pml_drive

                    ceding_folder_name = normalize_folder_name(selected_account)

                    ceding_drive = get_or_create_ceding_folders(
                        service=service,
                        period_folder_id=PERIOD_DRIVE_ID,
                        ceding_name=ceding_folder_name
                    )

                    CEDING_DRIVE_ID = ceding_drive["ceding_id"]

                    rate_exchange = get_exchange_rate(
                        service=service,
                        config_folder_id=CONFIG_FOLDER_ID,
                        currency=selected_rows.iloc[0]["Curr"],
                        month=month
                    )

                    due_date = calculate_due_date(
                        account_with=selected_account,
                        year=year,
                        month=month,
                        service=service
                    )

                    # ==========================
                    # LOAD / CREATE LOG
                    # ==========================
                    log_drive_id = find_drive_file(
                        service=service,
                        filename=get_log_filename(int(year), int(month)),
                        parent_id=PERIOD_DRIVE_ID,
                        mime_type="application/vnd.google-apps.spreadsheet"
                    )

                    if not log_drive_id:

                        log_drive_id = create_log_gsheet(
                            service=service,
                            parent_id=PERIOD_DRIVE_ID,
                            filename=get_log_filename(int(year), int(month)),
                            columns=LOG_COLUMNS
                        )

                    # ==========================
                    # VALIDATION STAGE
                    # ==========================
                    for _, row in selected_rows.iterrows():

                        try:

                            # ==========================
                            # GET PML FILE
                            # ==========================
                            pml_file_id = find_drive_file(
                                service=service,
                                filename=f"{row['PML ID']}.xlsx",
                                parent_id=PML_DRIVE_ID,
                                mime_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )

                            if not pml_file_id:

                                validation_errors.append(
                                    f"{row['PML ID']} → file tidak ditemukan"
                                )

                                continue

                            file_stream = download_file_from_drive(
                                service,
                                pml_file_id
                            )

                            df = pd.read_excel(file_stream)

                            # ==========================
                            # VALIDATE
                            # ==========================
                            errors = validate_calculate(
                                df,
                                row["Biz Type"],
                                reins_type
                            )

                            if errors:

                                validation_errors.append(
                                    f"{row['PML ID']} → {', '.join(errors)} (Kolom Tidak Unik)"
                                )

                                continue

                            # ==========================
                            # SAVE VALIDATED DATA
                            # ==========================
                            validated_data.append({
                                "row": row,
                                "df": df
                            })

                        except Exception as e:

                            validation_errors.append(
                                f"{row['PML ID']} → {e}"
                            )

                # ==========================
                # STOP IF VALIDATION FAILED
                # ==========================
                if validation_errors:

                    st.error("❌ Validation gagal")

                    for err in validation_errors:
                        st.write(f"- {err}")

                    st.stop()

                # ==========================
                # POSTING STAGE
                # ==========================
                with st.spinner("⏳ Calculation sedang berjalan, mohon tunggu..."):

                    try:
                        
                        # ==========================
                        # LOAD RATE FILE
                        # ==========================
                        rate_stream = download_file_from_drive(service, rate_file_id)

                        rate_df = pd.read_excel(rate_stream)

                        rate_df.columns = rate_df.columns.str.strip()

                        # ==========================
                        # NORMALIZE RATE COLUMNS
                        # ==========================
                        rate_df["Gender"]           = rate_df["Gender"].astype(str).str.strip().str.upper()
                        rate_df["Smoker"]           = rate_df["Smoker"].astype(str).str.strip().str.upper()
                        rate_df["Ced Product Code"] = rate_df["Ced Product Code"].astype(str).str.strip().str.upper()
                        rate_df["Age At"]           = pd.to_numeric(rate_df["Age At"], errors="coerce").fillna(0).astype(int)
                        rate_df["Rate"]             = pd.to_numeric(rate_df["Rate"], errors="coerce")
                        rate_df["Effective Start"]  = pd.to_datetime(rate_df["Effective Start"], errors="coerce", dayfirst=True)
                        rate_df["Effective End"]    = pd.to_datetime(rate_df["Effective End"], errors="coerce", dayfirst=True)

                        # ✅ TAMBAHKAN DI SINI — sebelum loop
                        review_results = []

                        # ==========================
                        # LOOP EACH PML
                        # ==========================
                        for validated in validated_data:

                            row = validated["row"]
                            df  = validated["df"]

                            review_df = df.copy()

                            # ==========================
                            # NORMALIZE PML COLUMNS
                            # ==========================
                            review_df["Gender"]           = review_df["Gender"].astype(str).str.strip().str.upper()
                            review_df["Smoker"]           = review_df["Smoker"].astype(str).str.strip().str.upper()
                            review_df["Ced Product Code"] = review_df["Ced Product Code"].astype(str).str.strip().str.upper()
                            review_df["Age At"]           = pd.to_numeric(review_df["Age At"], errors="coerce").fillna(0).astype(int)
                            review_df["Issue Date"]       = pd.to_datetime(review_df["Issue Date"], errors="coerce", dayfirst=True)

                            # ==========================
                            # INSERT CALC COLUMNS
                            # ==========================
                            calc_pairs = [
                                "Reins Premium",
                                "Reins EM Premium",
                                "Reins ER Premium",
                                "Reins Oth. Premium",
                                "Reins Total Premium",
                                "Reins Overriding",
                                "Reins Total Comm",
                                "Reins Tabarru",
                                "Reins Ujrah",
                                "Reins Nett Premium"
                            ]

                            premium_idx = review_df.columns.get_loc("Reins Premium")
                            if "Rate (Calc)" not in review_df.columns:
                                review_df.insert(premium_idx, "Rate (Calc)", None)

                            rate_calc_idx = review_df.columns.get_loc("Rate (Calc)")
                            if "Rate" not in review_df.columns:
                                review_df.insert(rate_calc_idx, "Rate", None)

                            for col in calc_pairs:
                                idx = review_df.columns.get_loc(col)
                                review_df.insert(idx + 1, f"{col} (Calc)", 0.0)

                            if "Calculation Status" not in review_df.columns:
                                review_df["Calculation Status"] = ""

                            # ==========================
                            # LOOP EACH ROW
                            # ==========================
                            for idx, data in review_df.iterrows():

                                gender       = str(data["Gender"]).strip().upper()
                                smoker       = str(data["Smoker"]).strip().upper()
                                product_code = str(data["Ced Product Code"]).strip().upper()
                                age          = int(pd.to_numeric(data["Age At"], errors="coerce") or 0)
                                issue_date   = data["Issue Date"]

                                # ==========================
                                # CARI RATE BERDASARKAN
                                # Gender, Smoker, Product Code, Age,
                                # dan Issue Date di antara Effective Start & End
                                # ==========================
                                matched_rate = rate_df[
                                    (rate_df["Gender"]           == gender) &
                                    (rate_df["Smoker"]           == smoker) &
                                    (rate_df["Ced Product Code"] == product_code) &
                                    (rate_df["Age At"]           == age) &
                                    (rate_df["Effective Start"]  <= issue_date) &
                                    (rate_df["Effective End"]    >= issue_date)
                                ]

                                if matched_rate.empty:
                                    review_df.at[idx, "Calculation Status"] = "RATE NOT FOUND"
                                    continue

                                rate = matched_rate.iloc[0]["Rate"]

                                review_df.at[idx, "Rate (Calc)"] = rate

                                # ==========================
                                # NUMERIC CONVERSION
                                # ==========================
                                sum_at_risk  = pd.to_numeric(data["Reins Sum At Risk"],  errors="coerce")
                                em_rate      = pd.to_numeric(data["Ced EM Rate"],        errors="coerce")
                                er_rate      = pd.to_numeric(data["Ced ER Rate"],        errors="coerce")
                                overriding   = pd.to_numeric(data["Reins Overriding"],   errors="coerce")
                                commission   = pd.to_numeric(data["Reins Total Comm"],   errors="coerce")
                                premium      = pd.to_numeric(data["Reins Premium"],      errors="coerce")
                                nett_premium = pd.to_numeric(data["Reins Nett Premium"], errors="coerce")

                                # ==========================
                                # TABARRU & UJRAH %
                                # ==========================
                                if nett_premium and nett_premium != 0:
                                    tabarru_percentage = data["Reins Tabarru"] / nett_premium
                                    ujrah_percentage   = data["Reins Ujrah"]   / nett_premium
                                else:
                                    tabarru_percentage = 0
                                    ujrah_percentage   = 0

                                # ==========================
                                # OVERRIDING %
                                # ==========================
                                reins_total_premium = pd.to_numeric(data["Reins Total Premium"], errors="coerce")

                                if pd.notna(overriding) and pd.notna(reins_total_premium) and reins_total_premium != 0:
                                    overriding_percentage = overriding / reins_total_premium
                                else:
                                    overriding_percentage = 0

                                # ==========================
                                # CALCULATION
                                # ==========================
                                rate_ced           = (premium / sum_at_risk) * 1000
                                premium_calc       = (sum_at_risk * rate) / 1000
                                em_premium_calc    = (premium_calc * em_rate) / 100   
                                er_premium_calc    = (sum_at_risk * er_rate) / 1000
                                total_premium_calc = premium_calc + em_premium_calc + er_premium_calc
                                overriding_calc    = total_premium_calc * overriding_percentage
                                total_comm_calc    = overriding_calc
                                nett_premium_calc  = total_premium_calc - total_comm_calc
                                tabarru_calc       = nett_premium_calc * tabarru_percentage
                                ujrah_calc         = nett_premium_calc * ujrah_percentage


                                # ==========================
                                # SAVE RESULT
                                # ==========================
                                review_df.at[idx, "Rate"]                       = rate_ced
                                review_df.at[idx, "Reins Premium (Calc)"]       = premium_calc
                                review_df.at[idx, "Reins EM Premium (Calc)"]    = em_premium_calc
                                review_df.at[idx, "Reins ER Premium (Calc)"]    = er_premium_calc
                                review_df.at[idx, "Reins Total Premium (Calc)"] = total_premium_calc
                                review_df.at[idx, "Reins Overriding (Calc)"]    = overriding_calc
                                review_df.at[idx, "Reins Total Comm (Calc)"]    = total_comm_calc
                                review_df.at[idx, "Reins Nett Premium (Calc)"]  = nett_premium_calc
                                review_df.at[idx, "Reins Tabarru (Calc)"]       = tabarru_calc
                                review_df.at[idx, "Reins Ujrah (Calc)"]         = ujrah_calc

                            # ==========================
                            # SUMMARY
                            # ==========================
                            total_original = (pd.to_numeric(review_df["Reins Nett Premium"], errors="coerce").fillna(0).sum())

                            total_calc = (pd.to_numeric(review_df["Reins Nett Premium (Calc)"],errors="coerce").fillna(0).sum())

                            total_diff = (total_calc - total_original)

                            missing_rate = len(review_df[review_df["Calculation Status"] == "RATE NOT FOUND"])

                            # ==========================
                            # CREATE REVIEW SPREADSHEET
                            # ==========================
                            review_spreadsheet_url = (create_review_spreadsheet(service=service, review_df=review_df, pml_id=row["PML ID"], parent_folder_id=PML_DRIVE_ID))

                            # ==========================
                            # SAVE RESULT
                            # ==========================
                            review_results.append({

                                "pml_id": row["PML ID"],

                                "review_df": review_df,

                                "spreadsheet_url":
                                    review_spreadsheet_url,

                                "total_rows":
                                    len(review_df),

                                "missing_rate":
                                    missing_rate,

                                "total_original":
                                    total_original,

                                "total_calc":
                                    total_calc,

                                "total_diff":
                                    total_diff,

                                "approved":
                                    False
                            })

                        # ==========================
                        # SAVE SESSION
                        # ==========================
                        st.session_state[
                            "review_results"
                        ] = review_results

                        # ==========================
                        # CUSTOM CSS
                        # ==========================
                        st.markdown("""
                        <style>

                        div[data-testid="stMetricValue"] {
                            font-size: 2rem;
                        }

                        div[data-testid="stMetricLabel"] {
                            font-size: 0.9rem;
                        }

                        .small-status {
                            font-size: 0.95rem !important;
                            padding: 0.3rem 0.6rem !important;
                        }

                        </style>
                        """, unsafe_allow_html=True)

                        # ==========================
                        # REVIEW RESULT SECTION
                        # ==========================
                        if "review_results" in st.session_state:

                            review_results = (
                                st.session_state["review_results"]
                            )

                            st.subheader(
                                "📋 Review Calculation Result"
                            )

                            st.caption(
                                "Review hasil calculation sebelum proses approval dan posting."
                            )

                            # ==========================
                            # LOOP RESULT
                            # ==========================
                            for result in review_results:

                                pml_id = result["pml_id"]

                                total_original = (
                                    result["total_original"]
                                )

                                total_calc = (
                                    result["total_calc"]
                                )

                                total_diff = (
                                    result["total_diff"]
                                )

                                missing_rate = (
                                    result["missing_rate"]
                                )

                                spreadsheet_url = (
                                    result["spreadsheet_url"]
                                )

                                review_df = (
                                    result["review_df"]
                                )

                                # ==========================
                                # STATUS
                                # ==========================
                                if missing_rate > 0:

                                    status = "⚠️ Missing Rate"

                                    status_type = "warning"

                                elif abs(total_diff) > 1:

                                    status = "❌ Difference Detected"

                                    status_type = "error"

                                else:

                                    status = "✅ Ready for Approval"

                                    status_type = "success"

                                # ==========================
                                # CONTAINER
                                # ==========================
                                with st.container(
                                    border=True
                                ):

                                    # ==========================
                                    # HEADER
                                    # ==========================
                                    col1, col2 = st.columns(
                                        [4,2]
                                    )

                                    with col1:

                                        st.markdown(
                                            f"## 📄 {pml_id}"
                                        )

                                    with col2:

                                        if status_type == "success":

                                            st.markdown(
                                                f"""
                                                <div class="small-status">
                                                    ✅ Ready
                                                </div>
                                                """,
                                                unsafe_allow_html=True
                                            )

                                        elif status_type == "warning":

                                            st.markdown(
                                                f"""
                                                <div class="small-status">
                                                    ⚠️ Missing Rate
                                                </div>
                                                """,
                                                unsafe_allow_html=True
                                            )

                                        else:

                                            st.markdown(
                                                f"""
                                                <div class="small-status">
                                                    
                                                </div>
                                                """,
                                                unsafe_allow_html=True
                                            )

                                    # ==========================
                                    # METRICS
                                    # ==========================
                                    metric1, metric2, metric3, metric4 = (
                                        st.columns(4)
                                    )

                                    metric1.metric(
                                        "Original",
                                        f"{total_original:,.2f}"
                                    )

                                    metric2.metric(
                                        "Calculated",
                                        f"{total_calc:,.2f}"
                                    )

                                    metric3.metric(
                                        "Difference",
                                        f"{total_diff:,.2f}"
                                    )

                                    metric4.metric(
                                        "Missing Rate",
                                        missing_rate
                                    )

                                    st.write("")

                                    # ==========================
                                    # ACTION BUTTONS
                                    # ==========================
                                    btn1, btn2 = st.columns(
                                        [2,1]
                                    )

                                    # ==========================
                                    # OPEN SPREADSHEET
                                    # ==========================
                                    with btn1:

                                        st.link_button(
                                            "📂 Open Review Spreadsheet",
                                            spreadsheet_url,
                                            use_container_width=True
                                        )

                                    # ==========================
                                    # APPROVE
                                    # ==========================
                                    with btn2:

                                        approve_clicked = (
                                            st.button(
                                                "✅ Approve",
                                                key=f"approve_{pml_id}",
                                                type="primary",
                                                use_container_width=True,
                                                disabled=(
                                                    missing_rate > 0
                                                    or
                                                    abs(total_diff) > 1
                                                )
                                            )
                                        )

                                        if approve_clicked:

                                            result["approved"] = True

                                            st.success(
                                                f"{pml_id} approved"
                                            )

                                    # ==========================
                                    # PREVIEW DATA
                                    # ==========================
                                    with st.expander(
                                        "🔍 Preview Review Data"
                                    ):

                                        st.dataframe(
                                            review_df,
                                            use_container_width=True
                                        )

                                    # ==========================
                                    # APPROVAL INFO
                                    # ==========================
                                    if missing_rate > 0:

                                        st.warning(
                                            "Masih terdapat RATE NOT FOUND."
                                        )

                                    elif abs(total_diff) > 1:

                                        st.error(
                                            "Difference terlalu besar."
                                        )

                                    else:

                                        st.success(
                                            "Review lolos validasi."
                                        )

                                    st.write("") 

                        # ==========================
                        # SUCCESS
                        # ==========================
                        st.success(
                            "✅ Calculation selesai"
                        )

                    finally:

                        release_drive_lock(service, PERIOD_DRIVE_ID)


    elif reins_type == "OUTWARD":

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
            folder_name="Folder PML (Outward)",
            parent_id=PERIOD_DRIVE_ID
        )

        PML_DRIVE_ID = pml_drive

        # ==========================
        # GET LOG PML FILE
        # ==========================
        log_pml_drive_id = find_drive_file(
            service=service,
            filename=f"{get_log_pml_filename(int(year), int(month))} (Outward)",
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
        # SNAPSHOT LOG KE SESSION STATE
        # ==========================
        log_snapshot_key = f"log_snapshot_outward_{year}_{month}"

        if log_snapshot_key not in st.session_state:
            st.session_state[log_snapshot_key] = log_df.copy()

        df_working = st.session_state[log_snapshot_key]

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
        df_posted = df_working[df_working["STATUS"] == "POSTED"].copy()

        if df_posted.empty:
            st.info("Tidak ada data dengan status POSTED")
            st.stop()

        # ==========================
        # REFRESH LOG PML
        # ==========================
        if st.button("🔄 Refresh Log PML", key="btn_refresh_pml"):
            if log_snapshot_key in st.session_state:
                del st.session_state[log_snapshot_key]
            st.cache_data.clear()
            st.rerun()

        # Warning di bawah tombol
        current_posted_count  = len(log_df[log_df["STATUS"] == "POSTED"])
        snapshot_posted_count = len(df_working[df_working["STATUS"] == "POSTED"])

        if current_posted_count != snapshot_posted_count:
            st.warning(
                f"⚠️ Log PML telah diperbarui oleh user lain "
                f"({snapshot_posted_count} → {current_posted_count} baris POSTED). "
                f"Klik Refresh jika ingin memuat data terbaru."
            )

        # ==========================
        # SEARCH (OPSIONAL)
        # ==========================
        search = st.text_input("🔍 Cari PML ID")

        if search:
            df_posted = df_posted[
                df_posted["PML ID"].astype(str).str.contains(search, case=False, na=False)
            ]

        # ==========================
        # FILTER
        # ==========================
        with st.expander("🔽 Filter Data", expanded=False):

            col_f1, col_f2, col_f3, col_f4 = st.columns(4)

            with col_f1:
                cedant_options = ["(Semua)"] + sorted(df_posted["Account With"].dropna().unique().tolist())
                filter_cedant = st.selectbox("Account With", cedant_options, key="filter_cedant_outward")

            with col_f2:
                filter_product = st.text_input("🔍 Product", key="filter_product_outward")

            with col_f3:
                cby_options = ["(Semua)"] + sorted(df_posted["CBY"].dropna().unique().tolist())
                filter_cby = st.selectbox("CBY", cby_options, key="filter_cby_outward")

            with col_f4:
                cbm_options = ["(Semua)"] + sorted(df_posted["CBM"].dropna().unique().tolist())
                filter_cbm = st.selectbox("CBM", cbm_options, key="filter_cbm_outward")

        # Terapkan filter
        df_filtered = df_posted.copy()

        if filter_cedant != "(Semua)":
            df_filtered = df_filtered[df_filtered["Account With"] == filter_cedant]

        if filter_product.strip():
            df_filtered = df_filtered[
                df_filtered["Product"].astype(str).str.contains(filter_product.strip(), case=False, na=False)
            ]

        if filter_cby != "(Semua)":
            df_filtered = df_filtered[df_filtered["CBY"] == filter_cby]

        if filter_cbm != "(Semua)":
            df_filtered = df_filtered[df_filtered["CBM"] == filter_cbm]

        st.write(f"Total PML POSTED: {len(df_filtered)} {'(difilter)' if len(df_filtered) != len(df_posted) else ''}")

        # ==========================
        # UI SELECT
        # ==========================
        st.info("Centang pada kolom **'Pilih'** untuk menentukan baris yang akan diproses.")

        if not df_filtered.empty:

            # ==========================
            # SELECT ALL / DESELECT ALL
            # ==========================
            col_btn1, col_btn2, col_btn3 = st.columns([2.5, 3, 10])

            with col_btn1:
                select_all = st.button("✅ Select All", key="btn_select_all_outward")

            with col_btn2:
                deselect_all = st.button("⬜ Deselect All", key="btn_deselect_all_outward")

            if "pilih_state_outward" not in st.session_state:
                st.session_state["pilih_state_outward"] = {}

            if select_all:
                for idx in df_filtered.index:
                    st.session_state["pilih_state_outward"][idx] = True

            if deselect_all:
                for idx in df_filtered.index:
                    st.session_state["pilih_state_outward"][idx] = False

            # Tambahkan checkbox column berdasarkan state
            df_to_edit = df_filtered.copy()
            df_to_edit.insert(
                0, "Pilih",
                df_to_edit.index.map(lambda i: st.session_state["pilih_state_outward"].get(i, False))
            )

            # ==========================
            # CLEAN NUMERIC
            # ==========================
            cols_numeric = ["Total Contribution", "Commission", "Overriding", "Total Commission", "Gross Premium Income", "Tabarru", "Ujrah", "Claim", "Balance"]

            for col in cols_numeric:
                def clean_number(x):
                    x = str(x)
                    if "." in x and "," in x:
                        x = x.replace(".", "").replace(",", ".")
                    else:
                        x = x.replace(",", "")
                    return pd.to_numeric(x, errors="coerce")
                df_to_edit[col] = df_to_edit[col].apply(clean_number)

            # ==========================
            # DATA EDITOR
            # ==========================
            edited_df = st.data_editor(
                df_to_edit,
                column_config={
                    "Pilih": st.column_config.CheckboxColumn(
                        "Pilih",
                        help="Pilih baris ini untuk di-calculate",
                        default=False,
                    ),
                    "PML ID":  st.column_config.Column(disabled=True),
                    "STATUS":  st.column_config.Column(disabled=True),
                    "Product": st.column_config.Column("Product", disabled=True),
                    "CBY":     st.column_config.Column("CBY", disabled=True),
                    "CBM":     st.column_config.Column("CBM", disabled=True),
                    "Total Contribution": st.column_config.NumberColumn("Total Contribution", format="%,.0f"),
                    "Commission":         st.column_config.NumberColumn("Commission",         format="%,.0f"),
                    "Overriding":         st.column_config.NumberColumn("Overriding",         format="%,.0f"),
                    "Total Commission":   st.column_config.NumberColumn("Total Commission",   format="%,.0f"),
                    "Gross Premium Income": st.column_config.NumberColumn("Gross Premium Income", format="%,.0f"),
                    "Tabarru": st.column_config.NumberColumn("Tabarru", format="%,.0f"),
                    "Ujrah":   st.column_config.NumberColumn("Ujrah",   format="%,.0f"),
                    "Claim":   st.column_config.NumberColumn("Claim",   format="%,.0f"),
                    "Balance": st.column_config.NumberColumn("Balance", format="%,.0f"),
                },
                disabled=[
                    "No", "PML ID", "STATUS", "Product", "CBY", "CBM",
                    "Total Contribution", "Commission", "Overriding",
                    "Total Commission", "Gross Premium Income",
                    "Tabarru", "Ujrah", "Claim", "Balance"
                ],
                hide_index=True,
                use_container_width=True,
                key="data_editor_calculate_outward"
            )

            # ==========================
            # SYNC STATE DARI EDITED DF
            # ==========================
            for idx, row in edited_df.iterrows():
                st.session_state["pilih_state_outward"][idx] = row["Pilih"]

            # ==========================
            # AMBIL YANG DIPILIH
            # ==========================
            selected_rows = edited_df[edited_df["Pilih"] == True]

            if selected_rows.empty:
                st.info("Silakan pilih minimal satu baris untuk melanjutkan.")
                selected_rows = None
            else:
                st.success(f"✅ {len(selected_rows)} baris terpilih")

        
            # ==========================
            # CEK RATE FILE
            # ==========================
            if selected_rows is None or selected_rows.empty:
                selected_account = None

            else:
                # Ambil unique account
                unique_accounts = selected_rows["Account With"].dropna().unique()

                if len(unique_accounts) > 1:
                    st.error("❌ Account With harus sama untuk proses calculate")
                    selected_account = None
                else:
                    selected_account = unique_accounts[0]

            # ==========================
            # CEK RATE FILE DI DRIVE
            # ==========================
            rate_file_id = None

            if selected_account:
                rate_file_id = find_drive_file(
                    service=service,
                    filename=f"{selected_account}.xlsx",
                    parent_id=RATE_FOLDER_ID
                )

            has_rate = rate_file_id is not None

            # ==========================
            # LAYOUT BUTTON
            # ==========================
            col1, col2, col3, col4 = st.columns([1, 1, 1.5, 1.5])

            with col3:
                ceding_clicked = st.button(
                    "📥 Ceding Calculation",
                    use_container_width=True,
                    disabled=(selected_account is None),
                    type="primary"
                )

            with col4:
                our_clicked = st.button(
                    "🧮 Our Calculation",
                    use_container_width=True,
                    disabled=(selected_account is None or not has_rate),
                    help="Rate belum tersedia" if not has_rate else ""
                )

            # ==========================
            # INFO RATE
            # ==========================
            if selected_account is None:
                st.info("ℹ️ Pilih data terlebih dahulu (pastikan Account With sama)")

            elif not has_rate:
                st.info("ℹ️ Rate belum tersedia → Our Calculation dinonaktifkan")

            else:
                st.success(f"✅ Rate ditemukan untuk account: {selected_account}")


            # ==========================
            # POST VOUCHER (MULTI - LOCKED)
            # ==========================
            if ceding_clicked:
                with st.spinner("🔍 Validating PML..."):
                    start_time = time.time()

                    # ==========================
                    # INIT
                    # ==========================
                    validation_errors = []
                    validated_data = []

                    service = get_drive_service()

                    drive_folders = get_period_drive_folders(
                        year=int(year),
                        month=int(month),
                        root_folder_id=ROOT_DRIVE_FOLDER_ID
                    )

                    PERIOD_DRIVE_ID = drive_folders["period_id"]

                    # ==========================
                    # PREPARE GLOBAL
                    # ==========================
                    outward_drive_folder = get_or_create_outward_folders(
                        service=service,
                        period_folder_id=PERIOD_DRIVE_ID,
                    )

                    OUTWARD_DRIVE_ID = outward_drive_folder["outward_id"]

                    pml_drive = get_or_create_folder(
                        service=service,
                        folder_name="Folder PML (Outward)",
                        parent_id=PERIOD_DRIVE_ID
                    )

                    PML_DRIVE_ID = pml_drive

                    ceding_folder_name = normalize_folder_name(selected_account)

                    ceding_drive = get_or_create_ceding_folders(
                        service=service,
                        period_folder_id=OUTWARD_DRIVE_ID,
                        ceding_name=ceding_folder_name
                    )

                    CEDING_DRIVE_ID = ceding_drive["ceding_id"]

                    rate_exchange = get_exchange_rate(
                        service=service,
                        config_folder_id=CONFIG_FOLDER_ID,
                        currency=selected_rows.iloc[0]["Curr"],
                        month=month
                    )

                    due_date = calculate_due_date(
                        account_with=selected_account,
                        year=year,
                        month=month,
                        service=service
                    )

                    # ==========================
                    # LOAD / CREATE LOG
                    # ==========================
                    log_drive_id = find_drive_file(
                        service=service,
                        filename=f"{get_log_filename(int(year), int(month))} (Outward)",
                        parent_id=OUTWARD_DRIVE_ID,
                        mime_type="application/vnd.google-apps.spreadsheet"
                    )

                    if not log_drive_id:

                        log_drive_id = create_log_gsheet(
                            service=service,
                            parent_id=OUTWARD_DRIVE_ID,
                            filename=f"{get_log_filename(int(year), int(month))} (Outward)",
                            columns=LOG_COLUMNS_OUTWARD
                        )

                    # ==========================
                    # VALIDATION STAGE
                    # ==========================
                    for _, row in selected_rows.iterrows():

                        try:

                            pml_file_id = find_drive_file(
                                service=service,
                                filename=f"{row["PML ID"]}.xlsx",
                                parent_id=PML_DRIVE_ID
                            )

                            if not pml_file_id:

                                validation_errors.append(
                                    f"{row['PML ID']} → file tidak ditemukan"
                                )

                                continue

                            file_stream = download_file_from_drive(
                                service,
                                pml_file_id
                            )

                            df = pd.read_excel(file_stream)

                            errors = validate_calculate(
                                df,
                                row["Biz Type"],
                                reins_type
                            )

                            if errors:

                                validation_errors.append(
                                    f"{row['PML ID']} → {', '.join(errors)} (Kolom Tidak Unik)"
                                )

                                continue

                            validated_data.append({
                                "row": row,
                                "df": df
                            })

                        except Exception as e:

                            validation_errors.append(
                                f"{row['PML ID']} → {e}"
                            )

                # ==========================
                # STOP IF VALIDATION FAILED
                # ==========================
                if validation_errors:

                    st.error("❌ Validation gagal")

                    for err in validation_errors:
                        st.write(f"- {err}")

                    st.stop()

                # ==========================
                # POSTING STAGE
                # ==========================
                with st.spinner("⏳ Calculation sedang berjalan, mohon tunggu..."):

                    try:

                        acquire_drive_lock(service, OUTWARD_DRIVE_ID)

                        sheets_service = init_sheets_service(creds)

                        success_count = 0

                        for item in validated_data:

                            row = item["row"]
                            df = item["df"]

                            try:

                                biz_type = row["Biz Type"]

                                # ==========================
                                # GENERATE VOUCHER
                                # ==========================
                                voucher, seq_no, _ = generate_vou_from_drive(
                                    service=service,
                                    period_folder_id=OUTWARD_DRIVE_ID,
                                    year=int(year),
                                    month=int(month),
                                    find_drive_file=find_drive_file,
                                    biz_type=biz_type
                                )

                                # ==========================
                                # BUILD LOG ENTRY
                                # ==========================
                                if biz_type in [
                                    "Kontribusi",
                                    "Refund",
                                    "Alteration",
                                    "Retur",
                                    "Revise",
                                    "Batal",
                                    "Cancel"
                                ]:

                                    total_contribution = df["Retro Total Premium"].sum()

                                    commission = df["Retro Total Comm"].sum()

                                    overriding = (
                                        df["Retro Overriding"].sum()
                                        if "Retro Overriding" in df.columns
                                        else 0
                                    )

                                    total_commission = commission + overriding

                                    claim_amount = (
                                        df["Claim"].sum()
                                        if "Claim" in df.columns
                                        else 0
                                    )

                                    balance = (
                                        total_contribution
                                        - total_commission
                                        - claim_amount
                                    )

                                    log_entry = {
                                        "Seq No": seq_no,
                                        "Department": row["Department"],
                                        "Biz Type": row["Biz Type"],
                                        "Retro Type": df["Retro Type"].iloc[0],
                                        "Inward VIN Ref": df["Inw Vouc ID"].iloc[0],
                                        "Voucher No": voucher,
                                        "Account With": row["Account With"],
                                        "Cedant Company": row["Cedant Company"],
                                        "PIC": row["PIC"],
                                        "Product": df["References No"].iloc[0],
                                        "CBY": df["Ced Book Year"].iloc[0],
                                        "CBM": df["Ced Book Month"].iloc[0],
                                        "OBY": int(year),
                                        "OBM": int(month),
                                        "KOB": df["KOB Code"].iloc[0],
                                        "COB": df["COB"].iloc[0],
                                        "MOP": df["Out Pay Period Type"].iloc[0],
                                        "Curr": df["Premium Ccy"].iloc[0],

                                        "Total Contribution": total_contribution,
                                        "Commission": commission,
                                        "Overiding": overriding,
                                        "Total Commission": total_commission,
                                        "Gross Premium Income": total_contribution - total_commission,
                                        "Tabarru": df["Retro Tabarru"].sum(),
                                        "Ujrah": df["Retro Ujrah"].sum(),
                                        "Claim": 0,
                                        "Balance": balance,
                                        "Check Balance": "",

                                        "Rate Exchange": rate_exchange,

                                        "Kontribusi (IDR)": total_contribution * rate_exchange,
                                        "Commission (IDR)": commission * rate_exchange,
                                        "Overiding (IDR)": overriding * rate_exchange,
                                        "Total Commission (IDR)": total_commission * rate_exchange,
                                        "Gross Premium Income (IDR)": (
                                            total_contribution - total_commission
                                        ) * rate_exchange,

                                        "Tabarru (IDR)": (
                                            df["Retro Tabarru"].sum()
                                            * rate_exchange
                                        ),

                                        "Ujrah (IDR)": (
                                            df["Retro Ujrah"].sum()
                                            * rate_exchange
                                        ),

                                        "Claim (IDR)": 0,

                                        "Balance (IDR)": (
                                            balance * rate_exchange
                                        ),

                                        "Check Balance (IDR)": "",

                                        "REMARKS": "-",
                                        "PML ID": row["PML ID"],
                                        "STATUS": "POSTED",
                                        "CREATED AT": now_wib_naive(),
                                        "CREATED BY": row["PIC"],
                                        "Due Date": due_date,
                                        "Subject Email": row["Subject Email"],
                                        "Email Date": row["Email Date"],
                                        "CANCELED AT": "-",
                                        "CANCELED BY": "-",
                                        "CANCEL OF VOUCHER": "-",
                                        "CANCEL REASON": "-"
                                    }

                                elif biz_type == "Claim":

                                    claim_amount = (
                                        df["Your Share"].sum()
                                        if "Your Share" in df.columns
                                        else 0
                                    )

                                    balance = -claim_amount

                                    log_entry = {
                                        "Seq No": seq_no,
                                        "Department": row["Department"],
                                        "Biz Type": row["Biz Type"],
                                        "Retro Type": df["Retro Type"].iloc[0],
                                        "Inward VIN Ref": df["Voucher ID"].iloc[0],
                                        "Voucher No": voucher,
                                        "Account With": row["Account With"],
                                        "Cedant Company": row["Cedant Company"],
                                        "PIC": row["PIC"],
                                        "Product": df["Voucher Desc"].iloc[0],
                                        "CBY": df["Ced Book Year"].iloc[0],
                                        "CBM": df["Ced Book Month"].iloc[0],
                                        "OBY": int(year),
                                        "OBM": int(month),
                                        "KOB": df["KOB Code"].iloc[0],
                                        "COB": df["COB Detail"].iloc[0],
                                        "MOP": df["Method of Payment"].iloc[0],
                                        "Curr": df["Curr"].iloc[0],

                                        "Total Contribution": 0,
                                        "Commission": 0,
                                        "Overriding": 0,
                                        "Total Commission": 0,
                                        "Gross Premium Income": 0,
                                        "Tabarru": 0,
                                        "Ujrah": 0,
                                        "Claim": claim_amount,
                                        "Balance": balance,
                                        "Check Balance": "",

                                        "Rate Exchange": rate_exchange,

                                        "Kontribusi (IDR)": 0,
                                        "Commission (IDR)": 0,
                                        "Overiding (IDR)": 0,
                                        "Total Commission (IDR)": 0,
                                        "Gross Premium Income (IDR)": 0,
                                        "Tabarru (IDR)": 0,
                                        "Ujrah (IDR)": 0,

                                        "Claim (IDR)": (
                                            claim_amount * rate_exchange
                                        ),

                                        "Balance (IDR)": (
                                            balance * rate_exchange
                                        ),

                                        "Check Balance (IDR)": "",

                                        "REMARKS": "-",
                                        "PML ID": row["PML ID"],
                                        "STATUS": "POSTED",
                                        "CREATED AT": now_wib_naive(),
                                        "CREATED BY": row["PIC"],
                                        "Due Date": due_date,
                                        "Subject Email": row["Subject Email"],
                                        "Email Date": row["Email Date"],
                                        "CANCELED AT": "-",
                                        "CANCELED BY": "-",
                                        "CANCEL OF VOUCHER": "-",
                                        "CANCEL REASON": "-"
                                    }

                                # ==========================
                                # APPEND LOG
                                # ==========================
                                append_gsheet(
                                    service=sheets_service,
                                    spreadsheet_id=log_drive_id,
                                    row_dict=log_entry
                                )

                                # ==========================
                                # UPLOAD FILE
                                # ==========================
                                upload_dataframe_to_drive_outward(
                                    service=service,
                                    df=df,
                                    template_columns=(
                                        columns_template_outward
                                        if biz_type != "Claim"
                                        else columns_template_claim_outward
                                    ),
                                    voucher_id=voucher,
                                    filename=f"{voucher}.xlsx",
                                    folder_id=CEDING_DRIVE_ID,
                                    biz_type=biz_type,
                                    pic=row["PIC"],
                                    date=now_wib_naive()
                                )

                                success_count += 1

                            except Exception as e:

                                st.error(
                                    f"❌ Error posting PML {row['PML ID']}: {e}"
                                )

                        # ==========================
                        # UPDATE STATUS
                        # ==========================
                        update_pml_status_to_calculated(
                            service=sheets_service,
                            spreadsheet_id=log_pml_drive_id,
                            pml_id=selected_rows["PML ID"].astype(str).tolist()
                        )

                        end_time = time.time()
                        duration = int(end_time - start_time)

                        if success_count > 0:

                            st.success(
                                f"✅ {success_count} voucher berhasil diposting "
                                f"({duration} detik)"
                            )

                    except RuntimeError:

                        st.error(
                            "⛔ Log sedang digunakan user lain. "
                            "Silakan coba lagi."
                        )

                    finally:

                        release_drive_lock(service, OUTWARD_DRIVE_ID)


