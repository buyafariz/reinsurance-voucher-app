import streamlit as st
import pandas as pd
from datetime import datetime
import os

from validator import validate_voucher
from vin_generator import generate_vin, create_cancel_row, get_log_path
from drive_utils import upload_or_update_drive_file, get_period_drive_folders, get_or_create_ceding_folders, get_drive_service, find_drive_file, acquire_drive_lock, release_drive_lock
from lock_utils import acquire_lock, release_lock
from zoneinfo import ZoneInfo


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



# ==========================
# CONFIG
# ==========================
st.set_page_config(
    page_title="Reinsurance Voucher System",
    layout="centered"
)

if "log_period" not in st.session_state:
    now = datetime.now(ZoneInfo("Asia/Jakarta"))
    st.session_state["log_period"] = {
        "year": now.year,
        "month": now.month
    }


BASE_PATH = "data"
ROOT_DRIVE_FOLDER_ID = st.secrets["drive_folder_id"]

st.title("📄 Reinsurance Voucher System")
st.write("")

tab_post, tab_cancel = st.tabs([
    "📥 Create Voucher",
    "🚫 Cancel Voucher",
])

# ==========================
# CLAIM
# ==========================

# with tab_claim:
#     st.subheader("📄 Klaim")

#     uploaded_claim = st.file_uploader(
#         "Upload File Klaim (.xlsx)",
#         type=["xlsx"],
#         key="upload_claim"
#     )

#     if uploaded_claim is None:
#         st.info("Silakan upload file klaim")
#     else:
#         df_claim = pd.read_excel(uploaded_claim)
#         st.success("File klaim berhasil dibaca")
#         st.dataframe(df_claim, use_container_width=True)


# ==========================
# SIMPAN VOUCHER
# ==========================

with tab_post:
    st.subheader("📥 Create Voucher")
    
    col1, col2 = st.columns(2)

    with col1:
        department = st.selectbox(
                "Department",
                [
                    "ADMIN",
                    "CLAIM"
                ]
        )
    with col2:
        biz_type = st.selectbox(
            "Biz Type",
            [
                "Kontribusi",
                "Claim",
                "Refund",
                "Alteration",
                "Retur",
                "Revise",
                "Batal"
            ]
        )

   
    #business_event = st.radio("Jenis Transaksi", options=["NEW BUSINESS", "TERMINATED"], horizontal=True)
    #business_event_code = ("NEW" if business_event == "NEW BUSINESS" else "TERMINATED")

    uploaded_file = st.file_uploader(
        "Upload Voucher (.xlsx)",
        type=["xlsx"],
        key="upload_post"
    )

    if uploaded_file:

        # ==========================
        # READ FILE
        # ==========================
        df = pd.read_excel(uploaded_file)
        df.columns = df.columns.str.strip().str.lower()

        for col in ["certificate no", "pol holder no"]:
            if col in df.columns:
                df[col] = df[col].astype(str).str.strip()

        # ==========================
        # VALIDATION
        # ==========================
        errors = validate_voucher(df, biz_type)

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

            valid_columns = get_non_empty_columns(filtered_df)

            col1, col2 = st.columns([2, 4])

            with col1:
                filter_col = st.selectbox(
                    "Filter berdasarkan kolom",
                    options=valid_columns
                )

            col_series = filtered_df[filter_col]

            with col2:

                # TEXT FILTER
                if col_series.dtype == "object":
                    keyword = st.text_input(
                        f"Cari pada kolom `{filter_col}`"
                    )

                    if keyword:
                        filtered_df = filtered_df[
                            col_series.astype(str)
                            .str.contains(keyword, case=False, na=False)
                        ]

                # NUMERIC FILTER
                elif pd.api.types.is_numeric_dtype(col_series):
                    numeric_series = col_series.dropna()

                    if numeric_series.empty:
                        st.info(f"Kolom `{filter_col}` kosong")

                    else:
                        unique_vals = numeric_series.unique()

                        # 🚫 HANYA 1 NILAI UNIK → TIDAK BOLEH SLIDER
                        if len(unique_vals) == 1:
                            st.info(
                                f"Kolom `{filter_col}` hanya memiliki satu nilai: "
                                f"**{unique_vals[0]}**"
                            )

                        else:
                            is_integer = pd.api.types.is_integer_dtype(numeric_series)

                            if is_integer:
                                min_val = int(numeric_series.min())
                                max_val = int(numeric_series.max())

                                selected_range = st.slider(
                                    f"Range `{filter_col}`",
                                    min_value=min_val,
                                    max_value=max_val,
                                    value=(min_val, max_val),
                                    step=1
                                )
                            else:
                                min_val = float(numeric_series.min())
                                max_val = float(numeric_series.max())

                                selected_range = st.slider(
                                    f"Range `{filter_col}`",
                                    min_value=min_val,
                                    max_value=max_val,
                                    value=(min_val, max_val)
                                )

                            filtered_df = filtered_df[
                                col_series.between(*selected_range)
                            ]


                # DATETIME FILTER
                elif pd.api.types.is_datetime64_any_dtype(col_series):
                    valid_dates = col_series.dropna()

                    start_date, end_date = st.date_input(
                        f"Range tanggal `{filter_col}`",
                        value=(valid_dates.min(), valid_dates.max())
                    )

                    filtered_df = filtered_df[
                        col_series.between(
                            pd.to_datetime(start_date),
                            pd.to_datetime(end_date)
                        )
                    ]

            st.caption(f"Menampilkan {len(filtered_df):,} baris")

            # ==========================
            # DISPLAY DF (ACCOUNTING VIEW)
            # ==========================

            MAX_PREVIEW_ROWS = 2000

            display_df = filtered_df.copy()

            for col in ACCOUNTING_COLS:
                if col in display_df.columns:
                    display_df[col] = display_df[col].apply(accounting_format)

            total_rows = len(display_df)

            if total_rows > MAX_PREVIEW_ROWS:
                st.warning(
                    f"⚠️ Data sangat besar ({total_rows:,} baris). "
                    f"Hanya menampilkan {MAX_PREVIEW_ROWS:,} baris pertama untuk preview."
                )
                preview_df = display_df.head(MAX_PREVIEW_ROWS)
            else:
                preview_df = display_df

            st.caption(f"Menampilkan {len(preview_df):,} dari {total_rows:,} baris")

            st.dataframe(
                preview_df,
                height=450,
                use_container_width=True
            )



        # ==========================
        # PERIOD & LOG
        # ==========================
        year = st.session_state["log_period"]["year"]
        month = st.session_state["log_period"]["month"]

        log_path = get_log_path(BASE_PATH, year, month)

        if os.path.exists(log_path):
            log_df = pd.read_excel(log_path)
        else:
            log_df = pd.DataFrame()


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


                pic = st.selectbox("PIC", ["Ardelia", "Buya", "Khansa"])
                product = st.text_input("Product")

            with col2:
                years = list(range(2015, year + 1))
                months = list(range(1, 13))

                cby = st.selectbox("CBY", years, index=years.index(year))
                cbm = st.selectbox("CBM", months)#, index=months.index(month))
                st.text_input("OBY", value=year, disabled=True)
                st.text_input("OBM", value=month, disabled=True)

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

            remarks = st.text_area("Remarks (WAJIB)")


        # ==========================
        # FINANCIAL SUMMARY
        # ==========================
        st.subheader("💰 Ringkasan Finansial")

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

        st.dataframe(
            summary_df.style.format({"Nilai": "{:,.2f}"}),
            use_container_width=True
        )


        # ==========================
        # POST VOUCHER (LOCKED)
        # ==========================
        if st.button("💾 Simpan Voucher"):

            if not product.strip() or not remarks.strip():
                st.error("Product dan Remarks wajib diisi")
                st.stop()

            #lock_path = log_path + ".lock"
            service = get_drive_service()

            with st.spinner("⏳ Menyimpan voucher, mohon tunggu..."):

                try:
                    acquire_drive_lock(service, PERIOD_DRIVE_ID)

                    # reload log terbaru setelah lock
                    if os.path.exists(log_path):
                        log_df = pd.read_excel(log_path)
                    else:
                        log_df = pd.DataFrame()

                    voucher, seq_no, _ = generate_vin(BASE_PATH, year, month)

                    ceding_folder_name = normalize_folder_name(account_with)

                    local_folder = os.path.join(
                        f"{year}_{month:02d}",
                        ceding_folder_name
                        #"vouchers"
                    )

                    os.makedirs(os.path.join(BASE_PATH, local_folder), exist_ok=True)

                    voucher_path = os.path.join(BASE_PATH, local_folder, f"{voucher}.xlsx")
                    df.to_excel(voucher_path, index=False)


                    service = get_drive_service()

                    ceding_folder_name = normalize_folder_name(account_with)

                    ceding_drive = get_or_create_ceding_folders(
                        service=service,
                        period_folder_id=PERIOD_DRIVE_ID,
                        ceding_name=ceding_folder_name
                    )

                    CEDING_DRIVE_ID = ceding_drive["ceding_id"]


                    # Upload voucher (selalu CREATE)
                    upload_or_update_drive_file(
                        file_path=voucher_path,
                        filename=f"{voucher}.xlsx",
                        folder_id=CEDING_DRIVE_ID
                    )

                    #if business_event_code == "NEW":
                    #    entry_type = "POST"
                    #elif business_event_code == "TERMINATED":
                    #    entry_type = "TERMINATE"

                    rate_exchange = 1 if curr == "IDR" else (1000 if curr == "USD" else 0)

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
                        "OBY": year,
                        "OBM": month,
                        "KOB": kob,
                        "COB": cob,
                        "MOP": mop,
                        "Curr":curr,
                        "Total Contribution": df["reins total premium"].sum(),
                        "Commission": df["reins total comm"].sum(),
                        "Overiding": df["overiding"].sum() if "overiding" in df.columns else 0,
                        "Total Commission": (df["reins total comm"].sum()) + (df["overiding"].sum() if "overiding" in df.columns else 0),
                        "Gross Premium Income": df["reins total premium"].sum() - ((df["reins total comm"].sum()) + (df["overiding"].sum() if "overiding" in df.columns else 0)),
                        "Tabarru": df["reins tabarru"].sum(),
                        "Ujrah": df["reins ujrah"].sum(),
                        "Claim": df["claim"].sum() if "claim" in df.columns else 0,
                        "Balance": df["reins total premium"].sum() - df["reins total comm"].sum() - (df["overiding"].sum() if "overiding" in df.columns else 0) - (df["claim"].sum() if "claim" in df.columns else 0),
                        "Rate Exchange": rate_exchange,
                        "Kontribusi (IDR)": (df["reins total premium"].sum())*rate_exchange,
                        "Commission (IDR)": (df["reins total comm"].sum())*rate_exchange,
                        "Overiding (IDR)": (df["overiding"].sum() if "overiding" in df.columns else 0)*rate_exchange,
                        "Total Commission (IDR)": ((df["reins total comm"].sum()) + (df["overiding"].sum() if "overiding" in df.columns else 0))*rate_exchange,
                        "Gross Premium Income (IDR)": (df["reins total premium"].sum() - ((df["reins total comm"].sum()) + (df["overiding"].sum() if "overiding" in df.columns else 0)))*rate_exchange,
                        "Tabarru (IDR)": (df["reins tabarru"].sum())*rate_exchange,
                        "Ujrah (IDR)": (df["reins ujrah"].sum())*rate_exchange,
                        "Claim (IDR)": (df["claim"].sum() if "claim" in df.columns else 0)*rate_exchange,
                        "Balance": (df["reins total premium"].sum() - df["reins total comm"].sum() - (df["overiding"].sum() if "overiding" in df.columns else 0) - (df["claim"].sum() if "claim" in df.columns else 0))*rate_exchange,
                        "REMARKS": remarks,
                        "STATUS": "POSTED",
                        #"ENTRY_TYPE": entry_type,
                        "CREATED_AT": now_wib_naive(),
                        "CREATED_BY": pic,
                    }

                    log_df = pd.concat([log_df, pd.DataFrame([log_entry])], ignore_index=True)
                    log_df.to_excel(log_path, index=False)

                    # Upload / update log (SATU FILE)
                    service = get_drive_service()

                    log_drive_id = find_drive_file(
                        service=service,
                        filename="log_produksi.xlsx",
                        parent_id=PERIOD_DRIVE_ID
                    )

                    if log_drive_id:
                        upload_or_update_drive_file(
                            file_path=log_path,
                            filename="log_produksi.xlsx",
                            folder_id=PERIOD_DRIVE_ID,
                            file_id=log_drive_id
                        )
                    else:
                        upload_or_update_drive_file(
                            file_path=log_path,
                            filename="log_produksi.xlsx",
                            folder_id=PERIOD_DRIVE_ID
                        )

                    st.success(f"✅ Voucher berhasil diposting: {voucher}")
                    st.code(voucher)

                except RuntimeError as e:
                        st.error("⛔ Log sedang digunakan user lain. Silakan coba lagi.")
                        st.stop()

                finally:
                    release_drive_lock(service, PERIOD_DRIVE_ID)


with tab_cancel:
    st.subheader("🚫 Cancel Voucher")

    year = st.session_state["log_period"]["year"]
    month = st.session_state["log_period"]["month"]

    log_path = get_log_path(BASE_PATH, year, month)

    # 🔑 PASTIKAN log_df SELALU ADA
    if not os.path.exists(log_path):
        st.info("Belum ada voucher")
        st.stop()

    log_df = pd.read_excel(log_path)

    if log_df.empty:
        st.info("Belum ada voucher")
    else:
        posted_df = log_df[
            (log_df["STATUS"] == "POSTED") #&
            #(log_df["ENTRY_TYPE"] == "POST")
        ]

        if posted_df.empty:
            st.info("Tidak ada voucher POSTED")
        else:
            selected_voucher = st.selectbox(
                "Pilih Voucher",
                posted_df["Voucher No"].tolist()
            )

            pic = st.selectbox(
                "PIC",
                ["Ardelia", "Buya", "Khansa"],
                key="cancel pic"
            )

            cancel_reason = st.text_area("Alasan Cancel (WAJIB)")

            if st.button("❌ Cancel Voucher"):

                if not cancel_reason.strip():
                    st.error("Alasan Cancel wajib diisi")
                    st.stop()

                original_row = log_df[
                    log_df["Voucher No"] == selected_voucher
                ].iloc[0]

                service = get_drive_service()

                with st.spinner("⏳ Cancel voucher, mohon tunggu..."):

                    year = st.session_state["log_period"]["year"]
                    month = st.session_state["log_period"]["month"]

                    drive_folders = get_period_drive_folders(
                        year=year,
                        month=month,
                        root_folder_id=ROOT_DRIVE_FOLDER_ID
                    )

                    PERIOD_DRIVE_ID = drive_folders["period_id"]

                    ceding_folder_name = normalize_folder_name(original_row["Account With"])

                    ceding_drive = get_or_create_ceding_folders(
                        service=service,
                        period_folder_id=PERIOD_DRIVE_ID,
                        ceding_name=ceding_folder_name
                    )

                    CEDING_DRIVE_ID = ceding_drive["ceding_id"]

                    voucher_filename = f"{selected_voucher}.xlsx"

                    voucher_file_id = find_drive_file(
                        service=service,
                        filename=voucher_filename,
                        parent_id=CEDING_DRIVE_ID
                    )


                    cancel_voucher, cancel_seq, _ = generate_vin(BASE_PATH, year, month)

                    cancel_row = create_cancel_row(
                        original_row=original_row,
                        new_voucher=cancel_voucher,
                        seq_no=cancel_seq,
                        user=pic,
                        reason=cancel_reason
                    )

                    log_df.loc[
                        log_df["Voucher No"] == selected_voucher,
                        ["STATUS", "CANCELLED_AT", "CANCELLED_BY", "CANCEL_REASON"]
                    ] = ["CANCELED", now_wib_naive(), pic, cancel_reason]

                    log_df = pd.concat(
                        [log_df, pd.DataFrame([cancel_row])],
                        ignore_index=True
                    )

                    log_df.to_excel(log_path, index=False)

                    service = get_drive_service()

                    log_drive_id = find_drive_file(
                        service=service,
                        filename="log_produksi.xlsx",
                        parent_id=PERIOD_DRIVE_ID
                    )

                    if log_drive_id:
                        upload_or_update_drive_file(
                            file_path=log_path,
                            filename="log_produksi.xlsx",
                            folder_id=PERIOD_DRIVE_ID,
                            file_id=log_drive_id
                        )
                    else:
                        upload_or_update_drive_file(
                            file_path=log_path,
                            filename="log_produksi.xlsx",
                            folder_id=PERIOD_DRIVE_ID
                        )

                    if voucher_file_id:
                        service.files().delete(
                            fileId=voucher_file_id,
                            supportsAllDrives=True
                        ).execute()

                        st.success(f"File voucher {voucher_filename} berhasil dihapus dari Drive")
                    else:
                        st.warning("File voucher tidak ditemukan di Drive (mungkin sudah terhapus)")

                    st.success(
                        f"Voucher {selected_voucher} dibatalkan → VIN cancel {cancel_voucher}"
                    )
                st.rerun()


