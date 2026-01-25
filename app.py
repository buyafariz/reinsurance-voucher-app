import streamlit as st
import pandas as pd
from datetime import datetime
import os

from validator import validate_voucher
from vin_generator import generate_vin, create_cancel_row, get_log_path
from drive_utils import upload_or_update_drive_file, get_period_drive_folders
from lock_utils import acquire_lock, release_lock
from zoneinfo import ZoneInfo


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


BASE_PATH = "data"
ROOT_DRIVE_FOLDER_ID = st.secrets["drive_folder_id"]

st.title("📄 Reinsurance Voucher System")
st.write("")

tab_post, tab_cancel, tab_claim = st.tabs([
    "📥 Simpan VIN",
    "🚫 Cancel Voucher",
    "📄 Klaim"
])

# ==========================
# CLAIM
# ==========================

with tab_claim:
    st.subheader("📄 Klaim")

    uploaded_claim = st.file_uploader(
        "Upload File Klaim (.xlsx)",
        type=["xlsx"],
        key="upload_claim"
    )

    if uploaded_claim is None:
        st.info("Silakan upload file klaim")
    else:
        df_claim = pd.read_excel(uploaded_claim)
        st.success("File klaim berhasil dibaca")
        st.dataframe(df_claim, use_container_width=True)


# ==========================
# SIMPAN VOUCHER
# ==========================

with tab_post:
    st.subheader("📥 Simpan VIN (Posting Voucher)")

    business_event = st.radio("Jenis Transaksi", options=["NEW BUSINESS", "TERMINATED"], horizontal=True)
    business_event_code = ("NEW" if business_event == "NEW BUSINESS" else "TERMINATED")

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
        errors = validate_voucher(df, business_event_code)

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
            display_df = filtered_df.copy()

            for col in ACCOUNTING_COLS:
                if col in display_df.columns:
                    display_df[col] = display_df[col].apply(accounting_format)

            st.caption(f"Menampilkan {len(display_df):,} baris")

            st.dataframe(
                display_df.style.set_properties(
                    subset=[c for c in ACCOUNTING_COLS if c in display_df.columns],
                    **{"text-align": "right"}
                ),
                height=450,
                use_container_width=True
            )



        # ==========================
        # PERIOD & LOG
        # ==========================
        today = datetime.today()
        year, month = today.year, today.month

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
        VOUCHER_DRIVE_ID = drive_folders["voucher_folder_id"]


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
                        "ASTRA AVIVA LIFE",
                        "AVRIST ASSURANCE SYARIAH",
                        "AXA MANDIRI FINANCIAL SERVICES SYARIAH",
                        "BNI LIFE SYARIAH",
                        "BRINGIN LIFE SYARIAH",
                        "CAPITAL LIFE SYARIAH",
                        "CENTRAL ASIA RAYA SYARIAH",
                        "FWD LIFE INDONESIA SYARIAH",
                        "GENERALI INDONESIA LIFE ASSURANCE SYARIAH",
                        "GREAT EASTERN LIFE SYARIAH",
                        "JASA MITRA ABADI SYARIAH",
                        "MANULIFE INDONESIA SYARIAH",
                        "MEGA LIFE INSURANCE SYARIAH",
                        "PFI",
                        "PRUDENTIAL LIFE SYARIAH",
                        "RELIANCE SYARIAH",
                        "SINARMAS SYARIAH",
                        "SUN LIFE SYARIAH",
                    ]
                )

                pic = st.selectbox("PIC", ["Ardelia", "Buya", "Khansa"])
                product = st.text_input("Product")

            with col2:
                cby = st.selectbox("CBY", list(range(2015, year + 1)))
                cbm = st.selectbox("CBM", list(range(1, 13)))
                st.text_input("OBY", value=year, disabled=True)
                st.text_input("OBM", value=month, disabled=True)

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

            lock_path = log_path + ".lock"

            try:
                acquire_lock(lock_path)

                # reload log terbaru setelah lock
                if os.path.exists(log_path):
                    log_df = pd.read_excel(log_path)
                else:
                    log_df = pd.DataFrame()

                vin, seq_no, _ = generate_vin(BASE_PATH, year, month)

                local_folder = f"{year}_{month:02d}/vouchers"
                os.makedirs(os.path.join(BASE_PATH, local_folder), exist_ok=True)

                voucher_path = os.path.join(BASE_PATH, local_folder, f"{vin}.xlsx")
                df.to_excel(voucher_path, index=False)

                # Upload voucher (selalu CREATE)
                upload_or_update_drive_file(
                    file_path=voucher_path,
                    filename=f"{vin}.xlsx",
                    folder_id=VOUCHER_DRIVE_ID
                )

                if business_event_code == "NEW":
                    entry_type = "POST"
                elif business_event_code == "TERMINATED":
                    entry_type = "TERMINATE"

                now_wib = datetime.now(ZoneInfo("Asia/Jakarta"))
                now_wib_naive = now_wib.replace(tzinfo=None)


                log_entry = {
                    "Seq No": seq_no,
                    "VIN No": vin,
                    "Account With": account_with,
                    "PIC": pic,
                    "Product": product,
                    "CBY": cby,
                    "CBM": cbm,
                    "OBY": year,
                    "OBM": month,
                    "COB": cob,
                    "MOP": mop,
                    "Total Contribution": df["reins total premium"].sum(),
                    "Commission": df["reins total comm"].sum(),
                    "Tabarru": df["reins tabarru"].sum(),
                    "Ujrah": df["reins ujrah"].sum(),
                    "Overiding": df["overiding"].sum() if "overiding" in df.columns else 0,
                    "Claim": df["claim"].sum() if "claim" in df.columns else 0,
                    "Balance": df["reins nett premium"].sum(),
                    "REMARKS": remarks,
                    "BUSINESS EVENT": business_event_code,
                    "STATUS": "POSTED",
                    "ENTRY_TYPE": entry_type,
                    "CREATED_AT": now_wib_naive,
                    "CREATED_BY": pic,
                }

                log_df = pd.concat([log_df, pd.DataFrame([log_entry])], ignore_index=True)
                log_df.to_excel(log_path, index=False)

                # Upload / update log (SATU FILE)
                if "log_drive_id" not in st.session_state:
                    st.session_state["log_drive_id"] = upload_or_update_drive_file(
                        file_path=log_path,
                        filename="log_produksi.xlsx",
                        folder_id=PERIOD_DRIVE_ID
                    )
                else:
                    upload_or_update_drive_file(
                        file_path=log_path,
                        filename="log_produksi.xlsx",
                        folder_id=PERIOD_DRIVE_ID,
                        file_id=st.session_state["log_drive_id"]
                    )

                st.success(f"✅ Voucher berhasil diposting: {vin}")
                st.code(vin)

            finally:
                release_lock(lock_path)


with tab_cancel:
    st.subheader("🚫 Cancel Voucher")

    today = datetime.today()
    year, month = today.year, today.month
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
            (log_df["STATUS"] == "POSTED") &
            (log_df["ENTRY_TYPE"] == "POST")
        ]

        if posted_df.empty:
            st.info("Tidak ada voucher POSTED")
        else:
            selected_vin = st.selectbox(
                "Pilih VIN",
                posted_df["VIN No"].tolist()
            )

            cancel_reason = st.text_area("Alasan Pembatalan (WAJIB)")

            if st.button("❌ Batalkan Voucher"):

                if not cancel_reason.strip():
                    st.error("Alasan pembatalan wajib diisi")
                    st.stop()

                original_row = log_df[
                    log_df["VIN No"] == selected_vin
                ].iloc[0]

                cancel_vin, cancel_seq, _ = generate_vin(BASE_PATH, year, month)

                cancel_row = create_cancel_row(
                    original_row=original_row,
                    new_vin=cancel_vin,
                    seq_no=cancel_seq,
                    user=pic,
                    reason=cancel_reason
                )

                log_df.loc[
                    log_df["VIN No"] == selected_vin,
                    ["STATUS", "CANCELLED_AT", "CANCELLED_BY", "CANCEL_REASON"]
                ] = ["CANCELLED", now_wib_naive, pic, cancel_reason]

                log_df = pd.concat(
                    [log_df, pd.DataFrame([cancel_row])],
                    ignore_index=True
                )

                log_df.to_excel(log_path, index=False)

                upload_or_update_drive_file(
                    file_path=log_path,
                    filename="log_produksi.xlsx",
                    folder_id=PERIOD_DRIVE_ID,
                    file_id=st.session_state["log_drive_id"]
                )

                st.success(
                    f"Voucher {selected_vin} dibatalkan → VIN cancel {cancel_vin}"
                )
                st.rerun()


