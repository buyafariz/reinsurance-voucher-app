import streamlit as st
import pandas as pd
from datetime import datetime
import os

from validator import validate_voucher
from vin_generator import generate_vin, create_cancel_row, get_log_path
from drive_utils import upload_or_update_drive_file


# ==========================
# CONFIG
# ==========================
st.set_page_config(
    page_title="Reinsurance Voucher System",
    layout="centered"
)

BASE_PATH = "data"
DRIVE_FOLDER_ID = st.secrets["drive_folder_id"]

st.title("📄 Reinsurance Voucher System")
st.write("Upload voucher Excel untuk divalidasi, diposting, dan dicancel")


# ==========================
# UPLOAD FILE
# ==========================
uploaded_file = st.file_uploader(
    "Upload Voucher (.xlsx)",
    type=["xlsx"]
)

if not uploaded_file:
    st.stop()


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
errors = validate_voucher(df)

if errors:
    st.error("❌ VALIDASI GAGAL")
    for err in errors:
        st.write(f"- {err}")
    st.stop()

st.success("✅ Validasi berhasil")

with st.expander("📊 Preview Data Voucher"):
    st.dataframe(df, height=450)


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
# POST VOUCHER
# ==========================
if st.button("💾 Simpan Voucher"):

    if not product.strip() or not remarks.strip():
        st.error("Product dan Remarks wajib diisi")
        st.stop()

    vin, seq_no, _ = generate_vin(BASE_PATH, year, month)

    folder = f"{year}_{month:02d}/vouchers"
    os.makedirs(os.path.join(BASE_PATH, folder), exist_ok=True)

    voucher_path = os.path.join(BASE_PATH, folder, f"{vin}.xlsx")
    df.to_excel(voucher_path, index=False)

    # Upload voucher (SELALU CREATE)
    upload_or_update_drive_file(
        file_path=voucher_path,
        filename=f"{vin}.xlsx",
        folder_id=DRIVE_FOLDER_ID
    )

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
        "STATUS": "POSTED",
        "ENTRY_TYPE": "POST",
        "CREATED_AT": datetime.now(),
        "CREATED_BY": pic,
    }

    log_df = pd.concat([log_df, pd.DataFrame([log_entry])], ignore_index=True)
    log_df.to_excel(log_path, index=False)

    # Upload / update log (SATU FILE)
    if "log_drive_id" not in st.session_state:
        st.session_state["log_drive_id"] = upload_or_update_drive_file(
            file_path=log_path,
            filename="log_produksi.xlsx",
            folder_id=DRIVE_FOLDER_ID
        )
    else:
        upload_or_update_drive_file(
            file_path=log_path,
            filename="log_produksi.xlsx",
            folder_id=DRIVE_FOLDER_ID,
            file_id=st.session_state["log_drive_id"]
        )

    st.success(f"✅ Voucher berhasil diposting: {vin}")
    st.code(vin)


# ==========================
# CANCEL VOUCHER
# ==========================
st.divider()
st.subheader("🚫 Cancel Voucher")

if log_df.empty:
    st.info("Belum ada voucher")
    st.stop()

posted_df = log_df[
    (log_df["STATUS"] == "POSTED") &
    (log_df["ENTRY_TYPE"] == "POST")
]

if posted_df.empty:
    st.info("Tidak ada voucher POSTED")
    st.stop()

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
    ] = ["CANCELLED", datetime.now(), pic, cancel_reason]

    log_df = pd.concat(
        [log_df, pd.DataFrame([cancel_row])],
        ignore_index=True
    )

    log_df.to_excel(log_path, index=False)

    upload_or_update_drive_file(
        file_path=log_path,
        filename="log_produksi.xlsx",
        folder_id=DRIVE_FOLDER_ID,
        file_id=st.session_state["log_drive_id"]
    )

    st.success(
        f"Voucher {selected_vin} dibatalkan → VIN cancel {cancel_vin}"
    )
    st.rerun()
