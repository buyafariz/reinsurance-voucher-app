import streamlit as st
import pandas as pd
from datetime import datetime
import os

from validator import validate_voucher
from vin_generator import (
    generate_vin,
    create_cancel_row,
    get_log_path,
    load_or_create_log
)
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
st.write("Upload voucher Excel untuk divalidasi dan diposting")


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


# ==========================
# PERIOD & LOG
# ==========================
today = datetime.today()
year, month = today.year, today.month

log_path = get_log_path(BASE_PATH, year, month)
log_df = load_or_create_log(log_path)


# ==========================
# FORM INPUT
# ==========================
with st.expander("🧾 Informasi Voucher", expanded=True):
    account_with = st.text_input("Account With")
    pic = st.selectbox("PIC", ["Ardelia", "Buya", "Khansa"])
    product = st.text_input("Product")
    remarks = st.text_area("Remarks (WAJIB)")


# ==========================
# SAVE VOUCHER
# ==========================
if st.button("💾 Simpan Voucher"):

    if not product.strip() or not remarks.strip():
        st.error("Product dan Remarks wajib diisi")
        st.stop()

    vin, seq_no, _ = generate_vin(BASE_PATH, year, month)

    voucher_dir = os.path.join(BASE_PATH, f"{year}_{month:02d}", "vouchers")
    os.makedirs(voucher_dir, exist_ok=True)

    voucher_path = os.path.join(voucher_dir, f"{vin}.xlsx")
    df.to_excel(voucher_path, index=False)

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
        "REMARKS": remarks,
        "STATUS": "POSTED",
        "ENTRY_TYPE": "POST",
        "CREATED_AT": datetime.now(),
        "CREATED_BY": pic,
    }

    log_df = pd.concat([log_df, pd.DataFrame([log_entry])], ignore_index=True)
    log_df.to_excel(log_path, index=False)

    log_drive_id = st.session_state.get("log_drive_id")

    st.session_state["log_drive_id"] = upload_or_update_drive_file(
        file_path=log_path,
        filename="log_produksi.xlsx",
        folder_id=DRIVE_FOLDER_ID,
        file_id=log_drive_id
    )

    st.success(f"✅ Voucher berhasil diposting: {vin}")


# ==========================
# CANCEL VOUCHER
# ==========================
st.divider()
st.subheader("🚫 Cancel Voucher")

posted_df = log_df[
    (log_df["STATUS"] == "POSTED") &
    (log_df["ENTRY_TYPE"] == "POST")
]

if posted_df.empty:
    st.info("Tidak ada voucher POSTED")
    st.stop()

selected_vin = st.selectbox("Pilih VIN", posted_df["VIN No"].tolist())
cancel_reason = st.text_area("Alasan Pembatalan (WAJIB)")

if st.button("❌ Batalkan Voucher"):

    if not cancel_reason.strip():
        st.error("Alasan pembatalan wajib diisi")
        st.stop()

    original_row = log_df[log_df["VIN No"] == selected_vin].iloc[0]

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

    log_df = pd.concat([log_df, pd.DataFrame([cancel_row])], ignore_index=True)
    log_df.to_excel(log_path, index=False)

    upload_or_update_drive_file(
        file_path=log_path,
        filename="log_produksi.xlsx",
        folder_id=DRIVE_FOLDER_ID,
        file_id=st.session_state["log_drive_id"]
    )

    st.success(f"Voucher {selected_vin} dibatalkan → {cancel_vin}")
    st.rerun()
