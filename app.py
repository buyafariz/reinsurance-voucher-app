import streamlit as st
import pandas as pd
from validator import validate_voucher
from vin_generator import generate_vin, create_cancel_row
from datetime import datetime
import os

# ==========================
# PAGE CONFIG
# ==========================
st.set_page_config(
    page_title="Voucher Upload",
    layout="centered"
)

st.title("📄 Reinsurance Voucher System")
st.write("Upload voucher Excel untuk divalidasi")

# ==========================
# UPLOAD FILE
# ==========================
uploaded_file = st.file_uploader(
    "Upload Voucher (.xlsx)",
    type=["xlsx"]
)

if uploaded_file:

    # ==========================
    # BACA FILE
    # ==========================
    df = pd.read_excel(uploaded_file)

    # Normalisasi kolom
    df.columns = df.columns.str.strip().str.lower()

    # Identifier wajib string
    for col in ["certificate no", "pol holder no"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    # ==========================
    # VALIDASI DATA
    # ==========================
    errors = validate_voucher(df)

    if errors:
        st.error("❌ VALIDASI GAGAL")
        for err in errors:
            st.write(f"- {err}")
        st.stop()

    st.success("✅ Validasi berhasil. File siap diproses.")

    with st.expander("📊 Lihat Detail Data Voucher"):
        st.dataframe(df, height=450)

    # ==========================
    # VIN & SESSION STATE
    # ==========================
    today = datetime.today()
    year = today.year
    month = today.month
    base_path = "data"

    if "vin" not in st.session_state:
        vin, seq_no, log_path = generate_vin(base_path, year, month)
        st.session_state.vin = vin
        st.session_state.seq_no = seq_no
        st.session_state.log_path = log_path
        st.session_state.saved = False

    vin = st.session_state.vin
    seq_no = st.session_state.seq_no
    log_path = st.session_state.log_path

    st.subheader("🔐 Voucher Information")
    st.write(f"**VIN No:** `{vin}`")

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
                ],
                key="account_with"
            )

            pic = st.selectbox(
                "PIC",
                ["Ardelia", "Buya", "Khansa"],
                key="pic"
            )

            product = st.text_input(
                "Product",
                key="product"
            )

        with col2:
            cby = st.selectbox(
                "Ceding Book Year (CBY)",
                list(range(2015, year + 1)),
                key="cby"
            )

            cbm = st.selectbox(
                "Ceding Book Month (CBM)",
                list(range(1, 13)),
                key="cbm"
            )

            st.text_input("Our Book Year (OBY)", value=year, disabled=True)
            st.text_input("Our Book Month (OBM)", value=month, disabled=True)

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
            ],
            key="cob"
        )

        mop = st.selectbox(
            "Mode of Payment (MOP)",
            [
                "Monthly",
                "Quarterly",
                "Half Yearly",
                "Yearly",
                "Single Premium",
            ],
            key="mop"
        )

        remarks = st.text_area(
            "Remarks",
            key="remarks"
        )

    # ==========================
    # RINGKASAN FINANSIAL
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
    # SIMPAN VOUCHER
    # ==========================
    if st.button("💾 Simpan Voucher", disabled=st.session_state.saved):

        validation_errors = []
        if not product.strip():
            validation_errors.append("Product wajib diisi")
        if not remarks.strip():
            validation_errors.append("Remarks wajib diisi")

        if validation_errors:
            st.warning("⚠️ Lengkapi data berikut:")
            for e in validation_errors:
                st.write(f"- {e}")
            st.stop()

        st.session_state.saved = True

        folder = f"{year}_{month:02d}/vouchers"
        os.makedirs(os.path.join("data", folder), exist_ok=True)
        file_path = os.path.join("data", folder, f"{vin}.xlsx")

        df.to_excel(file_path, index=False)

        log_entry = {
            "Seq No": seq_no,
            "VIN No": vin,
            "Account With": account_with,
            "PIC": pic,
            "PRODUCT": product,
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

            # 🔍 AUDIT TRAIL
            "STATUS": "POSTED",
            "CREATED_AT": datetime.now(),
            "CREATED_BY": pic
        }


        log_df = pd.read_excel(log_path)
        log_df = pd.concat([log_df, pd.DataFrame([log_entry])], ignore_index=True)
        log_df.to_excel(log_path, index=False)

        st.success(f"✅ Voucher `{vin}.xlsx` berhasil disimpan & log diperbarui")

    # ==========================
    # RESET
    # ==========================
    st.divider()
    if st.button("🔄 Upload Voucher Baru"):
        st.session_state.clear()
        st.rerun()

    st.divider()
    st.subheader("🚫 Cancel Voucher")

    log_df = pd.read_excel(log_path)

    posted_df = log_df[log_df["STATUS"] == "POSTED"]

    if posted_df.empty:
        st.info("Tidak ada voucher yang bisa dibatalkan.")
    else:
        selected_vin = st.selectbox(
            "Pilih VIN yang akan dibatalkan",
            posted_df["VIN No"].tolist()
        )

    cancel_reason = st.text_area(
        "Alasan Pembatalan (WAJIB)",
        placeholder="Contoh: Salah CBY / data premi tidak sesuai"
    )

    if st.button("❌ Batalkan Voucher"):

        if not cancel_reason.strip():
            st.error("Alasan pembatalan wajib diisi")
            st.stop()

        # ==========================
        # Ambil voucher original
        # ==========================
        original_row = log_df[log_df["VIN No"] == selected_vin]

        if original_row.empty:
            st.error("Voucher tidak ditemukan")
            st.stop()

        original_row = original_row.iloc[0]

        # ==========================
        # Generate VIN CANCEL
        # ==========================
        cancel_vin, cancel_seq, _ = generate_vin(base_path, year, month)

        # ==========================
        # Buat cancel row (NEGATIVE)
        # ==========================
        cancel_row = create_cancel_row(
            original_row=original_row,
            new_vin=cancel_vin,
            seq_no=cancel_seq,
            user=pic,
            reason=cancel_reason
        )

        # ==========================
        # Update voucher asli
        # ==========================
        log_df.loc[
            log_df["VIN No"] == selected_vin,
            ["STATUS", "CANCELLED_AT", "CANCELLED_BY", "CANCEL_REASON"]
        ] = [
            "CANCELLED",
            datetime.now(),
            pic,
            cancel_reason
        ]

        # ==========================
        # Append cancel entry
        # ==========================
        log_df = pd.concat(
            [log_df, pd.DataFrame([cancel_row])],
            ignore_index=True
        )

        log_df.to_excel(log_path, index=False)

        st.success(
            f"Voucher {selected_vin} berhasil dibatalkan "
            f"dengan VIN cancel {cancel_vin}"
        )

        st.rerun()
