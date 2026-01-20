import pandas as pd
import numpy as np

REQUIRED_COLUMNS = [
    "certificate no",
    "insured full name",
    "gender",
    "pol holder no",
    "policy holder",
    "birth date",
    "age at",
    "issue date",
    "term year",
    "term month",
    "expired date",
    "medical",
    "ced product code",
    "ced coverage code",
    "ccy code",
    "sum insured",
    "sum at risk",
    "reins sum insured",
    "reins sum at risk",
    "pay period type",
    "reins total premium",
    "reins total comm",
    "reins tabarru",
    "reins ujrah",
    "reins nett premium",
    "valuation date"
]

DATE_COLUMNS = [
    "birth date",
    "issue date",
    "expired date",
    "valuation date"
]

NUMERIC_COLUMNS = [
    "sum insured",
    "sum at risk",
    "reins sum insured",
    "reins sum at risk",
    "reins total premium",
    "reins total comm",
    "reins tabarru",
    "reins ujrah",
    "reins nett premium"
]

INTEGER_COLUMNS = [
    "age at",
    "term year",
    "term month"
]


def validate_voucher(df):
    errors = []

    # =========================
    # NORMALISASI KOLOM
    # =========================
    df.columns = (
        df.columns
        .str.strip()
        .str.lower()
        .str.replace(r"\s+", " ", regex=True)
    )

    # =========================
    # 1. KOLOM WAJIB
    # =========================
    missing_cols = set(REQUIRED_COLUMNS) - set(df.columns)
    if missing_cols:
        errors.append(f"Kolom tidak ditemukan: {sorted(missing_cols)}")
        return errors  # stop total

    # =========================
    # 2. DATE VALIDATION
    # =========================
    for col in DATE_COLUMNS:
        converted = pd.to_datetime(df[col], errors="coerce")
        if converted.isna().any():
            errors.append(f"Kolom {col} harus bertipe tanggal (date)")
        df[col] = converted

    # =========================
    # 3. NUMERIC >= 0
    # =========================
    for col in NUMERIC_COLUMNS:
        numeric = pd.to_numeric(df[col], errors="coerce")
        if numeric.isna().any():
            errors.append(f"Kolom {col} harus numerik")
        elif (numeric < 0).any():
            errors.append(f"Kolom {col} tidak boleh bernilai negatif")
        df[col] = numeric

    # =========================
    # 4. GENDER
    # =========================
    if not df["gender"].isin(["M", "F", "U"]).all():
        errors.append("Kolom gender hanya boleh M, F, atau U")

    # =========================
    # 5. INTEGER >= 0 (AGE AT, TERM)
    # =========================
    for col in INTEGER_COLUMNS:
        numeric = pd.to_numeric(df[col], errors="coerce")

        if numeric.isna().any():
            errors.append(f"Kolom {col} harus berupa angka integer ≥ 0")
            continue

        if not (numeric % 1 == 0).all():
            errors.append(f"Kolom {col} tidak boleh mengandung desimal")
            continue

        if (numeric < 0).any():
            errors.append(f"Kolom {col} harus ≥ 0")
            continue

        # ✅ AMAN untuk casting
        df[col] = numeric.astype("Int64")

    # =========================
    # 6. TERM YEAR & TERM MONTH
    # =========================
    both_zero = (df["term year"] == 0) & (df["term month"] == 0)
    if both_zero.any():
        errors.append("term year dan term month tidak boleh keduanya bernilai 0")

    # =========================
    # 7. MEDICAL (M / N)
    # =========================
    df["medical"] = df["medical"].astype(str).str.strip().str.upper()

    if not df["medical"].isin(["M", "N"]).all():
        errors.append("Kolom medical hanya boleh bernilai M atau N")


    # =========================
    # 8. CCY CODE (3 HURUF)
    # =========================
    if not df["ccy code"].str.match(r"^[A-Z]{3}$").all():
        errors.append("ccy code harus 3 huruf kapital (contoh: IDR, USD)")

    # =========================
    # 9. EXPIRED DATE > ISSUE DATE
    # =========================
    if not (df["expired date"] > df["issue date"]).all():
        errors.append("expired date harus lebih besar dari issue date")

    # =========================
    # 10. REINS VS ORIGINAL LIMIT
    # =========================
    if not (df["reins sum insured"] <= df["sum insured"]).all():
        errors.append("reins sum insured tidak boleh lebih besar dari sum insured")

    if not (df["reins sum at risk"] <= df["sum at risk"]).all():
        errors.append("reins sum at risk tidak boleh lebih besar dari sum at risk")

    # =========================
    # 11. FINANCIAL CONSISTENCY (TOLERANSI)
    # =========================
    diff_nett = (
        df["reins total premium"]
        - df["reins total comm"]
        - df["reins nett premium"]
    ).abs()

    if not (diff_nett < 0.01).all():
        errors.append("reins nett premium ≠ total premium - total comm")

    diff_tab = (
        df["reins tabarru"]
        + df["reins ujrah"]
        - df["reins nett premium"]
    ).abs()

    if not (diff_tab < 0.01).all():
        errors.append("tabarru + ujrah ≠ reins nett premium")

    return errors
