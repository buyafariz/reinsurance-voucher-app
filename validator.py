import pandas as pd
import numpy as np


# Admin
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

# Claim
REQUIRED_COLUMNS_CLAIM = [
    "bookyear",
    "bookmonth",
    "cedbookyear",
    "cedbookmonth",
    "company name",
    "policy holder no",
    "policy holder",
    "certificate no",
    "insured name",
    "birth date",
    "age",
    "gender",
    "sum insured idr",
    "sum reinsured idr",
    "medicalcategory",
    "product",
    "coverage code",
    "classofbusiness",
    "payperiodtype",
    "issue date",
    "term year",
    "term month",
    "end date policy",
    "claim date",
    "claim register date",
    "payment date",
    "currency",
    "exchangerate",
    "amount of claim idr",
    "reins claim idr",
    "marein share idr",
    "cause of claim"
]


DATE_COLUMNS = [
    "birth date", # Admin
    "issue date",
    "expired date",
    "valuation date"
]

DATE_COLUMNS_CLAIM = [
    "birth date", # Claim
    "issue date",
    "end date policy",
    "claim date"
]

NUMERIC_COLUMNS = [
    "sum insured", # Admin
    "sum at risk",
    "reins sum insured",
    "reins sum at risk",
    "reins total premium",
    "reins total comm",
    "reins tabarru",
    "reins ujrah",
    "reins nett premium"
]

NUMERIC_COLUMNS_CLAIM = [
    "sum insured idr", # Claim
    "sum reinsured idr",
    "amount of claim idr",
    "reins claim idr",
    "marein share idr"
]


INTEGER_COLUMNS = [
    "age at", # Admin
    "term year",
    "term month"
]

INTEGER_COLUMNS_CLAIM = [
    "term year", # Claim
    "term month",
    "bookyear", 
    "bookmonth",
    "cedbookyear",
    "cedbookmonth",
    "age"
]


def validate_voucher(df, biz_type: str):

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

    biz_type = str(biz_type).strip()
    allowed = ["Kontribusi", "Claim", "Refund", "Alteration", "Retur", "Revise", "Batal"]

    if biz_type not in allowed:
        errors.append("BUSINESS TYPE tidak valid")
        return errors   # ⛔ stop sekali saja

    # =========================
    # 1. KOLOM WAJIB
    # =========================
    if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal"]:
        missing_cols = set(REQUIRED_COLUMNS) - set(df.columns)
        if missing_cols:
            errors.append(f"Kolom tidak ditemukan: {sorted(missing_cols)}")
            return errors  # stop total

    elif biz_type == "Claim":
        missing_cols = set(REQUIRED_COLUMNS_CLAIM) - set(df.columns)
        if missing_cols:
            errors.append(f"Kolom tidak ditemukan: {sorted(missing_cols)}")
            return errors  # stop total

    # =========================
    # 2. DATE VALIDATION
    # =========================
    if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal"]:
        for col in DATE_COLUMNS:
            converted = pd.to_datetime(df[col], errors="coerce")
            if converted.isna().any():
                errors.append(f"Kolom {col} harus bertipe tanggal (date)")
            df[col] = converted

    elif biz_type == "Claim":
        for col in DATE_COLUMNS_CLAIM:
            converted = pd.to_datetime(df[col], errors="coerce")
            if converted.isna().any():
                errors.append(f"Kolom {col} harus bertipe tanggal (date)")
            df[col] = converted
        

    # =========================
    # 3. NUMERIC VALIDATION (BY BUSINESS EVENT)
    # =========================


    if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal"]:

        for col in NUMERIC_COLUMNS:

            # 🔹 CEK DULU ADA ATAU TIDAK
            if col not in df.columns:
                errors.append(f"Kolom {col} tidak ditemukan di file")
                continue

            numeric = pd.to_numeric(df[col], errors="coerce")

            if numeric.isna().any():
                errors.append(f"Kolom {col} harus numerik")
                continue

            # 🔹 Kontribusi → tidak boleh negatif
            if biz_type == "Kontribusi":
                if (numeric < 0).any():
                    errors.append(
                        f"Kolom {col} tidak boleh bernilai negatif ({biz_type})"
                    )

            # 🔹 Refund, Retur, Batal → harus negatif
            if biz_type in ["Refund", "Retur", "Batal"]:
                if col in [
                    "reins total premium",
                    "reins total comm",
                    "reins tabarru",
                    "reins ujrah",
                    "reins nett premium"
                ]:
                    if (numeric > 0).any():
                        errors.append(
                            f"Kolom {col} harus bernilai negatif ({biz_type})"
                        )

            df[col] = numeric


    elif biz_type == "Claim":
        for col in NUMERIC_COLUMNS_CLAIM:
            numeric = pd.to_numeric(df[col], errors="coerce")

            if numeric.isna().any():
                errors.append(f"Kolom {col} harus numerik")
                continue

            # 🔹 Kontribusi → tidak boleh negatif
            if biz_type == "Claim":
                if (numeric < 0).any():
                    errors.append(
                        f"Kolom {col} tidak boleh bernilai negatif ({biz_type})"
                    )
    
            df[col] = numeric


    # =========================
    # 4. GENDER
    # =========================
    if not df["gender"].isin(["M", "F", "U"]).all():
        errors.append("Kolom gender hanya boleh M, F, atau U")

    # =========================
    # 5. INTEGER >= 0 (AGE AT, TERM)
    # =========================
    if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal"]:
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

    elif biz_type == "Claim":
        for col in INTEGER_COLUMNS_CLAIM:
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
    if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal"]:
        df["medical"] = df["medical"].astype(str).str.strip().str.upper()

        if not df["medical"].isin(["M", "N"]).all():
            errors.append("Kolom medical hanya boleh bernilai M atau N")

    elif biz_type == "Claim":
        df["medicalcategory"] = df["medicalcategory"].astype(str).str.strip().str.upper()

        if not df["medicalcategory"].isin(["M", "N"]).all():
            errors.append("Kolom medical hanya boleh bernilai M atau N")


    # =========================
    # 8. CCY CODE (3 HURUF)
    # =========================
    if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal"]:
        if not df["ccy code"].str.match(r"^[A-Z]{3}$").all():
            errors.append("ccy code harus 3 huruf kapital (contoh: IDR, USD)")

    elif biz_type == "Claim":
        if not df["currency"].str.match(r"^[A-Z]{3}$").all():
            errors.append("Currency harus 3 huruf kapital (contoh: IDR, USD)")

    # =========================
    # 9. EXPIRED DATE > ISSUE DATE
    # =========================
    if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal"]:
        if not (df["expired date"] > df["issue date"]).all():
            errors.append("expired date harus lebih besar dari issue date")

    elif biz_type == "Claim":
        if not (df["end date policy"] > df["issue date"]).all():
            errors.append("end date policy harus lebih besar dari issue date")      

    # =========================
    # 10. REINS VS ORIGINAL LIMIT
    # =========================
    if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal"]:
        if not (df["reins sum insured"] <= df["sum insured"]).all():
            errors.append("reins sum insured tidak boleh lebih besar dari sum insured")

        if not (df["reins sum at risk"] <= df["sum at risk"]).all():
            errors.append("reins sum at risk tidak boleh lebih besar dari sum at risk")

    elif biz_type == "Claim":
        if not (df["sum insured idr"] <= df["sum reinsured idr"]).all():
            errors.append("sum reinsured idr tidak boleh lebih besar dari sum insured idr")

        if not (df["amount of claim idr"] <= df["reins claim idr"]).all():
            errors.append("reins claim idr tidak boleh lebih besar dari amount of claim idr")

        if not (df["amount of claim idr"] <= df["marein share idr"]).all():
            errors.append("marein share idr tidak boleh lebih besar dari amount of claim idr")


    # =========================
    # 11. FINANCIAL CONSISTENCY (TOLERANSI)
    # =========================
    if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal"]:
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
