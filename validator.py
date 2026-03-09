import pandas as pd
import numpy as np


# Admin
REQUIRED_COLUMNS_INWARD = [
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

# Admin
REQUIRED_COLUMNS_OUTWARD = [
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
    # "ri period from",
    # "ri period until",
    "medical",
    "ced product code",
    "ced coverage code",
    "ccy code",
    "sum insured",
    "sum at risk",
    "reins sum insured",
    "reins sum at risk",
    "retro sum insured",
    "retro sum at risk",
    "out pay period type",
    "retro total premium",
    "retro total comm",
    "retro tabarru",
    "retro ujrah",
    "retro overriding",
    "retro nett premium",
    "valuation date"
]

# Claim
REQUIRED_COLUMNS_CLAIM_INWARD = [
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

# Claim
REQUIRED_COLUMNS_CLAIM_OUTWARD = [
    "cedant name",
    "main pol no",
    "main policy",
    "pol holder no",
    "policy holder",
    "certificate no",
    "insured name",
    "birth date",
    "age",
    "gender",
    "ced product code",
    "ced coverage code",
    "cob detail",
    "issue date",
    "term year",
    "term month",
    "kob code",
    "smoker",
    "medical",
    "claim date",
    "cause of claim",
    "inw book year",
    "inw book month",
    "ced book year",
    "ced book month",
    "curr",
    "reins claim",
    "your share",
    "reinsurer name",
    "voucher id",
    "voucher desc"
]

# Admin
DATE_COLUMNS = [
    "birth date", 
    "issue date",
    # "ri period from",
    # "ri period until",
    "valuation date"
]

# Claim
DATE_COLUMNS_CLAIM_INWARD = [
    "birth date", 
    "issue date",
    "end date policy",
    "claim date"
]

DATE_COLUMNS_CLAIM_OUTWARD = [
    "birth date", 
    "issue date",
    "claim date"
]

# Admin
NUMERIC_COLUMNS_INWARD = [
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

NUMERIC_COLUMNS_OUTWARD = [
    "sum insured", 
    "sum at risk",
    "reins sum insured",
    "reins sum at risk",
    "retro sum insured",
    "retro sum at risk",
    "retro total premium",
    "retro total comm",
    "retro tabarru",
    "retro ujrah",
    "retro overriding",
    "retro nett premium"
]

#Claim
NUMERIC_COLUMNS_CLAIM_INWARD = [
    "sum insured idr",
    "sum reinsured idr",
    "amount of claim idr",
    "reins claim idr",
    "marein share idr"
]

NUMERIC_COLUMNS_CLAIM_OUTWARD = [
    "reins claim",
    "your share"
]

# Admin
INTEGER_COLUMNS = [
    "age at", 
    "term year",
    "term month"
]

# Claim
INTEGER_COLUMNS_CLAIM_INWARD = [
    "term year", 
    "term month",
    "bookyear", 
    "bookmonth",
    "cedbookyear",
    "cedbookmonth",
    "age"
]

INTEGER_COLUMNS_CLAIM_OUTWARD = [
    "term year", 
    "term month",
    "inw book year", 
    "inw book month",
    "ced book year",
    "ced book month",
    "age"
]


def validate_voucher(df, biz_type: str, reins_type:str):

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
    allowed = ["Kontribusi", "Claim", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]

    if biz_type not in allowed:
        errors.append("BUSINESS TYPE tidak valid")
        return errors   # ⛔ stop sekali saja

    # =========================
    # 1. KOLOM WAJIB
    # =========================
    if reins_type == "INWARD":
        if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
            missing_cols = set(REQUIRED_COLUMNS_INWARD) - set(df.columns)
            if missing_cols:
                errors.append(f"Kolom tidak ditemukan: {sorted(missing_cols)}")
                return errors  # stop total

        elif biz_type == "Claim":
            missing_cols = set(REQUIRED_COLUMNS_CLAIM_INWARD) - set(df.columns)
            if missing_cols:
                errors.append(f"Kolom tidak ditemukan: {sorted(missing_cols)}")
                return errors  # stop total
            
    elif reins_type == "OUTWARD":
        if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
            missing_cols = set(REQUIRED_COLUMNS_OUTWARD) - set(df.columns)
            if missing_cols:
                errors.append(f"Kolom tidak ditemukan: {sorted(missing_cols)}")
                return errors  # stop total

        elif biz_type == "Claim":
            missing_cols = set(REQUIRED_COLUMNS_CLAIM_OUTWARD) - set(df.columns)
            if missing_cols:
                errors.append(f"Kolom tidak ditemukan: {sorted(missing_cols)}")
                return errors  # stop total

    # =========================
    # 2. DATE VALIDATION
    # =========================
    if reins_type == "INWARD":
        if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
            for col in DATE_COLUMNS:
                converted = pd.to_datetime(df[col], errors="coerce")
                if converted.isna().any():
                    errors.append(f"Kolom {col} harus bertipe tanggal (date)")
                df[col] = converted

        elif biz_type == "Claim":
            for col in DATE_COLUMNS_CLAIM_INWARD:
                converted = pd.to_datetime(df[col], errors="coerce")
                if converted.isna().any():
                    errors.append(f"Kolom {col} harus bertipe tanggal (date)")
                df[col] = converted
        
    if reins_type == "OUTWARD":
        if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
            for col in DATE_COLUMNS:
                converted = pd.to_datetime(df[col], errors="coerce")
                if converted.isna().any():
                    errors.append(f"Kolom {col} harus bertipe tanggal (date)")
                df[col] = converted

        elif biz_type == "Claim":
            for col in DATE_COLUMNS_CLAIM_OUTWARD:
                converted = pd.to_datetime(df[col], errors="coerce")
                if converted.isna().any():
                    errors.append(f"Kolom {col} harus bertipe tanggal (date)")
                df[col] = converted

    # =========================
    # 3. NUMERIC VALIDATION (BY BUSINESS EVENT)
    # =========================
    if reins_type == "INWARD":
        if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
        
            for col in NUMERIC_COLUMNS_INWARD:

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
                if biz_type in ["Refund", "Retur", "Batal", "Cancel"]:
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
            for col in NUMERIC_COLUMNS_CLAIM_INWARD:
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

    elif reins_type == "OUTWARD":
        if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
    
            for col in NUMERIC_COLUMNS_OUTWARD:

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
                if biz_type in ["Refund", "Retur", "Batal", "Cancel"]:
                    if col in [
                        "retro total premium",
                        "retro total comm",
                        "retro tabarru",
                        "retro ujrah",
                        "retro overriding",
                        "retro nett premium"
                    ]:
                        if (numeric > 0).any():
                            errors.append(
                                f"Kolom {col} harus bernilai negatif ({biz_type})"
                            )

                df[col] = numeric


        elif biz_type == "Claim":
            for col in NUMERIC_COLUMNS_CLAIM_OUTWARD:
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
    if reins_type == "INWARD":
        if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
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
            for col in INTEGER_COLUMNS_CLAIM_INWARD:
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


    elif reins_type == "OUTWARD":
        if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
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
            for col in INTEGER_COLUMNS_CLAIM_OUTWARD:
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
    # both_zero = (df["term year"] == 0) & (df["term month"] == 0)
    # if both_zero.any():
    #     errors.append("term year dan term month tidak boleh keduanya bernilai 0")

    # =========================
    # 7. MEDICAL (M / N)
    # =========================
    if reins_type == "INWARD":
        if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
            df["medical"] = df["medical"].astype(str).str.strip().str.upper()

            if not df["medical"].isin(["M", "N"]).all():
                errors.append("Kolom medical hanya boleh bernilai M atau N")

        elif biz_type == "Claim":
            df["medicalcategory"] = df["medicalcategory"].astype(str).str.strip().str.upper()

            if not df["medicalcategory"].isin(["M", "N"]).all():
                errors.append("Kolom medical hanya boleh bernilai M atau N")

    elif reins_type == "OUTWARD":
        if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
            df["medical"] = df["medical"].astype(str).str.strip().str.upper()

            if not df["medical"].isin(["M", "N"]).all():
                errors.append("Kolom medical hanya boleh bernilai M atau N")

        elif biz_type == "Claim":
            df["medical"] = df["medical"].astype(str).str.strip().str.upper()

            if not df["medical"].isin(["M", "N"]).all():
                errors.append("Kolom medical hanya boleh bernilai M atau N")

    # =========================
    # 8. CCY CODE (3 HURUF)
    # =========================
    if reins_type == "INWARD":
        if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
            if not df["ccy code"].str.match(r"^[A-Z]{3}$").all():
                errors.append("ccy code harus 3 huruf kapital (contoh: IDR, USD)")

        elif biz_type == "Claim":
            if not df["currency"].str.match(r"^[A-Z]{3}$").all():
                errors.append("Currency harus 3 huruf kapital (contoh: IDR, USD)")

    elif reins_type == "OUTWARD":
        if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
            if not df["ccy code"].str.match(r"^[A-Z]{3}$").all():
                errors.append("ccy code harus 3 huruf kapital (contoh: IDR, USD)")

        elif biz_type == "Claim":
            if not df["curr"].str.match(r"^[A-Z]{3}$").all():
                errors.append("Currency harus 3 huruf kapital (contoh: IDR, USD)")

    # =========================
    # 9. EXPIRED DATE > ISSUE DATE
    # =========================
    if reins_type == "INWARD":
        if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
            if not (df["expired date"] > df["issue date"]).all():
                errors.append("expired date harus lebih besar dari issue date")

        elif biz_type == "Claim":
            if not (df["end date policy"] > df["issue date"]).all():
                errors.append("end date policy harus lebih besar dari issue date")      

    # elif reins_type == "OUTWARD":
    #     if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
    #         if not (df["ri period until"] > df["issue date"]).all():
    #             errors.append("ri period until harus lebih besar dari issue date")

        # elif biz_type == "Claim":
        #     if not (df["end date policy"] > df["issue date"]).all():
        #         errors.append("end date policy harus lebih besar dari issue date")      


    # =========================
    # 10. REINS VS ORIGINAL LIMIT
    # =========================
    if reins_type == "INWARD":
        if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
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

    if reins_type == "OUTWARD":
        if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
            if not (df["reins sum insured"] <= df["sum insured"]).all():
                errors.append("reins sum insured tidak boleh lebih besar dari sum insured")

            if not (df["retro sum insured"] <= df["reins sum insured"]).all():
                errors.append("retro sum insured tidak boleh lebih besar dari reins sum insured")

            if not (df["reins sum at risk"] <= df["sum at risk"]).all():
                errors.append("reins sum at risk tidak boleh lebih besar dari sum at risk")

            if not (df["retro sum at risk"] <= df["reins sum at risk"]).all():
                errors.append("retro sum at risk tidak boleh lebih besar dari reins sum at risk")

        elif biz_type == "Claim":
            if not (df["reins claim"] <= df["your share"]).all():
                errors.append("reins claim tidak boleh lebih besar dari your share")


    # =========================
    # 11. FINANCIAL CONSISTENCY (TOLERANSI)
    # =========================
    if reins_type == "INWARD":
        if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
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
        
    elif reins_type == "OUTWARD":
        if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
            diff_nett = (
                df["retro total premium"]
                - df["retro total comm"]
                - df["retro overriding"]
                - df["retro nett premium"]
            ).abs()

            if not (diff_nett < 0.01).all():
                errors.append("retro nett premium ≠ (total premium - total comm - overriding)")

            diff_tab = (
                df["retro tabarru"]
                + df["retro ujrah"]
                - df["retro nett premium"]
            ).abs()

            if not (diff_tab < 0.01).all():
                errors.append("tabarru + ujrah ≠ retro nett premium")

            return errors
