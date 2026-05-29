import pandas as pd
import numpy as np


# Admin
REQUIRED_COLUMNS_INWARD = [
    "trans category",
    "policy category",
    "certificate no",
    "insured full name",
    "gender",
    "main pol no",
    "moin policy",
    "pol holder no",
    "policy holder",
    "birth date",
    "age at",
    "issue date",
    "term year",
    "term month",
    "expired date",
    "medical",
    "smoker",
    "k.o.b code",
    "ced product code",
    "ced coverage code",
    "ccy code",
    "sum insured",
    "sum at risk",
    "reins sum insured",
    "reins sum at risk",
    "pay period type",
    "ced em rate",
    "ced er rate",
    "reins total premium",
    "reins total comm",
    "reins tabarru",
    "reins ujrah",
    "reins nett premium",
    "valuation date",
    "cby",
    "cbm",
    "cob",
    "voucher id",
    "references no"
]

# Admin
REQUIRED_COLUMNS_OUTWARD = [
    "retro type",
    "acc with name",
    "policy category",
    "kob code",
    "main pol no",
    "main policy",
    "pol holder no",
    "policy holder",
    "certificate no",
    "insured full name",
    "birth date",
    "gender",
    "issue date",
    "age at",
    "issue date",
    "term year",
    "term month",
    "expired date",
    "smoker",
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
    "valuation date",
    "inw vouc id",
    "retro type",
    "cob",
    "references no"
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
    "cause of claim",
    "voucher id",
    "references no"
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
    "voucher desc",
    "method of payment"
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
        errors.append("BUSINESS TYPE harus bernilai salah satu dari")

    # =========================
    # 1. KOLOM WAJIB
    # =========================
    if reins_type == "INWARD":
        if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
            missing_cols = set(REQUIRED_COLUMNS_INWARD) - set(df.columns)
            if missing_cols:
                errors.append(f"Kolom tidak ditemukan: {sorted(missing_cols)}")

        elif biz_type == "Claim":
            missing_cols = set(REQUIRED_COLUMNS_CLAIM_INWARD) - set(df.columns)
            if missing_cols:
                errors.append(f"Kolom tidak ditemukan: {sorted(missing_cols)}")
  
            
    elif reins_type == "OUTWARD":
        if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
            missing_cols = set(REQUIRED_COLUMNS_OUTWARD) - set(df.columns)
            if missing_cols:
                errors.append(f"Kolom tidak ditemukan: {sorted(missing_cols)}")


        elif biz_type == "Claim":
            missing_cols = set(REQUIRED_COLUMNS_CLAIM_OUTWARD) - set(df.columns)
            if missing_cols:
                errors.append(f"Kolom tidak ditemukan: {sorted(missing_cols)}")


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
    # if not df["gender"].isin(["M", "F", "U"]).all():
    #     errors.append("Kolom gender hanya boleh M, F, atau U")

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
    # 7. SMOKER (S / N)
    # =========================
    if reins_type == "INWARD":
        if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
            if "smoker" not in df.columns:
                errors.append("Tambahkan kolom Smoker")
            else:
                df["smoker"] = df["smoker"].astype(str).str.strip().str.upper()

                if not df["smoker"].isin(["S", "N"]).all():
                    errors.append("Kolom smoker hanya boleh bernilai S atau N")

        # elif biz_type == "Claim":
        #     df["medicalcategory"] = df["medicalcategory"].astype(str).str.strip().str.upper()

        #     if not df["medicalcategory"].isin(["M", "N"]).all():
        #         errors.append("Kolom medical hanya boleh bernilai M atau N")

    # elif reins_type == "OUTWARD":
    #     if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
    #         df["medical"] = df["medical"].astype(str).str.strip().str.upper()

    #         if not df["medical"].isin(["M", "N"]).all():
    #             errors.append("Kolom medical hanya boleh bernilai M atau N")

    #     elif biz_type == "Claim":
    #         df["medical"] = df["medical"].astype(str).str.strip().str.upper()

    #         if not df["medical"].isin(["M", "N"]).all():
    #             errors.append("Kolom medical hanya boleh bernilai M atau N")


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

            if not (df["sum at risk"] <= df["sum insured"]).all():
                errors.append("sum at risk tidak boleh lebih besar dari sum insured")

            if not (df["reins sum at risk"] <= df["reins sum insured"]).all():
                errors.append("reins sum at risk tidak boleh lebih besar dari reins sum insured")

        elif biz_type == "Claim":
            if not (df["sum insured idr"] >= df["sum reinsured idr"]).all():
                errors.append("sum reinsured idr tidak boleh lebih besar dari sum insured idr")

            if not (df["amount of claim idr"] >= df["reins claim idr"]).all():
                errors.append("reins claim idr tidak boleh lebih besar dari amount of claim idr")

            if not (df["amount of claim idr"] >= df["marein share idr"]).all():
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
            if not (df["reins claim"] >= df["your share"]).all():
                errors.append("your share tidak boleh lebih besar dari reins claim")


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

            # diff_tab = (
            #     df["retro tabarru"]
            #     + df["retro ujrah"]
            #     - df["retro nett premium"]
            # ).abs()

            # if not (diff_tab < 0.01).all():
            #     errors.append("tabarru + ujrah ≠ retro nett premium")

    # =========================
    # 12. KOB Code, Pay Period Type, COB
    # =========================
    if reins_type == "INWARD":
        if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
            # Ced EM Rate
            if "ced em rate" not in df.columns:
                errors.append("Tambahkan kolom Ced EM Rate")
            else:
                em_rate = (df["ced em rate"].fillna("").astype(str).str.strip())
                if (em_rate == "").any():
                    errors.append("Ced EM Rate tidak boleh kosong")

            # Ced ER Rate
            if "ced er rate" not in df.columns:
                errors.append("Tambahkan kolom Ced ER Rate")
            else:
                er_rate = (df["ced er rate"].fillna("").astype(str).str.strip())
                if (er_rate == "").any():
                    errors.append("Ced ER Rate tidak boleh kosong")


            # KOB Code
            if "k.o.b code" not in df.columns:
                errors.append("Tambahkan kolom K.O.B Code")
            else:
                kob_series = (df["k.o.b code"].fillna("").astype(str).str.strip())
                if (kob_series == "").any():
                    errors.append("K.O.B Code tidak boleh kosong")
                allowed_kob = {"TTY", "FAC"}
                invalid_kob = set(kob_series[kob_series != ""].unique()) - allowed_kob
                if invalid_kob:
                    errors.append(f"K.O.B Code harus bernilai salah satu dari: "f"{', '.join(sorted(allowed_kob))}")

            # Pay Period Type
            if "pay period type" not in df.columns:
                errors.append("Tambahkan kolom Pay Period Type")
            else:
                pay_period_series = (df["pay period type"].fillna("").astype(str).str.strip())
                if (pay_period_series == "").any():
                    errors.append("Pay Period Type tidak boleh kosong")
                allowed_pay_period = {"Monthly", "Quarterly", "Half Yearly", "Yearly", "Single Premium"}
                invalid_pay_period = set(pay_period_series[pay_period_series != ""].unique()) - allowed_pay_period
                if invalid_pay_period:
                    errors.append(f"Pay Period Type harus bernilai salah satu dari: "f"{', '.join(sorted(allowed_pay_period))}")

            # COB
            if "cob" not in df.columns:
                errors.append("Tambahkan kolom COB")
            else:
                cob_series = (df["cob"].fillna("").astype(str).str.strip())
                if (cob_series == "").any():
                    errors.append("COB tidak boleh kosong")
                allowed_cob = {"CREDIT GROUP", "HEALTH GROUP", "HEALTH INDIVIDUAL", "LIFE GROUP", "LIFE INDIVIDUAL", "P.A GROUP", "P.A INDIVIDUAL"}
                invalid_cob = set(cob_series[cob_series != ""].unique()) - allowed_cob
                if invalid_cob:
                    errors.append(f"COB harus bernilai salah satu dari: "f"{', '.join(sorted(allowed_cob))}")

            # CBY
            if "cby" not in df.columns:
                errors.append("Tambahkan kolom CBY")
            else:
                cby_series = (df["cby"].fillna("").astype(str).str.strip())
                if (cby_series == "").any():
                    errors.append("CBY tidak boleh kosong")

            # CBM
            if "cbm" not in df.columns:
                errors.append("Tambahkan kolom CBM")
            else:
                cbm_series = (df["cbm"].fillna("").astype(str).str.strip())
                if (cbm_series == "").any():
                    errors.append("CBM tidak boleh kosong")


        elif biz_type == "Claim":
            # COB
            if "classofbusiness" not in df.columns:
                errors.append("Tambahkan kolom ClassofBusiness")
            else:
                cob_series = (df["classofbusiness"].fillna("").astype(str).str.strip())
                if (cob_series == "").any():
                    errors.append("ClassOfBusiness tidak boleh kosong")
                allowed_cob = {"CREDIT GROUP", "HEALTH GROUP", "HEALTH INDIVIDUAL", "LIFE GROUP", "LIFE INDIVIDUAL", "P.A GROUP", "P.A INDIVIDUAL"}
                invalid_cob = set(cob_series[cob_series != ""].unique()) - allowed_cob
                if invalid_cob:
                    errors.append(f"ClassOfBusiness harus bernilai salah satu dari: "f"{', '.join(sorted(allowed_cob))}")

            # Pay Period Type
            if "payperiodtype" not in df.columns:
                errors.append("Tambahkan kolom PayPeriodType")
            else:
                pay_period_series = (df["payperiodtype"].fillna("").astype(str).str.strip())
                if (pay_period_series == "").any():
                    errors.append("PayPeriodType tidak boleh kosong")
                allowed_pay_period = {"Monthly", "Quarterly", "Half Yearly", "Yearly", "Single Premium"}
                invalid_pay_period = set(pay_period_series[pay_period_series != ""].unique()) - allowed_pay_period
                if invalid_pay_period:
                    errors.append(f"PayPeriodType harus bernilai salah satu dari: "f"{', '.join(sorted(allowed_pay_period))}")

            # KOB Code
            if "kindofbusiness" not in df.columns:
                errors.append("Tambahkan kolom KindOfBusiness")
            else:
                kob_series = (df["kindofbusiness"].fillna("").astype(str).str.strip())
                if (kob_series == "").any():
                    errors.append("KindOfBusiness tidak boleh kosong")
                allowed_kob = {"TTY", "FAC"}
                invalid_kob = set(kob_series[kob_series != ""].unique()) - allowed_kob
                if invalid_kob:
                    errors.append(f"KindOfBusiness harus bernilai salah satu dari: "f"{', '.join(sorted(allowed_kob))}")

            # CedBookYear
            if "cedbookyear" not in df.columns:
                errors.append("Tambahkan kolom CedBookYear")
            else:
                cby_series = (df["cedbookyear"].fillna("").astype(str).str.strip())
                if (cby_series == "").any():
                    errors.append("CedBookYear tidak boleh kosong")

            # CedBookMonth
            if "cedbookmonth" not in df.columns:
                errors.append("Tambahkan kolom CedBookMonth")
            else:
                cbm_series = (df["cedbookmonth"].fillna("").astype(str).str.strip())
                if (cbm_series == "").any():
                    errors.append("CedBookMonth tidak boleh kosong")


    if reins_type == "OUTWARD":
        if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
            # Retro Type
            if "retro type" not in df.columns:
                errors.append("Tambahkan kolom Retro Type")
            else:
                retro_type_series = (df["retro type"].fillna("").astype(str).str.strip())
                if (retro_type_series == "").any():
                    errors.append("Retro Type tidak boleh kosong")
                allowed_retro = {"Sp Program", "Sp Arrangement", "Panel"}
                invalid_retro = set(retro_type_series[retro_type_series != ""].unique()) - allowed_retro
                if invalid_retro:
                    errors.append(f"Retro Type harus bernilai salah satu dari: "f"{', '.join(sorted(allowed_retro))}")

            # KOB Code
            if "kob code" not in df.columns:
                errors.append("Tambahkan kolom KOB Code")
            else:
                kob_series = (df["kob code"].fillna("").astype(str).str.strip())
                if (kob_series == "").any():
                    errors.append("KOB Code tidak boleh kosong")
                allowed_kob = {"TTY", "FAC"}
                invalid_kob = set(kob_series[kob_series != ""].unique()) - allowed_kob
                if invalid_kob:
                    errors.append(f"KOB Code harus bernilai salah satu dari: "f"{', '.join(sorted(allowed_kob))}")

            # Ced Book Year
            cby_series = (df["ced book year"].fillna("").astype(str).str.strip())
            if (cby_series == "").any():
                errors.append("Ced Book Year tidak boleh kosong")

            # Ced Book Month
            cbm_series = (df["ced book month"].fillna("").astype(str).str.strip())
            if (cbm_series == "").any():
                errors.append("Ced Book Month tidak boleh kosong")

            # Inw Pay Period Type
            inw_pay_period_series = (df["inw pay period type"].fillna("").astype(str).str.strip())
            if (inw_pay_period_series == "").any():
                errors.append("Inw Pay Period Type tidak boleh kosong")
            allowed_inw_pay_period = {"Monthly", "Quarterly", "Half Yearly", "Yearly", "Single Premium"}
            invalid_inw_pay_period = set(inw_pay_period_series[inw_pay_period_series != ""].unique()) - allowed_inw_pay_period
            if invalid_inw_pay_period:
                errors.append(f"Inw Pay Period Type harus bernilai salah satu dari: "f"{', '.join(sorted(allowed_inw_pay_period))}")
            
            # Out Pay Period Type
            out_pay_period_series = (df["out pay period type"].fillna("").astype(str).str.strip())
            if (out_pay_period_series == "").any():
                errors.append("Out Pay Period Type tidak boleh kosong")
            allowed_out_pay_period = {"Monthly", "Quarterly", "Half Yearly", "Yearly"}
            invalid_out_pay_period = set(out_pay_period_series[out_pay_period_series != ""].unique()) - allowed_out_pay_period
            if invalid_out_pay_period:
                errors.append(f"Out Pay Period Type harus bernilai salah satu dari: "f"{', '.join(sorted(allowed_out_pay_period))}")

            # COB
            if "cob" not in df.columns:
                errors.append("Tambahkan kolom COB")
            else:
                cob_series = (df["cob"].fillna("").astype(str).str.strip())
                if (cob_series == "").any():
                    errors.append("COB tidak boleh kosong")
                allowed_cob = {"CREDIT GROUP", "HEALTH GROUP", "HEALTH INDIVIDUAL", "LIFE GROUP", "LIFE INDIVIDUAL", "P.A GROUP", "P.A INDIVIDUAL"}
                invalid_cob = set(cob_series[cob_series != ""].unique()) - allowed_cob
                if invalid_cob:
                    errors.append(f"COB harus bernilai salah satu dari: "f"{', '.join(sorted(allowed_cob))}")

                         
        elif biz_type == "Claim":
            # Retro Type
            if "retro type" not in df.columns:
                errors.append("Tambahkan kolom Retro Type")
            else:
                retro_type_series = (df["retro type"].fillna("").astype(str).str.strip())
                if (retro_type_series == "").any():
                    errors.append("Retro Type tidak boleh kosong")
                allowed_retro = {"Sp Program", "Sp Arrangement", "Panel"}
                invalid_retro = set(retro_type_series[retro_type_series != ""].unique()) - allowed_retro
                if invalid_retro:
                    errors.append(f"Retro Type harus bernilai salah satu dari: "f"{', '.join(sorted(allowed_retro))}")

            # COB
            if "cob" not in df.columns:
                errors.append("Tambahkan kolom COB")
            else:
                cob_series = (df["cob detail"].fillna("").astype(str).str.strip())
                if (cob_series == "").any():
                    errors.append("COB Detail tidak boleh kosong")
                allowed_cob = {"CREDIT GROUP", "HEALTH GROUP", "HEALTH INDIVIDUAL", "LIFE GROUP", "LIFE INDIVIDUAL", "P.A GROUP", "P.A INDIVIDUAL"}
                invalid_cob = set(cob_series[cob_series != ""].unique()) - allowed_cob
                if invalid_cob:
                    errors.append(f"COB Detail harus bernilai salah satu dari: "f"{', '.join(sorted(allowed_cob))}")

            # KOB Code
            if "kob code" not in df.columns:
                errors.append("Tambahkan kolom KOB Code")
            else:
                kob_series = (df["kob code"].fillna("").astype(str).str.strip())
                if (kob_series == "").any():
                    errors.append("KOB Code tidak boleh kosong")
                allowed_kob = {"TTY", "FAC"}
                invalid_kob = set(kob_series[kob_series != ""].unique()) - allowed_kob
                if invalid_kob:
                    errors.append(f"KOB Code harus bernilai salah satu dari: "f"{', '.join(sorted(allowed_kob))}")

            # Ced Book Year
            cby_series = (df["ced book year"].fillna("").astype(str).str.strip())
            if (cby_series == "").any():
                errors.append("Ced Book Year tidak boleh kosong")

            # Ced Book Month
            cbm_series = (df["ced book month"].fillna("").astype(str).str.strip())
            if (cbm_series == "").any():
                errors.append("Ced Book Month tidak boleh kosong")

            # Method of Payment
            method_of_payment_series = (df["method of payment"].fillna("").astype(str).str.strip())
            if (method_of_payment_series == "").any():
                errors.append("Method of Payment tidak boleh kosong")
            allowed_method_of_payment = {"Monthly", "Quarterly", "Half Yearly", "Yearly"}
            invalid_method_of_payment = set(method_of_payment_series[method_of_payment_series != ""].unique()) - allowed_method_of_payment
            if invalid_method_of_payment:
                errors.append(f"Method of Payment harus bernilai salah satu dari: "f"{', '.join(sorted(allowed_method_of_payment))}")

    return errors


def validate_calculate(df, biz_type: str, reins_type: str):

    if reins_type == "INWARD":
        if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
            errors = []

            for col in ["K.O.B Code", "Ccy Code", "Pay Period Type", "CBY", "CBM", "COB", "References No"]:

                if col not in df.columns:
                    errors.append(f"Kolom {col} tidak ditemukan")
                    continue

                series = (
                    df[col]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                )

                # cek kosong
                if (series == "").any():
                    errors.append(f"Kolom {col} terdapat data kosong")
                    continue

                unique = series.unique()

                if len(unique) > 1:
                    errors.append(col)

        if biz_type == "Claim":
            errors = []

            for col in ["CedBookYear", "CedBookMonth", "ClassOfBusiness", "PayPeriodType", "KindOfBusiness", "Currency", "References No"]:

                if col not in df.columns:
                    errors.append(f"Kolom {col} tidak ditemukan")
                    continue

                series = (
                    df[col]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                )

                # cek kosong
                if (series == "").any():
                    errors.append(f"Kolom {col} terdapat data kosong")
                    continue

                unique = series.unique()

                if len(unique) > 1:
                    errors.append(col)

    elif reins_type == "OUTWARD":
        if biz_type in ["Kontribusi", "Refund", "Alteration", "Retur", "Revise", "Batal", "Cancel"]:
            errors = []

            for col in ["Retro Type", "Acc With Name", "KOB Code", "Ccy Code", "Premium Ccy", "Ced Book Year", "Ced Book Month", "Out Pay Period Type", "COB", "References No"]:

                if col not in df.columns:
                    errors.append(f"Kolom {col} tidak ditemukan")
                    continue

                series = (
                    df[col]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                )

                # cek kosong
                if (series == "").any():
                    errors.append(f"Kolom {col} terdapat data kosong")
                    continue

                unique = series.unique()

                if len(unique) > 1:
                    errors.append(col)

        if biz_type == "Claim":
            errors = []

            for col in ["Retro Type", "Cedant Name", "COB Detail", "KOB Code", "Ced Book Year", "Ced Book Month", "Method of Payment", "Curr", "Reinsurer Name", "Voucher Desc"]:

                if col not in df.columns:
                    errors.append(f"Kolom {col} tidak ditemukan")
                    continue

                series = (
                    df[col]
                    .fillna("")
                    .astype(str)
                    .str.strip()
                )

                # cek kosong
                if (series == "").any():
                    errors.append(f"Kolom {col} terdapat data kosong")
                    continue

                unique = series.unique()

                if len(unique) > 1:
                    errors.append(col)
    
    return errors

