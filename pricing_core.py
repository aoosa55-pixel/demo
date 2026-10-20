"""สูตรคำนวณราคาลูกค้า — ไฟล์กลางที่ใช้ร่วมกันระหว่าง app.py (8501) และ chat_gateway
(คัดลอกจาก app.py ตรงตัว ไม่ได้แก้ตรรกะ) แก้สูตรที่ไฟล์นี้ที่เดียว"""
import math

import pandas as pd


def parse_pct(val):
    if pd.isna(val) or val is None:
        return 0.0
    val_str = str(val).replace("%", "").strip()
    try:
        return float(val_str) / 100.0
    except ValueError:
        return 0.0


def calculate_custom_price(retail_price: float, category: str, cust_row: pd.Series, supplier_cost: float = 0.0):
    cat = str(category).upper()

    honda_big = parse_pct(cust_row.get("HONDA_BIG_BIKE"))
    honda_small = parse_pct(cust_row.get("HONDA_SMALL_BIKE"))
    yamaha_big = parse_pct(cust_row.get("YAMAHA_BIG_BIKE"))
    yamaha_small = parse_pct(cust_row.get("YAMAHA_SMALL_BIKE"))
    kawasaki_big = parse_pct(cust_row.get("KAWASAKI_BIG_BIKE"))
    kawasaki_small = parse_pct(cust_row.get("KAWASAKI_SMALL_BIKE"))

    net_markup = parse_pct(cust_row.get("NET_PRICE"))
    vat_rate = parse_pct(cust_row.get("VAT"))
    fob_rate = parse_pct(cust_row.get("FOB"))

    if "NET" in cat:
        base_cost = supplier_cost if supplier_cost > 0 else retail_price
        base_price = base_cost * (1.0 + net_markup)
        discount_label = f"NET (+{int(net_markup * 100)}%)"
    elif "HONDA" in cat:
        is_bigbike = "BIGBIKE" in cat or "BIG BIKE" in cat
        disc = honda_big if is_bigbike else honda_small
        base_price = retail_price * (1.0 - disc)
        discount_label = f"HONDA {'BIG' if is_bigbike else 'SMALL'} ({int(disc * 100)}%)"
    elif "YAMAHA" in cat:
        is_bigbike = "BIGBIKE" in cat or "BIG BIKE" in cat
        disc = yamaha_big if is_bigbike else yamaha_small
        base_price = retail_price * (1.0 - disc)
        discount_label = f"YAMAHA {'BIG' if is_bigbike else 'SMALL'} ({int(disc * 100)}%)"
    elif "KAWASAKI" in cat:
        is_bigbike = "BIGBIKE" in cat or "BIG BIKE" in cat
        disc = kawasaki_big if is_bigbike else kawasaki_small
        base_price = retail_price * (1.0 - disc)
        discount_label = f"KAWASAKI {'BIG' if is_bigbike else 'SMALL'} ({int(disc * 100)}%)"
    else:
        base_price = retail_price
        discount_label = "NET (0%)"

    final_raw_price = base_price * (1.0 + vat_rate + fob_rate)
    final_unit_price = int(math.ceil(round(final_raw_price, 4)))

    return final_unit_price, discount_label
