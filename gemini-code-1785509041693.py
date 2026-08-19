import pandas as pd
import re
import numpy as np

# Load Data
file_path = "Planilha base Porsche.xlsx"
df = pd.read_excel(file_path)

# --- 1. Dates ---
# Rule update: Specifically handle formats like 18/04/2024 and 2024/15/07 (which implies yyyy/dd/mm)
def sanitize_date(val):
    val = str(val).strip()
    val = re.sub(r'(\d+)(st|nd|rd|th)', r'\1', val)
    
    # Try multiple parsing strategies
    # if it matches YYYY/DD/MM
    m = re.match(r'^(\d{4})[/-](\d{2})[/-](\d{2})$', val)
    if m:
        p1, p2, p3 = int(m.group(1)), int(m.group(2)), int(m.group(3))
        # if p2 > 12, it must be DD and p3 must be MM
        if p2 > 12 and p3 <= 12:
            val = f"{p1}-{p3:02d}-{p2:02d}"
    
    # if it matches DD/MM/YYYY
    m2 = re.match(r'^(\d{2})[/-](\d{2})[/-](\d{4})$', val)
    if m2:
        p1, p2, p3 = int(m2.group(1)), int(m2.group(2)), int(m2.group(3))
        # if p1 > 12, it must be DD/MM/YYYY
        if p1 > 12 and p2 <= 12:
            val = f"{p3}-{p2:02d}-{p1:02d}"

    try:
        dt = pd.to_datetime(val, errors='raise')
        return dt.strftime('%Y-%m-%d')
    except:
        return 'INVALID'

df['SaleDateSanitized'] = df['sale_date'].apply(sanitize_date)

# --- 2. Porsche Models ---
models_list = [
    "911 Carrera", "911 Carrera S", "911 Carrera GTS", "911 Turbo", "911 Turbo S",
    "911 GT3", "911 GT3 RS", "911 Dakar", "911 Targa 4", "911 Targa 4S",
    "718 Cayman", "718 Cayman S", "718 Cayman GT4 RS", "718 Boxster", "718 Boxster GTS",
    "718 Spyder RS", "Cayenne", "Cayenne S", "Cayenne Coupe", "Cayenne E-Hybrid",
    "Cayenne Turbo", "Cayenne Turbo GT", "Macan", "Macan S", "Macan T", "Macan GTS",
    "Macan Electric", "Panamera", "Panamera 4", "Panamera 4S", "Panamera Turbo",
    "Panamera Turbo S", "Panamera 4 E-Hybrid", "Taycan", "Taycan 4S", "Taycan GTS",
    "Taycan Turbo", "Taycan Turbo S", "Taycan Cross Turismo"
]
model_map = {m.lower(): m for m in models_list}

def sanitize_model(val):
    val = str(val).strip()
    lower_val = val.lower()
    if lower_val in model_map:
        return model_map[lower_val]
    return val.title()

df['PorscheModelSanitized'] = df['porsche_model'].apply(sanitize_model)

# --- 3. Model Year ---
# Added 2025 and 2026
word_to_num = {
    'twenty twenty six': 2026, 'two thousand twenty six': 2026,
    'twenty twenty five': 2025, 'two thousand twenty five': 2025,
    'twenty twenty four': 2024, 'two thousand twenty four': 2024,
    'twenty twenty three': 2023, 'two thousand twenty three': 2023,
    'twenty twenty two': 2022, 'two thousand twenty two': 2022,
    'twenty twenty one': 2021, 'two thousand twenty one': 2021,
    'twenty twenty': 2020, 'two thousand twenty': 2020,
}
def sanitize_year(val):
    val = str(val).strip().lower()
    if val in word_to_num:
        val = str(word_to_num[val])
    val = re.sub(r'^20\s*[-_\.]?\s*(\d{2})$', r'20\1', val)
    try:
        yr = int(val)
        if 1990 <= yr <= 2035:
            return str(yr)
    except:
        pass
    return 'INVALID'

df['ModelYearSanitized'] = df['model_year'].apply(sanitize_year)

# --- 4. Sales Price ---
price_words = {
    'eighty two thousand': 82000,
    'two hundred thousand': 200000,
}
def sanitize_price(val):
    val = str(val).strip().lower()
    for k, v in price_words.items():
        if k in val:
            return f"{float(v):.2f}"
    
    is_k = 'k' in val
    clean_chars = re.sub(r'[^0-9\.,]', '', val)
    if clean_chars.endswith(',') or clean_chars.endswith('.'):
        clean_chars = clean_chars[:-1]
    
    if ',' in clean_chars and '.' in clean_chars:
        if clean_chars.rfind(',') > clean_chars.rfind('.'):
            clean_chars = clean_chars.replace('.', '').replace(',', '.')
        else:
            clean_chars = clean_chars.replace(',', '')
    elif ',' in clean_chars:
        clean_chars = clean_chars.replace(',', '')
    elif '.' in clean_chars:
        parts = clean_chars.split('.')
        if len(parts) > 1 and len(parts[-1]) == 3 and not is_k:
            clean_chars = clean_chars.replace('.', '')
    
    try:
        num = float(clean_chars)
        if is_k:
            num = num * 1000
        return f"{num:.2f}"
    except:
        return 'INVALID'

df['SalesPriceSanitized'] = df['sale_price'].apply(sanitize_price)

# --- 5. Vehicle Mileage ---
mile_words = {
    'twelve thousand': 12000,
    'zero': 0, 'new': 0
}
def sanitize_mileage(val):
    val_str = str(val).lower().strip()
    if 'zero' in val_str or 'new' in val_str:
        return '0'
    for k, v in mile_words.items():
        if k in val_str:
            return str(v)
            
    is_km = 'km' in val_str
    clean_chars = re.sub(r'[^0-9\.,]', '', val_str)
    if ',' in clean_chars and '.' in clean_chars:
        if clean_chars.rfind(',') > clean_chars.rfind('.'):
            clean_chars = clean_chars.replace('.', '').replace(',', '.')
        else:
            clean_chars = clean_chars.replace(',', '')
    elif ',' in clean_chars:
        clean_chars = clean_chars.replace(',', '')
    elif '.' in clean_chars:
        parts = clean_chars.split('.')
        if len(parts[-1]) == 3:
            clean_chars = clean_chars.replace('.', '')
            
    try:
        num = float(clean_chars)
        if is_km:
            num = num * 0.621371
        return str(int(round(num)))
    except:
        return 'INVALID'

df['VehicleMileageSanitized'] = df['vehicle_mileage'].apply(sanitize_mileage)

# --- 6. Payment Method ---
# Add Leasing -> Lease
def sanitize_pay(val):
    val = str(val).lower()
    if 'credit' in val: return 'Credit Card'
    if 'debit' in val: return 'Debit Card'
    if 'bank' in val: return 'Bank Transfer'
    if 'wire' in val: return 'Wire Transfer'
    if 'financ' in val: return 'Financing'
    if 'leas' in val: return 'Lease'
    if 'cash' in val: return 'Cash'
    if 'ach' in val: return 'ACH Payment'
    if 'crypto' in val: return 'Crypto Payment'
    return val.title()

df['PayMethodSanitized'] = df['payment_method'].apply(sanitize_pay)

# --- 7. City ---
df['CitySanitized'] = df['city'].astype(str).str.title()

# --- 8. State ---
us_states = {
    'alabama': 'AL', 'alaska': 'AK', 'arizona': 'AZ', 'arkansas': 'AR',
    'california': 'CA', 'colorado': 'CO', 'connecticut': 'CT', 'delaware': 'DE',
    'florida': 'FL', 'georgia': 'GA', 'hawaii': 'HI', 'idaho': 'ID',
    'illinois': 'IL', 'indiana': 'IN', 'iowa': 'IA', 'kansas': 'KS',
    'kentucky': 'KY', 'louisiana': 'LA', 'maine': 'ME', 'maryland': 'MD',
    'massachusetts': 'MA', 'michigan': 'MI', 'minnesota': 'MN', 'mississippi': 'MS',
    'missouri': 'MO', 'montana': 'MT', 'nebraska': 'NE', 'nevada': 'NV',
    'new hampshire': 'NH', 'new jersey': 'NJ', 'new mexico': 'NM', 'new york': 'NY',
    'north carolina': 'NC', 'north dakota': 'ND', 'ohio': 'OH', 'oklahoma': 'OK',
    'oregon': 'OR', 'pennsylvania': 'PA', 'rhode island': 'RI', 'south carolina': 'SC',
    'south dakota': 'SD', 'tennessee': 'TN', 'texas': 'TX', 'utah': 'UT',
    'vermont': 'VT', 'virginia': 'VA', 'washington': 'WA', 'west virginia': 'WV',
    'wisconsin': 'WI', 'wyoming': 'WY', 'dc': 'DC', 'district of columbia': 'DC'
}
valid_abbr = set(us_states.values())

def sanitize_state(val):
    val = str(val).strip().lower()
    if val in us_states:
        return us_states[val]
    val_upper = val.upper()
    if val_upper in valid_abbr:
        return val_upper
    return 'INVALID'

df['StateSanitized'] = df['state'].apply(sanitize_state)

# --- 9. Delivery Status ---
deliv_methods = [
    'Delivered', 'Pending', 'In Transit', 'Cancelled', 'Awaiting Delivery',
    'Awaiting Pickup', 'Pending Approval', 'Pending Review', 'Shipped', 'Awaiting Review'
]
def sanitize_delivery(val):
    val = str(val).lower().replace('-', ' ').replace('_', ' ').strip()
    val = re.sub(r'[^a-z ]', '', val)
    if 'deliver' in val and 'await' not in val: return 'Delivered'
    if 'transit' in val: return 'In Transit'
    if val == 'pending': return 'Pending'
    for dm in deliv_methods:
        if dm.lower() == val:
            return dm
    return val.title()

df['DeliveryStatusSanitized'] = df['delivery_status'].apply(sanitize_delivery)

# Reorder columns
new_cols = []
mapping = {
    'sale_date': 'SaleDateSanitized',
    'porsche_model': 'PorscheModelSanitized',
    'model_year': 'ModelYearSanitized',
    'sale_price': 'SalesPriceSanitized',
    'vehicle_mileage': 'VehicleMileageSanitized',
    'payment_method': 'PayMethodSanitized',
    'city': 'CitySanitized',
    'state': 'StateSanitized',
    'delivery_status': 'DeliveryStatusSanitized'
}

for col in df.columns:
    if col in mapping.values(): continue
    new_cols.append(col)
    if col in mapping:
        new_cols.append(mapping[col])

df_final = df[new_cols].copy()
df_final.to_excel('Planilha_Porsche_Sanitized_v2.xlsx', index=False)

# Let's generate the raw python script for the user as requested
with open('porsche_sanitization_script.py', 'w') as f:
    f.write('''import pandas as pd
import re
import numpy as np

def process_porsche_data(input_path, output_path):
    df = pd.read_excel(input_path)

    # --- 1. Dates ---
    def sanitize_date(val):
        val = str(val).strip()
        val = re.sub(r'(\d+)(st|nd|rd|th)', r'\\1', val)
        m = re.match(r'^(\d{4})[/-](\d{2})[/-](\d{2})$', val)
        if m:
            p1, p2, p3 = int(m.group(1)), int(m.group(2)), int(m.group(3))
            if p2 > 12 and p3 <= 12:
                val = f"{p1}-{p3:02d}-{p2:02d}"
        
        m2 = re.match(r'^(\d{2})[/-](\d{2})[/-](\d{4})$', val)
        if m2:
            p1, p2, p3 = int(m2.group(1)), int(m2.group(2)), int(m2.group(3))
            if p1 > 12 and p2 <= 12:
                val = f"{p3}-{p2:02d}-{p1:02d}"

        try:
            dt = pd.to_datetime(val, errors='raise')
            return dt.strftime('%Y-%m-%d')
        except:
            return 'INVALID'

    df['SaleDateSanitized'] = df['sale_date'].apply(sanitize_date)

    # --- 2. Porsche Models ---
    models_list = [
        "911 Carrera", "911 Carrera S", "911 Carrera GTS", "911 Turbo", "911 Turbo S",
        "911 GT3", "911 GT3 RS", "911 Dakar", "911 Targa 4", "911 Targa 4S",
        "718 Cayman", "718 Cayman S", "718 Cayman GT4 RS", "718 Boxster", "718 Boxster GTS",
        "718 Spyder RS", "Cayenne", "Cayenne S", "Cayenne Coupe", "Cayenne E-Hybrid",
        "Cayenne Turbo", "Cayenne Turbo GT", "Macan", "Macan S", "Macan T", "Macan GTS",
        "Macan Electric", "Panamera", "Panamera 4", "Panamera 4S", "Panamera Turbo",
        "Panamera Turbo S", "Panamera 4 E-Hybrid", "Taycan", "Taycan 4S", "Taycan GTS",
        "Taycan Turbo", "Taycan Turbo S", "Taycan Cross Turismo"
    ]
    model_map = {m.lower(): m for m in models_list}

    def sanitize_model(val):
        val = str(val).strip()
        lower_val = val.lower()
        if lower_val in model_map:
            return model_map[lower_val]
        return val.title()

    df['PorscheModelSanitized'] = df['porsche_model'].apply(sanitize_model)

    # --- 3. Model Year ---
    word_to_num = {
        'twenty twenty six': 2026, 'two thousand twenty six': 2026,
        'twenty twenty five': 2025, 'two thousand twenty five': 2025,
        'twenty twenty four': 2024, 'two thousand twenty four': 2024,
        'twenty twenty three': 2023, 'two thousand twenty three': 2023,
        'twenty twenty two': 2022, 'two thousand twenty two': 2022,
        'twenty twenty one': 2021, 'two thousand twenty one': 2021,
        'twenty twenty': 2020, 'two thousand twenty': 2020,
    }
    def sanitize_year(val):
        val = str(val).strip().lower()
        if val in word_to_num:
            val = str(word_to_num[val])
        val = re.sub(r'^20\s*[-_\.]?\s*(\d{2})$', r'20\\1', val)
        try:
            yr = int(val)
            if 1990 <= yr <= 2035:
                return str(yr)
        except:
            pass
        return 'INVALID'

    df['ModelYearSanitized'] = df['model_year'].apply(sanitize_year)

    # --- 4. Sales Price ---
    price_words = {
        'eighty two thousand': 82000,
        'two hundred thousand': 200000,
    }
    def sanitize_price(val):
        val = str(val).strip().lower()
        for k, v in price_words.items():
            if k in val:
                return f"{float(v):.2f}"
        
        is_k = 'k' in val
        clean_chars = re.sub(r'[^0-9\.,]', '', val)
        if clean_chars.endswith(',') or clean_chars.endswith('.'):
            clean_chars = clean_chars[:-1]
        
        if ',' in clean_chars and '.' in clean_chars:
            if clean_chars.rfind(',') > clean_chars.rfind('.'):
                clean_chars = clean_chars.replace('.', '').replace(',', '.')
            else:
                clean_chars = clean_chars.replace(',', '')
        elif ',' in clean_chars:
            clean_chars = clean_chars.replace(',', '')
        elif '.' in clean_chars:
            parts = clean_chars.split('.')
            if len(parts) > 1 and len(parts[-1]) == 3 and not is_k:
                clean_chars = clean_chars.replace('.', '')
        
        try:
            num = float(clean_chars)
            if is_k:
                num = num * 1000
            return f"{num:.2f}"
        except:
            return 'INVALID'

    df['SalesPriceSanitized'] = df['sale_price'].apply(sanitize_price)

    # --- 5. Vehicle Mileage ---
    mile_words = {
        'twelve thousand': 12000,
        'zero': 0, 'new': 0
    }
    def sanitize_mileage(val):
        val_str = str(val).lower().strip()
        if 'zero' in val_str or 'new' in val_str:
            return '0'
        for k, v in mile_words.items():
            if k in val_str:
                return str(v)
                
        is_km = 'km' in val_str
        clean_chars = re.sub(r'[^0-9\.,]', '', val_str)
        if ',' in clean_chars and '.' in clean_chars:
            if clean_chars.rfind(',') > clean_chars.rfind('.'):
                clean_chars = clean_chars.replace('.', '').replace(',', '.')
            else:
                clean_chars = clean_chars.replace(',', '')
        elif ',' in clean_chars:
            clean_chars = clean_chars.replace(',', '')
        elif '.' in clean_chars:
            parts = clean_chars.split('.')
            if len(parts[-1]) == 3:
                clean_chars = clean_chars.replace('.', '')
                
        try:
            num = float(clean_chars)
            if is_km:
                num = num * 0.621371
            return str(int(round(num)))
        except:
            return 'INVALID'

    df['VehicleMileageSanitized'] = df['vehicle_mileage'].apply(sanitize_mileage)

    # --- 6. Payment Method ---
    def sanitize_pay(val):
        val = str(val).lower()
        if 'credit' in val: return 'Credit Card'
        if 'debit' in val: return 'Debit Card'
        if 'bank' in val: return 'Bank Transfer'
        if 'wire' in val: return 'Wire Transfer'
        if 'financ' in val: return 'Financing'
        if 'leas' in val: return 'Lease'
        if 'cash' in val: return 'Cash'
        if 'ach' in val: return 'ACH Payment'
        if 'crypto' in val: return 'Crypto Payment'
        return val.title()

    df['PayMethodSanitized'] = df['payment_method'].apply(sanitize_pay)

    # --- 7. City ---
    df['CitySanitized'] = df['city'].astype(str).str.title()

    # --- 8. State ---
    us_states = {
        'alabama': 'AL', 'alaska': 'AK', 'arizona': 'AZ', 'arkansas': 'AR',
        'california': 'CA', 'colorado': 'CO', 'connecticut': 'CT', 'delaware': 'DE',
        'florida': 'FL', 'georgia': 'GA', 'hawaii': 'HI', 'idaho': 'ID',
        'illinois': 'IL', 'indiana': 'IN', 'iowa': 'IA', 'kansas': 'KS',
        'kentucky': 'KY', 'louisiana': 'LA', 'maine': 'ME', 'maryland': 'MD',
        'massachusetts': 'MA', 'michigan': 'MI', 'minnesota': 'MN', 'mississippi': 'MS',
        'missouri': 'MO', 'montana': 'MT', 'nebraska': 'NE', 'nevada': 'NV',
        'new hampshire': 'NH', 'new jersey': 'NJ', 'new mexico': 'NM', 'new york': 'NY',
        'north carolina': 'NC', 'north dakota': 'ND', 'ohio': 'OH', 'oklahoma': 'OK',
        'oregon': 'OR', 'pennsylvania': 'PA', 'rhode island': 'RI', 'south carolina': 'SC',
        'south dakota': 'SD', 'tennessee': 'TN', 'texas': 'TX', 'utah': 'UT',
        'vermont': 'VT', 'virginia': 'VA', 'washington': 'WA', 'west virginia': 'WV',
        'wisconsin': 'WI', 'wyoming': 'WY', 'dc': 'DC', 'district of columbia': 'DC'
    }
    valid_abbr = set(us_states.values())

    def sanitize_state(val):
        val = str(val).strip().lower()
        if val in us_states:
            return us_states[val]
        val_upper = val.upper()
        if val_upper in valid_abbr:
            return val_upper
        return 'INVALID'

    df['StateSanitized'] = df['state'].apply(sanitize_state)

    # --- 9. Delivery Status ---
    deliv_methods = [
        'Delivered', 'Pending', 'In Transit', 'Cancelled', 'Awaiting Delivery',
        'Awaiting Pickup', 'Pending Approval', 'Pending Review', 'Shipped', 'Awaiting Review'
    ]
    def sanitize_delivery(val):
        val = str(val).lower().replace('-', ' ').replace('_', ' ').strip()
        val = re.sub(r'[^a-z ]', '', val)
        if 'deliver' in val and 'await' not in val: return 'Delivered'
        if 'transit' in val: return 'In Transit'
        if val == 'pending': return 'Pending'
        for dm in deliv_methods:
            if dm.lower() == val:
                return dm
        return val.title()

    df['DeliveryStatusSanitized'] = df['delivery_status'].apply(sanitize_delivery)

    # Reorder columns
    new_cols = []
    mapping = {
        'sale_date': 'SaleDateSanitized',
        'porsche_model': 'PorscheModelSanitized',
        'model_year': 'ModelYearSanitized',
        'sale_price': 'SalesPriceSanitized',
        'vehicle_mileage': 'VehicleMileageSanitized',
        'payment_method': 'PayMethodSanitized',
        'city': 'CitySanitized',
        'state': 'StateSanitized',
        'delivery_status': 'DeliveryStatusSanitized'
    }

    for col in df.columns:
        if col in mapping.values(): continue
        new_cols.append(col)
        if col in mapping:
            new_cols.append(mapping[col])

    df_final = df[new_cols].copy()
    df_final.to_excel(output_path, index=False)
    print(f"File saved successfully to {output_path}")

if __name__ == "__main__":
    process_porsche_data('Planilha base Porsche.xlsx', 'Planilha_Porsche_Sanitized_v2.xlsx')
''')