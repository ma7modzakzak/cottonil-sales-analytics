import pandas as pd
import numpy as np

# 1. تحديد مسار الملف وقراءته
file_path = r"D:\data analysis\round final project\Cottonil Excel Files\الاساسيات\DISCOUNTS.xlsx"
df = pd.read_excel(file_path)

# 2. تنظيف أسماء الأعمدة الأساسية
df.columns = [str(col).strip() for col in df.columns]

# 3. 🧹 إزالة أسطر العناوين/الهيدر المكررة داخل الصفوف
header_mask = df.astype(str).apply(lambda row: row.str.contains('رقم الإذن|رقم اﻹذن|الصنف', na=False).any(), axis=1)
df = df[~header_mask].reset_index(drop=True)

# 4. دمج العمود الثاني والعمود الثالث (إن لم يكونا من أعمدة الأرقام)
col_2_name = df.columns[1]
col_3_name = df.columns[2]

if col_2_name not in ['السعر', 'الكمية', 'خصم', 'الخصم'] and col_3_name not in ['السعر', 'الكمية', 'خصم', 'الخصم']:
    merged_col_name = f"{col_2_name}_{col_3_name}"
    df[merged_col_name] = df[col_2_name].fillna('').astype(str).str.strip() + ' ' + df[col_3_name].fillna('').astype(str).str.strip()
    df = df.drop(columns=[col_2_name, col_3_name])

# 5. توحيد مسمى عمود الخصم
for col in df.columns:
    if 'خصم' in col:
        df.rename(columns={col: 'الخصم'}, inplace=True)

# 6. 🔢 تحويل أعمدة (السعر، الكمية، الخصم) لأرقام + تحويل القيم السالبة إلى موجبة (.abs())
num_cols = ['السعر', 'الكمية', 'الخصم']
for col in num_cols:
    if col in df.columns:
        df[col] = df[col].astype(str).str.strip().str.replace(',', '')
        # تحويل لأرقام ثم تحويل القيمة المطلقة (Absolute) لضمان القيم الموجبة
        df[col] = pd.to_numeric(df[col], errors='coerce').abs()
    else:
        df[col] = 0.0

# 7. الشرط: إذا كان السعر فارغاً (NaN) أو 0 مع وجود قيمة في الكمية
mask_price_empty = (df['السعر'].isna() | (df['السعر'] == 0)) & (df['الكمية'].notna() & (df['الكمية'] != 0))

# نقل القيمة من عمود الكمية إلى عمود السعر
df.loc[mask_price_empty, 'السعر'] = df.loc[mask_price_empty, 'الكمية']

# تصفير خلية الكمية للصفوف التي تم نقل قيمتها
df.loc[mask_price_empty, 'الكمية'] = 0.0

# تعبئة بقية القيم الفارغة بصفر
df['السعر'] = df['السعر'].fillna(0.0)
df['الكمية'] = df['الكمية'].fillna(0.0)
df['الخصم'] = df['الخصم'].fillna(0.0)

# 8. حساب عمود "الصافي" وتحويل ناتجه إلى قيمة موجبة صريحة
df['الصافي'] = ((df['السعر'] * df['الكمية']) - (df['الخصم'] * df['الكمية'])).abs()

# 9. حفظ الملف النهائي
output_filename = 'cleaned_DISCOUNTS_final.xlsx'
try:
    df.to_excel(output_filename, index=False)
    print(f"تم تنظيف الملف وتأكيد أن جميع القيم موجبة وحساب الصافي بنجاح في: {output_filename}")
except PermissionError:
    alt_filename = 'cleaned_DISCOUNTS_final_v2.xlsx'
    df.to_excel(alt_filename, index=False)
    print(f"تنبيه: الملف كان مفتوحاً، تم الحفظ بملف جديد باسم: {alt_filename}")

df.head(10)