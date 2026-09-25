import pandas as pd
import numpy as np

# 1. قراءة ملف المبيعات
file_path = r"D:\data analysis\round final project\Cottonil Excel Files\الاساسيات\SALES.xlsx"
df = pd.read_excel(file_path)

# تنظيف أسماء الأعمدة
df.columns = [str(col).strip() for col in df.columns]

# 2. 🧹 إزالة أسطر الهيدر المكررة داخل البيانات
header_mask = df.astype(str).apply(lambda row: row.str.contains('رقم الإذن|رقم اﻹذن|الصنف', na=False).any(), axis=1)
df = df[~header_mask].reset_index(drop=True)

# 3. دمج العمود الثاني والعمود الثالث (إذا لم يكونا أرقام أو تاريخ)
col_2_name = df.columns[1]
col_3_name = df.columns[2]

if col_2_name not in ['السعر', 'الكمية', 'خصم', 'الخصم', 'التاريخ'] and col_3_name not in ['السعر', 'الكمية', 'خصم', 'الخصم', 'التاريخ']:
    merged_col_name = f"{col_2_name}_{col_3_name}"
    df[merged_col_name] = df[col_2_name].fillna('').astype(str).str.strip() + ' ' + df[col_3_name].fillna('').astype(str).str.strip()
    df = df.drop(columns=[col_2_name, col_3_name])

# 4. توحيد مسمى الخصم
for col in df.columns:
    if 'خصم' in col:
        df.rename(columns={col: 'الخصم'}, inplace=True)

# 5. 📅 تعديل فكرة التواريخ: تقسيم تواريخ كل شهر بالإنصاف (50% لسنة 2024 و 50% لسنة 2025)
if 'التاريخ' in df.columns:
    # أ) تحويل القيمة لتاريخ
    df['التاريخ'] = pd.to_datetime(df['التاريخ'], dayfirst=True, errors='coerce')
    
    # ب) استبدال التواريخ الفارغة بتاريخ افتراضي لتجنب الأخطاء
    df['التاريخ'] = df['التاريخ'].fillna(pd.Timestamp('2024-01-01'))

    # ج) تجميع البيانات حسب الشهر وتنصيف السنة بين 2024 و 2025
    def split_years_per_month(group):
        n = len(group)
        half = n // 2  # النصف الأول
        
        # إنشاء السلسلة الجديدة للأعوام (النصف الأول 2024 والباقي 2025)
        years = [2024] * half + [2025] * (n - half)
        
        # تطبيق السنوات الجديدة مع الحفاظ على اليوم والشهر
        new_dates = []
        for i, row in enumerate(group.itertuples()):
            orig_date = row.التاريخ
            target_year = years[i]
            try:
                # تغيير السنة فقط مع إبقاء اليوم والشهر
                updated_date = orig_date.replace(year=target_year)
            except ValueError:
                # معالجة حالة سنة الكبيسة (29 فبراير)
                updated_date = orig_date.replace(year=target_year, day=28)
            new_dates.append(updated_date)
            
        group['التاريخ'] = new_dates
        return group

    # تطبيق التقسيم لكل شهر على حدة
    df = df.groupby(df['التاريخ'].dt.month, group_keys=False).apply(split_years_per_month)

    # د) إعادة تنسيق التاريخ إلى (يوم/شهر/سنة)
    df['التاريخ'] = df['التاريخ'].dt.strftime('%d/%m/%Y')

# 6. 🔢 تحويل الأرقام الحسابية إلى قيم موجبة (.abs())
num_cols = ['السعر', 'الكمية', 'الخصم']
for col in num_cols:
    if col in df.columns:
        df[col] = df[col].astype(str).str.strip().str.replace(',', '')
        df[col] = pd.to_numeric(df[col], errors='coerce').abs()
    else:
        df[col] = 0.0

# 7. نقل القيم بين السعر والكمية في حال الشغور
mask_price_empty = (df['السعر'].isna() | (df['السعر'] == 0)) & (df['الكمية'].notna() & (df['الكمية'] != 0))
df.loc[mask_price_empty, 'السعر'] = df.loc[mask_price_empty, 'الكمية']
df.loc[mask_price_empty, 'الكمية'] = 0.0

df['السعر'] = df['السعر'].fillna(0.0)
df['الكمية'] = df['الكمية'].fillna(0.0)
df['الخصم'] = df['الخصم'].fillna(0.0)

# 8. حساب الصافي بمجموع موجب
df['الصافي'] = ((df['السعر'] * df['الكمية']) - (df['الخصم'] * df['الكمية'])).abs()

# 9. 💾 حفظ النتيجة النهائية
output_filename = 'CLEANED SALES.xlsx'
try:
    df.to_excel(output_filename, index=False)
    print(f"تم بنجاح! تم حفظ {len(df)} صفاً وتقسيم تواريخ كل شهر بالنصف بين 2024 و 2025 في: {output_filename}")
except PermissionError:
    alt_filename = 'CLEANED SALES_v2.xlsx'
    df.to_excel(alt_filename, index=False)
    print(f"تنبيه: الملف مفتوح، تم الحفظ بملف جديد: {alt_filename}")

df.head(10)