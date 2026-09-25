import pandas as pd
import numpy as np

# 1. قراءة البيانات
file_path = r"D:\data analysis\round final project\Cottonil Excel Files\الاساسيات\المشتريات.xlsx"
df = pd.read_excel(file_path, header=None)

# 2. إنشاء الأعمدة الجديدة وتأكيد نوع البيانات كـ object
new_cols = ['رقم الإذن', 'التاريخ', 'المورد', 'أمين المخزن', 'المخزن', 'المستلم']
for col in new_cols:
    df[col] = np.nan
    df[col] = df[col].astype('object')

# قائمة لتجميع أي أسماء/قيم مستخرجة لمسحها لاحقاً
extracted_values = set()

# 3. دالة استخراج وتوزيع القيم المباشرة عبر فحص شامل للأسطر
for idx, row in df.iterrows():
    row_cells = [str(val).strip() for val in row.values if pd.notna(val) and str(val).strip() not in ['nan', 'None', '']]
    row_text = " ".join(row_cells)
    
    # استخراج رقم الإذن، التاريخ، والمورد عند وجود الكلمة المفتاحية
    if 'رقم إذن' in row_text or 'إذن' in row_text:
        for i, cell in enumerate(row_cells):
            if ('رقم إذن' in cell or 'إذن' in cell) and i + 1 < len(row_cells):
                val = str(row_cells[i + 1])
                df.at[idx, 'رقم الإذن'] = val
                extracted_values.add(val)
            elif ('التاريخ' in cell or 'التاري' in cell) and i + 1 < len(row_cells):
                val = str(row_cells[i + 1])
                df.at[idx, 'التاريخ'] = val
                extracted_values.add(val)
            elif 'المورد' in cell and i + 1 < len(row_cells):
                val = str(row_cells[i + 1])
                df.at[idx, 'المورد'] = val
                extracted_values.add(val)

        # اختيار آخر قيمة كمورد إن لم تُكتشف بجوار الكلمة مباشرة
        if pd.isna(df.at[idx, 'المورد']):
            for cell in reversed(row_cells):
                if not cell.isdigit() and '/' not in cell and cell not in ['رقم إذن', 'التاريخ', 'المورد']:
                    val = str(cell)
                    df.at[idx, 'المورد'] = val
                    extracted_values.add(val)
                    break

    # استخراج (أمين المخزن، المخزن، المستلم)
    for cell in row_cells:
        if 'أحمد' in cell or 'احمد' in cell:
            val = str(cell)
            df.at[idx, 'أمين المخزن'] = val
            extracted_values.add(val)
        if 'المخزن الرئيسي' in cell or 'مخزن' in cell:
            val = str(cell)
            df.at[idx, 'المخزن'] = val
            extracted_values.add(val)
        if 'أمين' in cell or 'امين' in cell:
            val = str(cell)
            df.at[idx, 'المستلم'] = val
            extracted_values.add(val)

# 4. التعبئة لأسفل (ffill) على كامل DataFrame
for col in new_cols:
    df[col] = df[col].ffill()

# 5. جلب عناوين الأعمدة الحقيقية للأصناف
header_mask = df.astype(str).apply(lambda row: row.str.contains('بارك|الباركوود|الباركود|الصنف', na=False).any(), axis=1)
header_idx = df[header_mask].index[0]
headers = df.iloc[header_idx, :13].values.tolist()

# 6. استبعاد أسطر الهيدر للحصول على أسطر الأصناف فقط
is_header_row = df.astype(str).apply(lambda row: row.str.contains('رقم إذن|الباركوود|الباركود|التاريخ', na=False).any(), axis=1)
clean_df = df[~is_header_row & df.iloc[:, 1:12].notna().any(axis=1)].copy()

# 7. اقتطاع بيانات الأصناف المكونة من 13 عموداً
product_data = clean_df.iloc[:, :13].copy()

# 8. 🧹 مسح الأسماء المكررة واستبعاد الفراغات تماماً من كل عمود بالموقع (iloc)
keywords_to_clean = set(['رقم إذن', 'إذن', 'التاريخ', 'المورد', 'أمين المخزن', 'المخزن', 'المستلم']).union(extracted_values)

cleaned_columns = {}
for i in range(13):
    # الوصول لكل عمود بالموقع لتفادي مشكلة تكرار أسماء الأعمدة
    col_series = product_data.iloc[:, i].dropna()
    
    # تصفية القيم المكررة والنصوص غير المرغوب فيها
    mask = ~col_series.astype(str).str.strip().isin(keywords_to_clean)
    valid_series = col_series[mask].reset_index(drop=True)
    
    # تسمية العمود باسم الهيدر المناسب
    col_name = str(headers[i]) if i < len(headers) else f"Column_{i}"
    cleaned_columns[col_name] = valid_series

# إعادة بناء جدول الأصناف الخالي تماماً من الفراغات والـ null
cleaned_product_df = pd.DataFrame(cleaned_columns)

# 9. دمج الأعمدة التعريفية الجديدة وإعادة الترتيب النهائي
for col in new_cols:
    cleaned_product_df[col] = clean_df[col].reset_index(drop=True)

# ترتيب الأعمدة النهائي
final_cols = new_cols + [str(h) for h in headers]
final_df = cleaned_product_df.reindex(columns=final_cols)

# إزالة أي أسطر فارغة بالكامل إن وجدت
final_df = final_df.dropna(subset=[str(h) for h in headers], how='all')

# 10. حفظ الملف النهائي
output_filename = 'cleaned_purchases_final.xlsx'
try:
    final_df.to_excel(output_filename, index=False)
    print(f"تم مسح القيم المكررة وضغط البيانات بدون أي خلايا null بنجاح في {output_filename}")
except PermissionError:
    alt_filename = 'cleaned_purchases_final_v2.xlsx'
    final_df.to_excel(alt_filename, index=False)
    print(f"تنبيه: الملف الأساسي كان مفتوحاً، تم الحفظ بملف جديد باسم: {alt_filename}")

final_df.head(10)