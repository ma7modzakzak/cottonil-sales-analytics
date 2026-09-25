import pandas as pd
import numpy as np
import re

# 1. مسار الملف المباشر
file_path = r"D:\data analysis\round final project\Cottonil Excel Files\الاساسيات\Excel جرد البضاعة.xlsx"

# قراءة الملف بدون عناوين افتراضية
df = pd.read_excel(file_path, header=None)

parsed_rows = []

for idx, row in df.iterrows():
    # استخراج الخلايا المليئة فقط
    row_cells = [val for val in row.values if pd.notna(val) and str(val).strip() not in ['nan', 'None', '']]
    str_cells = [str(val).strip() for val in row_cells]
    row_text = " ".join(str_cells)
    
    # تجنب أسطر الهيدر وعناوين التقرير والتواريخ
    if any(k in row_text for k in ['طباعة بتاريخ', 'المجموعة', 'المقاس', 'الرصيد', 'الباركود', 'مجموعة ف', 'الساعة']):
        continue
        
    barcode = np.nan
    product_name = np.nan
    product_type = np.nan
    balance = np.nan
    size = np.nan
    
    # 2. فحص كل خلية بناءً على القواعد والمترادفات المحددة
    for val in row_cells:
        s_val = str(val).strip()
        
        # أ) المقاس: أي تيكت إنجليزي فيه (X, L, M, S) أو مقاسات مرقمة مثل (11-12, 13-14)
        if re.search(r'\b(S|M|L|XL|2XL|3XL|4XL|5XL|X)\b', s_val, re.I) or re.match(r'^\d{1,2}-\d{1,2}$', s_val):
            size = s_val
            
        # ب) النوع: أي خلية تحتوي على "داخلي" أو "خارجي"
        elif 'داخلي' in s_val or 'خارجي' in s_val:
            product_type = s_val
            
        # ج) الصنف: نصوص المنتجات (بشكير، فوطة، سالوبيت، كولون، شراب، طقم، فايبر، إلخ)
        elif any(kw in s_val for kw in ['فوطه', 'فوطة', 'بشكير', 'سالوبيت', 'كولون', 'شراب', 'طقم', 'فايبر', 'حرام', 'سوكت', 'برا']):
            product_name = s_val
            
        # د) الأرقام: التفريق بين الباركود والرصيد
        elif s_val.isdigit() or (s_val.replace('.', '', 1).isdigit()):
            num = float(s_val)
            # الباركود: أي رقم أكبر من 100
            if num >= 100:
                barcode = int(num)
            # الرصيد: أي رقم أقل من 100
            elif num < 100:
                balance = int(num)

    parsed_rows.append({
        'الباركود': barcode,
        'الصنف': product_name,
        'النوع': product_type,
        'الرصيد': balance,
        'المقاس': size
    })

# 3. تحويل البيانات لـ DataFrame
result_df = pd.DataFrame(parsed_rows)

# 4. التعبئة لأسفل (ffill) للأصناف والأنواع لضمان عدم ضياع الترويسات للصفوف التابعة لها
result_df['الصنف'] = result_df['الصنف'].ffill()
result_df['النوع'] = result_df['النوع'].ffill()

# 5. استبعاد الأسطر التي لا تحتوي على باركود أو رصيد
result_df = result_df.dropna(subset=['الباركود', 'الرصيد'], how='all').reset_index(drop=True)

# 6. حفظ الملف المخرَج
output_filename = 'cleaned_inventory_products.xlsx'
try:
    result_df.to_excel(output_filename, index=False)
    print(f"تم تحويل جرد البضاعة بنجاح إلى 5 أعمدة وحفظه في: {output_filename}")
except PermissionError:
    alt_filename = 'cleaned_inventory_products_v2.xlsx'
    result_df.to_excel(alt_filename, index=False)
    print(f"تنبيه: الملف كان مفتوحاً، تم الحفظ باسم: {alt_filename}")

result_df.head(15)