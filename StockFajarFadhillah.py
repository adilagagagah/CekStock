import pandas as pd
from datetime import datetime
import locale

file_name = "DATA BASE DOS & STOCK 12 SEPT 2024.xlsx"
file_path = f"../Cek Stock Fajar Fadhillah/{file_name}"
df = pd.read_excel(file_path, sheet_name="dos by store-brand type & area")
site_codes = ['E423', 'E491', 'E288', 'E371']
result_rows = []

# Membuat writer Excel
locale.setlocale(locale.LC_TIME, 'Indonesian')
tanggal_sekarang = datetime.now()
output_file_name = file_name[22:-10]
output_file = f"../Cek Stock Fajar Fadhillah/{output_file_name}.xlsx"
tanggal_format = tanggal_sekarang.strftime('%A, %d %B %Y')

def merge_cells(worksheet, result_df, start_row, merge_format):
    """
    Fungsi untuk melakukan merge cells pada kolom BRAND TYPE, STOCK, SALES, dan DOS
    berdasarkan nilai yang sama.

    Args:
    worksheet: Worksheet xlsxwriter di mana merge akan diterapkan.
    result_df: DataFrame berisi data yang akan dimasukkan ke dalam worksheet.
    start_row: Baris awal untuk mulai menulis data pada worksheet.
    merge_format: Format yang digunakan untuk merge cells.
    """
    previous_value = None

    # Mulai iterasi dari baris awal data yang akan ditulis
    for i, product in enumerate(result_df['BRAND TYPE'], start=start_row):
        if product == previous_value:
            continue

        # Cek berapa kali nilai tersebut berulang
        same_products_count = (result_df['BRAND TYPE'] == product).sum()

        if same_products_count > 1:
            # Merge untuk kolom BRAND TYPE, STOCK, SALES, dan DOS
            worksheet.merge_range(i, 1, i + same_products_count - 1, 1, product, merge_format)
            worksheet.merge_range(i, 2, i + same_products_count - 1, 2, result_df['STOCK'].iloc[i - start_row].values[0], merge_format)
            worksheet.merge_range(i, 3, i + same_products_count - 1, 3, result_df['SALES'].iloc[i - start_row].values[0], merge_format)
            worksheet.merge_range(i, 4, i + same_products_count - 1, 4, result_df['DOS'].iloc[i - start_row].values[0], merge_format)
        
        previous_value = product

def format_cells(worksheet, result_df, start_row, center_format):
    for row in range(len(result_df)):
        for col in range(len(result_df.columns)):
            worksheet.write(start_row + row, col, result_df.iloc[row, col], center_format)

def auto_adjust_column_width(worksheet, df):
    """Menyesuaikan lebar kolom berdasarkan konten sel"""
    for idx, col in enumerate(df.columns):
        try:
            max_len_data = df[col].astype(str).map(len).max() if not df[col].isnull().all() else 0
        except:
            max_len_data = 0
        max_len_header = len(col)
        max_len = max(max_len_data, max_len_header) + 2 
        worksheet.set_column(idx, idx, max_len)
        
def set_row_heights(worksheet, result_df, start_row, height=30):
    for row in range(start_row, start_row + len(result_df)):
        worksheet.set_row(row, height)


with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    current_row = 4
    
    for code in site_codes:
        min_dos_df = df[(df['SITE CODE'] == code) & (df['DOS 30 days'] <= 15)]
        min_dos_df = min_dos_df.sort_values(by='DOS 30 days', ascending=True)
        if min_dos_df.empty:
            continue

        min_dos_city = list(min_dos_df['KOTA'])[0]
        min_dos_store = list(min_dos_df['STORE NAME'])[0] + f" ({code})"
        min_dos_product = list(min_dos_df['Article code no color'])

        rotation_df = df[(df['TSH'] == 'FAJAR FADHILLAH') & 
                        (df['KOTA'] == min_dos_city) & 
                        (-df['SITE CODE'].isin(site_codes)) & 
                        (df['DOS 30 days'] >= 45) & 
                        (df['Article code no color'].isin(min_dos_product))]

        # Iterasi untuk setiap produk di min_dos_product
        for product in min_dos_product:
            # Data dari min_dos_df untuk produk yang cocok
            min_dos_product_data = min_dos_df[min_dos_df['Article code no color'] == product]
            
            min_stock = min_dos_product_data['Stock'].values[0]
            min_sales = min_dos_product_data['Sales 30 days'].values[0]
            min_dos = min_dos_product_data['DOS 30 days'].values[0]
                
            # Data dari rotation_df untuk produk yang cocok
            rotation_product_data = rotation_df[rotation_df['Article code no color'] == product]
            rotation_product_data = rotation_product_data.sort_values(by='DOS 30 days', ascending=False)
            
            for index, row in rotation_product_data.iterrows():
                # Ambil data rotation untuk setiap toko yang cocok
                rotation_store = row['STORE NAME']
                rotation_stock = row['Stock']
                rotation_sales = row['Sales 30 days']
                rotation_dos = row['DOS 30 days']
                    
                # Menambahkan ke list hasil
                result_rows.append([
                    min_dos_store, product, min_stock, min_sales, min_dos,
                    rotation_store, rotation_stock, rotation_sales, rotation_dos
                ])

        result_df = pd.DataFrame(result_rows, columns=[
            'STORE NAME', 'BRAND TYPE', 'STOCK', 'SALES', 'DOS',
            'ROTASI DARI', 'STOCK', 'SALES', 'DOS'
        ])
        
        result_df.to_excel(writer, sheet_name='FAJAR KEMITRAAN', index=False, startrow=current_row)
        
        workbook = writer.book
        bold_format = workbook.add_format({'bold': True})
        center_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })
        merge_format = workbook.add_format({
            'align': 'center',
            'valign': 'vcenter',
            'bold': True,
            'border': 1
        })
        
        worksheet = writer.sheets['FAJAR KEMITRAAN']
        worksheet.write(0, 0, f"UPDATE : {tanggal_format}", bold_format)
        worksheet.write((current_row-2), 0, f"Store Name : {min_dos_store}", bold_format)
        merge_cells(worksheet, result_df, start_row=current_row + 1, merge_format=merge_format)
        format_cells(worksheet, result_df, start_row=current_row + 1, center_format=center_format)
        worksheet.merge_range((current_row+1), 0, (current_row + len(result_df)), 0, min_dos_store, merge_format)
        
        set_row_heights(worksheet, result_df, start_row=current_row + 1, height=30)
        auto_adjust_column_width(worksheet, result_df)
        
        # Update current_row untuk menambahkan jeda 5 baris kosong
        current_row += len(result_df) + 5

        result_rows = []

print(f"Hasil telah disimpan dengan nama {output_file}")
