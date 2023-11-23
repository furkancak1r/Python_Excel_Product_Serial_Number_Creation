import os
import win32com.client

# Specify the file path
file_path = "C:/Users/furkan.cakir/Desktop/FurkanPRS/Kodlar/SAP Garanti Çalışması/garanti sorgu dahil.xlsx"

# Check if the file exists
if os.path.exists(file_path):
    try:
        # Open Excel application
        excel = win32com.client.Dispatch("Excel.Application")
        workbook = excel.Workbooks.Open(file_path)
        worksheet = workbook.Sheets("Özet")
        result_sheet = workbook.Sheets("Sonuç")
        excel.Visible = False

        # Get the used range of the worksheet
        used_range = worksheet.UsedRange
        rows = used_range.Rows.Count
        cols = used_range.Columns.Count

        # Access data and perform operations
        result_row = 1 # Initialize the result row counter
        for row in range(2, rows + 1):  # Assuming data starts from the second row
            stok_kodu = worksheet.Cells(row, 1).Value
            urun_adi = worksheet.Cells(row, 2).Value
            sip_miktar = worksheet.Cells(row, 3).Value
            siparis_no = worksheet.Cells(row, 4).Value
            tarih = worksheet.Cells(row, 5).Value

            # Perform your operations, for example, creating Dolap Seri No
            for i in range(1, int(sip_miktar) + 1):
                dolap_seri_no = f"{siparis_no}_{stok_kodu}_{str(i).zfill(6)}" # Use i instead of sip_miktar and format it as a 6-character string with leading zeros
                result_sheet.Cells(result_row, 1).Value = stok_kodu # Write the result to the result sheet
                result_sheet.Cells(result_row, 2).Value = urun_adi
                result_sheet.Cells(result_row, 3).Value = 1
                result_sheet.Cells(result_row, 4).Value = siparis_no
                result_sheet.Cells(result_row, 5).Value = tarih
                result_sheet.Cells(result_row, 6).Value = dolap_seri_no
                result_row += 1 # Increment the result row counter

        # Save and close the workbook
        workbook.Save()
        workbook.Close()

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Quit the Excel application
        excel.Quit()
