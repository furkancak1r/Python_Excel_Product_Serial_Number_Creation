import os
import win32com.client

# Specify the file path
file_path = "C:\\Users\\furkan.cakir\\Desktop\\FurkanPRS\\Kodlar\\SSH\\SSH Exceller\\ekipman kartı aktarım şablonu çalışması.xlsx"

# Check if the file exists
if os.path.exists(file_path):
    try:
        # Open Excel application
        excel = win32com.client.Dispatch("Excel.Application")
        workbook = excel.Workbooks.Open(file_path)
        worksheet = workbook.Sheets("2. Aşama")
        result_sheet = workbook.Sheets("3. Aşama")
        excel.Visible = False

        # Get the used range of the worksheet
        used_range = worksheet.UsedRange
        rows = used_range.Rows.Count
        cols = used_range.Columns.Count

        # Access data and perform operations
        result_row = 1 # Initialize the result row counter
        for row in range(2, rows + 1):  # Assuming data starts from the second row
            cari_hesap_no = worksheet.Cells(row, 1).Value
            cari_hesap_unvani = worksheet.Cells(row, 2).Value
            sevkiyat_kodu= worksheet.Cells(row, 3).Value
            sevkiyat_aciklamasi = worksheet.Cells(row, 4).Value
            sevkiyat_adresi = worksheet.Cells(row, 5).Value
            kodu= worksheet.Cells(row, 6).Value
            aciklamasi= worksheet.Cells(row, 7).Value
            miktar= worksheet.Cells(row, 8).Value
            birim= worksheet.Cells(row, 9).Value
            fatura_numarasi= worksheet.Cells(row, 10).Value
            tarihi= worksheet.Cells(row, 11).Value
            siparis_no= worksheet.Cells(row, 12).Value
            gruplar= worksheet.Cells(row, 13).Value
            

            # Perform your operations, for example, creating Dolap Seri No
            for i in range(1, int(miktar) + 1):
                dolap_seri_no = f"{siparis_no}_{kodu}_{str(i).zfill(6)}" # Use i instead of sip_miktar and format it as a 6-character string with leading zeros
                result_sheet.Cells(result_row, 1).Value = cari_hesap_no # Write the result to the result sheet
                result_sheet.Cells(result_row, 2).Value = cari_hesap_unvani
                result_sheet.Cells(result_row, 3).Value = sevkiyat_kodu
                result_sheet.Cells(result_row, 4).Value = sevkiyat_aciklamasi
                result_sheet.Cells(result_row, 5).Value = sevkiyat_adresi
                result_sheet.Cells(result_row, 6).Value = kodu
                result_sheet.Cells(result_row, 7).Value = aciklamasi
                result_sheet.Cells(result_row, 8).Value = 1
                result_sheet.Cells(result_row, 9).Value = birim
                result_sheet.Cells(result_row, 10).Value = fatura_numarasi
                result_sheet.Cells(result_row, 11).Value = tarihi
                result_sheet.Cells(result_row, 12).Value = siparis_no
                result_sheet.Cells(result_row, 13).Value = gruplar
                result_sheet.Cells(result_row, 14).Value = dolap_seri_no
                
                result_row += 1 # Increment the result row counter

        # Save and close the workbook
        workbook.Save()
        workbook.Close()

    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Quit the Excel application
        excel.Quit()
