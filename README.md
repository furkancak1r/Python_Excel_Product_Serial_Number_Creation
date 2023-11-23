# Python_Excel_Product_Serial_Number_Creation

## Overview
This script utilizes the `win32com.client` library in Python to interact with Microsoft Excel. The goal is to process data from a specified Excel file, perform operations, and save the results to a new sheet.

## Prerequisites
- Python installed on your system
- Required Python packages: `os`, `win32com.client`

## Usage
1. Clone the repository or download the script.
2. Make sure the required libraries are installed by running:
   ```bash
   pip install pypiwin32
Update the file_path variable in the script to point to your Excel file.
Run the script:
```
python excel_data_processing.py
```
Check the "Sonuç" sheet in the Excel file for the processed data.

## Script Explanation
The script opens the specified Excel file, assuming there's a sheet named "Özet" with relevant data and a sheet named "Sonuç" for the results.
It iterates through the rows, extracting necessary information and creating a unique "Dolap Seri No" based on the provided logic.
The processed data is then written to the "Sonuç" sheet.

## Example Input

Ensure your Excel sheet "Özet" has columns with headers like:

| Stok Kodu | Ürün Adı | Sipariş Miktarı | Sipariş No | Tarih    |
|-----------|----------|-----------------|------------|----------|
| SKU123    | Product1 | 3               | Order123   | 20220101 |
| SKU456    | Product2 | 2               | Order124   | 20220102 |

## Example Output

The "Sonuç" sheet will contain processed data:

| Stok Kodu | Ürün Adı | Adet | Sipariş No | Tarih    | Dolap Seri No        |
|-----------|----------|------|------------|----------|----------------------|
| SKU123    | Product1 | 1    | Order123   | 20220101 | Order123_SKU123_000001|
| SKU123    | Product1 | 2    | Order123   | 20220101 | Order123_SKU123_000002|
| SKU123    | Product1 | 3    | Order123   | 20220101 | Order123_SKU123_000003|
| SKU456    | Product2 | 1    | Order124   | 20220102 | Order124_SKU456_000001|
| SKU456    | Product2 | 2    | Order124   | 20220102 | Order124_SKU456_000002|

## Notes
Make sure Microsoft Excel is installed on your system.
Adjustments to the script may be needed based on your specific Excel file structure.