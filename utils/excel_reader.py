import os
import openpyxl

def load_test_data(path=None):
    """
    Returns dict:
      {
        "vegetables": [..],
        "quantities": [..],
        "promocode": "...",
        "country": "IN"  # or whatever value your select expects
      }
    """
    if path is None:
        path = os.path.join(os.getcwd(), "GreenCartOrder.xlsx")
    wb = openpyxl.load_workbook(path)
    sheet = wb.active
    vegetables = []
    quantities = []
    # assuming data layout similar to original script: col1=name, col2=qty, col3=promo, col4=country (row2 onwards for veg)
    for i in range(2, sheet.max_row + 1):
        name = sheet.cell(row=i, column=1).value
        q = sheet.cell(row=i, column=2).value
        if name is None:
            continue
        vegetables.append(str(name))
        try:
            quantities.append(int(q))
        except Exception:
            quantities.append(1)
    promocode = sheet.cell(row=2, column=3).value or ""
    country = sheet.cell(row=2, column=4).value or ""
    return {
        "vegetables": vegetables,
        "quantities": quantities,
        "promocode": str(promocode).strip(),
        "country": str(country).strip()
    }