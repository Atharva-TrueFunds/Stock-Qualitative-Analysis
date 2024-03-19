from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import pandas as pd
from openpyxl import load_workbook

#######################################################################################################
chrome_service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=chrome_service)
url = "https://www.valueresearchonline.com/funds/66/quant-small-cap-fund#fund-portfolio"
driver.get(url)

parent_div_xpath = (
    "/html/body/section[2]/div[2]/div/div[5]/div/div/section[2]/div[2]/div/div[1]/div"
)
parent_div = driver.find_element(By.XPATH, parent_div_xpath)
list_items = parent_div.find_elements(By.TAG_NAME, "p")

combined_data = [item.text for item in list_items]

df1 = pd.DataFrame({"List Items": combined_data})

excel_filename = "combined_data.xlsx"
df1.to_excel(excel_filename, index=False)
print(f"All DataFrames saved to {excel_filename}")


############################################################################################
url = "https://www.valueresearchonline.com/funds/66/quant-small-cap-fund#fund-portfolio"
driver.get(url)

parent_div_xpath = (
    "/html/body/section[2]/div[2]/div/div[5]/div/div/section[2]/div[1]/div"
)
parent_div = driver.find_element(By.XPATH, parent_div_xpath)
list_items = parent_div.find_elements(By.TAG_NAME, "li")

combined_data_2 = [item.text for item in list_items]

wb = load_workbook(excel_filename)
ws = wb.active

start_column = ws.max_column + 1
for i, data in enumerate(combined_data_2, start=1):
    ws.cell(row=i, column=start_column, value=data)

wb.save(excel_filename)
print(f"Second section data appended to {excel_filename}")

#############################################################################################


def process_url(url):
    driver.get(url)
    tables = driver.find_elements(By.TAG_NAME, "table")

    dfs = []
    for table in tables:
        table_html = table.get_attribute("outerHTML")
        df = pd.read_html(table_html)
        if df:
            df = df[0].dropna(how="all").reset_index(drop=True)
            dfs.append(df)

    return dfs


urls = [
    "https://www.valueresearchonline.com/funds/66/quant-small-cap-fund#overview",
    "https://www.valueresearchonline.com/funds/66/quant-small-cap-fund#performance",
    "https://www.valueresearchonline.com/funds/66/quant-small-cap-fund#fund-portfolio",
    "https://www.valueresearchonline.com/funds/66/quant-small-cap-fund#risk",
    "https://www.valueresearchonline.com/funds/66/quant-small-cap-fund#other",
]

all_dfs = []
for url in urls:
    dfs = process_url(url)
    all_dfs.extend(dfs)

if all_dfs:
    combined_df = pd.concat(all_dfs, ignore_index=True)
    wb = load_workbook(excel_filename)
    ws = wb.active

    # Find the first empty column
    start_column = 1
    while start_column <= ws.max_column:
        empty_column = True
        for cell in ws.iter_rows(
            min_row=2, min_col=start_column, max_row=ws.max_row, max_col=start_column
        ):
            for row in cell:
                if row.value:
                    empty_column = False
                    break
            if not empty_column:
                break

        if empty_column:
            ws.delete_cols(start_column)
        else:
            start_column += 1
    start_column = ws.max_column + 1

    for col_index, column in enumerate(combined_df.columns, start=start_column):
        ws.cell(row=1, column=col_index, value=column)

    start_row = 2
    for _, row in combined_df.iterrows():
        for col_index, value in enumerate(row, start=start_column):
            ws.cell(row=start_row, column=col_index, value=value)
        start_row += 1

    wb.save(excel_filename)
    print(f"Table data appended to new empty columns in {excel_filename}")
else:
    print("No tables found on any of the webpages.")

driver.quit()