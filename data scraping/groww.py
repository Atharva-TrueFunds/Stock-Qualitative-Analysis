from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import pandas as pd
from openpyxl import load_workbook
from urllib.parse import urlparse
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

wait_duration = 10
chrome_service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=chrome_service)


def process_tables(url):
    driver.get(url)

    table_xpaths = [
        "/html/body/div/div[2]/div[2]/div[1]/div/div[1]/div[2]/div/div[1]/div/div[5]"
    ]
    WebDriverWait(driver, wait_duration).until(
        EC.visibility_of_element_located((By.XPATH, table_xpaths))
    )
    dfs = []
    from io import StringIO

    # Inside the process_tables function
    for xpath in table_xpaths:
        table = driver.find_element(By.XPATH, xpath)
        table_html = table.get_attribute("outerHTML")
        # Wrap table_html in a StringIO object
        table_string_io = StringIO(table_html)
        df = pd.read_html(table_string_io)
        if df:
            df = df[0].dropna(how="all").reset_index(drop=True)
            dfs.append(df)


# Single URL to process
url = "https://groww.in/mutual-funds/quant-small-cap-fund-direct-plan-growth"

# Process tables for the single URL
all_dfs = process_tables(url)

if not all_dfs:
    combined_df_temp = pd.DataFrame({"No Data": [float("nan")]})
else:
    combined_df_temp = pd.DataFrame()  # Initialize DataFrame
    # Concatenate all DataFrames
    for df_index, df in enumerate(all_dfs):
        for col_index, col in enumerate(df.columns):
            new_col_name = (
                f"{col} - {df_index}"  # Create new column name based on DataFrame index
            )
            combined_df_temp[new_col_name] = df[col]

combined_df_temp = combined_df_temp.dropna(axis=1, how="all")

parsed_url = urlparse(url)
fund_name = parsed_url.path.split("/")[2].replace("-", " ").title()[:31]
excel_filename = "groww_data.xlsx"

wb = load_workbook(excel_filename)

if fund_name in wb.sheetnames:
    ws = wb[fund_name]
    start_column = ws.max_column + 1 if ws.max_column else 1
else:
    ws = wb.create_sheet(fund_name)
    start_column = 1

for col_index, col_name in enumerate(combined_df_temp.columns, start=start_column):
    ws.cell(row=1, column=col_index).value = col_name
    for row_index, data in enumerate(combined_df_temp[col_name], start=2):
        ws.cell(row=row_index, column=col_index, value=data)

wb.save(excel_filename)
print(f"Data added to '{fund_name}' sheet in {excel_filename}")

driver.quit()
