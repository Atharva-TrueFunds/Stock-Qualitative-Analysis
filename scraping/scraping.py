from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
from urllib.parse import urlparse
from openpyxl import load_workbook

#####################################################################################################
#####################################################################################################

# portfolio


def process_tables(url):
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


chrome_service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=chrome_service)

urls = [
    "https://www.valueresearchonline.com/funds/66/quant-small-cap-fund#fund-portfolio",
    "https://www.valueresearchonline.com/funds/16198/sbi-bluechip-fund-direct-plan/#fund-portfolio",
    "https://www.valueresearchonline.com/funds/11213/mirae-asset-large-midcap-fund-regular-plan/?#fund-portfolio",
    "https://www.valueresearchonline.com/funds/2310/icici-prudential-value-discovery-fund/#fund-portfolio",
]

for index, url in enumerate(urls):
    chrome_service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=chrome_service)
    driver.get(url)

    parent_div_xpath_h3 = "/html/body/section[2]/div[2]/div/div[5]/div/div/section[2]/div[2]/div/div[1]/div/div/div[1]/div[1]"
    parent_div_h3 = driver.find_element(By.XPATH, parent_div_xpath_h3)
    list_items_h3 = parent_div_h3.find_elements(By.XPATH, ".//p")
    combined_data_h3 = [item.text for item in list_items_h3]

    parent_div_xpath_p = "/html/body/section[2]/div[2]/div/div[5]/div/div/section[2]/div[2]/div/div[1]/div/div/div[1]/div[3]"
    parent_div_p = driver.find_element(By.XPATH, parent_div_xpath_p)
    list_items_p = parent_div_p.find_elements(By.XPATH, ".//p")
    combined_data_p = [item.text for item in list_items_p]

    parent_div_xpath_li = (
        "/html/body/section[2]/div[2]/div/div[5]/div/div/section[2]/div[1]/div/ul"
    )
    parent_div_li = driver.find_element(By.XPATH, parent_div_xpath_li)
    list_items_li = parent_div_li.find_elements(By.XPATH, ".//span")
    combined_data_li = [item.text for item in list_items_li]

    all_table_dfs = process_tables(url)

    max_len = max(
        len(combined_data_h3),
        len(combined_data_p),
        len(combined_data_li),
        *[len(df) for df in all_table_dfs],
    )

    combined_data_h3.extend([""] * (max_len - len(combined_data_h3)))
    combined_data_p.extend([""] * (max_len - len(combined_data_p)))
    combined_data_li.extend([""] * (max_len - len(combined_data_li)))

    combined_df_temp = pd.DataFrame(
        {
            "List Items (h3)": combined_data_h3,
            "List Items (p)": combined_data_p,
            "List Items (li)": combined_data_li,
        }
    )

    for df in all_table_dfs:
        for col in df.columns:
            if col in combined_df_temp.columns:
                combined_df_temp[col] = (
                    combined_df_temp[col].astype(str) + "\n" + df[col].astype(str)
                )
                combined_df_temp[col] = combined_df_temp[col].str.replace(
                    r"(?:^|\n)nan(?:$|\n)", "\n", regex=True
                )
            else:
                combined_df_temp[col] = df[col]

    parsed_url = urlparse(url)
    fund_name = parsed_url.path.split("/")[3].replace("-", " ").title()[:31]

    combined_df_temp["Fund Name"] = fund_name

    # Saving data to Excel
    excel_filename = "combined_data.xlsx"
    if index == 0:
        combined_df_temp.to_excel(excel_filename, sheet_name=fund_name, index=False)
    else:
        with pd.ExcelWriter(excel_filename, engine="openpyxl", mode="a") as writer:
            combined_df_temp.to_excel(writer, sheet_name=fund_name, index=False)

    print(f"All Data saved to {excel_filename}")

driver.quit()

####################################################################################################
####################################################################################################
