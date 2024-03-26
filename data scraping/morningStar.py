import requests
import pdfplumber
import pandas as pd

# URL of the PDF document
pdf_url = "https://lt.morningstar.com/dk7pkae7kl/snapshotpdf/default.aspx?Id=f000000i4v&accesstoken=LchS1W2t%2fW3nNWJfvvX8%2f3CfUr1s%2f7Y16snhTV6SX%2bs%2fs08caBq0ng%3d%3d"

# Download the PDF file
response = requests.get(pdf_url)
with open("downloaded_pdf.pdf", "wb") as f:
    f.write(response.content)

# Open the downloaded PDF file
with pdfplumber.open("downloaded_pdf.pdf") as pdf:
    # Extract text data from all pages
    pdf_text = ""
    for page in pdf.pages:
        pdf_text += page.extract_text()

# Split the text into lines and create a DataFrame
lines = pdf_text.split("\n")
df = pd.DataFrame({"Data": lines})

# Save the DataFrame to an Excel file
excel_filename = "pdf_data.xlsx"
df.to_excel(excel_filename, index=False)

print(f"Data saved to {excel_filename}")
