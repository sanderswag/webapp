from flask import Flask, request, render_template, send_file
import pandas as pd
import requests
from bs4 import BeautifulSoup
import re
import os

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "output"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Function to scrape company data
def scrape_company_info(org_number):
    url = f"https://www.proff.no/aksjon%C3%A6rer/-/-/{org_number}"
    headers = {"User-Agent": "Mozilla/5.0"}
    
    try:
        response = requests.get(url, headers=headers, timeout=10)
        response.raise_for_status()
    except requests.RequestException:
        return "Unknown", []
    
    soup = BeautifulSoup(response.text, "html.parser")
    company_name_tag = soup.find("h1")
    company_name = company_name_tag.get_text(strip=True) if company_name_tag else "Unknown"
    
    shareholders = []
    table = soup.find("table")
    if table:
        for row in table.find_all("tr")[1:]:
            cols = row.find_all("td")
            if len(cols) >= 4:
                name = re.sub(r"\(Ordinære aksjer\)", "", cols[0].get_text(strip=True)).strip()
                name = re.sub(r"Org nr", ", Org nr:", name)
                name = re.sub(r"Født", ", Født", name)
                percentage = cols[3].get_text(strip=True)
                shareholders.append((name, percentage))
    
    return company_name, shareholders

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        if "file" not in request.files:
            return "No file uploaded."
        file = request.files["file"]
        if file.filename == "":
            return "No selected file."
        
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)
        
        df = pd.read_excel(file_path, header=None)
        df.rename(columns={df.columns[0]: "OrgNumber"}, inplace=True)
        
        result_data = []
        max_shareholders = 0
        
        for org_number in df["OrgNumber"].astype(str):
            company_name, shareholders = scrape_company_info(org_number)
            max_shareholders = max(max_shareholders, len(shareholders))
            row_data = {"OrgNumber": org_number, "CompanyName": company_name}
            
            for i, (name, percentage) in enumerate(shareholders):
                row_data[f"Shareholder_{i+1}"] = name
                row_data[f"Andel_{i+1}"] = percentage
            
            result_data.append(row_data)
        
        columns = ["OrgNumber", "CompanyName"] + [f"Shareholder_{i+1}" for i in range(max_shareholders)] + [f"Andel_{i+1}" for i in range(max_shareholders)]
        result_df = pd.DataFrame(result_data, columns=columns)
        
        output_file = os.path.join(OUTPUT_FOLDER, "processed_data.xlsx")
        result_df.to_excel(output_file, index=False)
        
        return send_file(output_file, as_attachment=True)
    
    return render_template("index.html")

if __name__ == "__main__":
    app.run(debug=True)

# Rename script for deployment
if __name__ != "__main__":
    eierweb = app
