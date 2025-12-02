# Radiology Quotation Generator

This app automatically fills an Excel quotation template using a multi-sheet radiology charges file.

## Features
- Reads your Excel quotation template exactly as-is  
- Detects where patient info & tariff table belong  
- Reads tariffs from multi-tab charge sheet  
- Automatically selects correct sheet and tariff  
- Generates a perfect Excel quotation  

## Files Required
Place these in the repository root:
- quotation_template.xlsx
- charges.xlsx
- app.py
- requirements.txt

## Running Locally
pip install -r requirements.txt
streamlit run app.py

## Deploy on Streamlit Cloud
1. Push repo to GitHub  
2. Go to https://share.streamlit.io  
3. Select your repo  
4. Deploy  

