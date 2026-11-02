# TotalEarning

Merged earnings tracker - Colab + GDrive

## Setup
1. Add secret `FINNHUB_API_KEY` in GitHub Settings -> Secrets -> Actions
2. (Optional) Add secret `GOOGLE_SHEET_ID`
3. Upload Watchlist.xlsx as Release asset or keep in Drive

## Run in Colab
```python
from google.colab import drive
drive.mount('/content/drive')
!git clone https://github.com/YOUR_USERNAME/TotalEarning.git
%cd TotalEarning
!pip install -r requirements.txt
!python merged_to_gdrive.py
