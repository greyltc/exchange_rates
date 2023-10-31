#!/usr/bin/env python3

# needs `pacman -Syu python-openpyxl`


import oauth1.authenticationutils as authenticationutils
from oauth1.oauth import OAuth
import requests
import pandas
import datetime
import pathlib
# from mysecrets_example import keyfilename, key_password, consumer_key
from mysecrets import keyfilename, key_password, consumer_key

quarter = "2023Q2"
curs = ["USD", "EUR", "JPY", "AUD", "CAD"]

ref_amt = str(100)
out_file_name = "exchange_rates.xlsx"

start_date = pandas.Period(quarter).start_time.to_pydatetime()
start_quarter = pandas.to_datetime(start_date).to_period("Q")

uri_base = "https://sandbox.api.mastercard.com/enhanced/settlement/currencyrate/subscribed/summary-rates"
signing_key = authenticationutils.load_signing_key(keyfilename, key_password)

if pathlib.Path(out_file_name).is_file():
    filemode = "a"
else:
    filemode = "w"

with pandas.ExcelWriter(out_file_name, mode=filemode, engine="openpyxl") as xls:
    for cur in curs:
        df = pandas.DataFrame()
        date = start_date
        quarter = start_quarter
        while quarter == start_quarter:
            row = {}
            datestring = date.strftime("%Y-%m-%d")
            uri = f"{uri_base}?rate_date={datestring}&trans_curr={cur}&trans_amt={ref_amt}&crdhld_bill_curr=GBP"
            authHeader = OAuth.get_authorization_header(uri, 'GET', None, consumer_key, signing_key)
            response = requests.get(uri, headers={'Authorization' : authHeader})
            rjdat = response.json()["data"]
            # print(rjdat)
            row["Date"] = rjdat["rateDate"]
            row["currency"] = rjdat["transCurr"]
            row[f"{rjdat['transCurr']} Rate"] = rjdat["effectiveConversionRate"]
            row["Diff"] = rjdat["pctDifferenceMastercardExclAllFeesAndEcb"]
            row["ECB Date"] = rjdat["ecb"]["ecbReferenceRateDate"]
            row["ECB Rate"] = rjdat["ecb"]["ecbReferenceRate"]
            print(row)
            df = pandas.concat([df, pandas.DataFrame([row])], ignore_index=True)
            date += datetime.timedelta(days=1)
            quarter = pandas.to_datetime(date).to_period("Q")
        df.to_excel(xls, sheet_name=f"{start_quarter.year}Q{start_quarter.quarter}-{cur}", index=False)
