#!/usr/bin/env python3

# needs `pacman -Syu python-openpyxl`

import argparse
import pandas
import datetime
import pathlib
import sys

quarter = "2024Q1"
curs = ["USD", "EUR", "JPY", "AUD", "CAD"]

ref_amt = str(100)
out_file_name = "exchange_rates.xlsx"
ECB=True  # if true, use the European Central Bank rates instead of the mastercard ones

parser = argparse.ArgumentParser(description="get exchange rates", formatter_class=argparse.ArgumentDefaultsHelpFormatter)
parser.add_argument("--benchmark-currency", "-b", default="GBP", help="benchmark currency")
parser.add_argument("--currencies", "-c", nargs='+', default=curs, help="exchange currencies")
parser.add_argument("--quarter", "-q", default=quarter, help="fiscal quarter")
parser.add_argument("--out-file", "-o", default=out_file_name, help="output file name")
parser.add_argument("--source", choices=['ECB', 'BoE', 'MC', 'Visa'], default="Visa")

parser.prog = sys.argv[0]
args = parser.parse_args(sys.argv[1:])
out_file_name = args.out_file
curs = args.currencies
quarter = args.quarter

start_date = pandas.Period(quarter).start_time.to_pydatetime()
end_date   = pandas.Period(quarter).end_time  .to_pydatetime()
start_quarter = pandas.to_datetime(start_date).to_period("Q")

before_start = start_date.date()-datetime.timedelta(days=10)

date_range_like = pandas.date_range(start=pandas.Timestamp(start_date), end=pandas.Timestamp(end_date), freq='1D')

if args.source == "ECB":
    import pandasdmx
    if "GBP" not in curs:
        curs = curs + ["GBP"]
    try:
        curs.remove("EUR")
    except:
        pass
    ecb = pandasdmx.Request('ECB')
    print("Warning ECB source always returns exchange rates to EUR")
    print("Fetching exchange rate data from the European Central Bank...")
    data_response = ecb.data(resource_id = 'EXR', key={'CURRENCY': curs, "FREQ": "D"}, params = {'startPeriod': str(before_start), 'endPeriod': str(end_date.date()), })
    if data_response.response.status_code == 200:
        print(f"Sucessful fetch from: {data_response.response.url}")
    else:
        print(f"Failed fetching from: {data_response.response.url}")
        print(f"Status code: {data_response.response.status_code}")
        print(f"Reason: {data_response.response.reason}")
        sys.exit(-1)
    data = data_response.data[0]
    pdata = pandasdmx.to_pandas(data, datetime='TIME_PERIOD')
    reid = pdata.reindex(date_range_like, method='ffill')
elif args.source == "MC":
    import oauth1.authenticationutils as authenticationutils
    from oauth1.oauth import OAuth
    import requests
    # from mysecrets_example import keyfilename, key_password, consumer_key
    from mysecrets import keyfilename, key_password, consumer_key
    uri_base = "https://sandbox.api.mastercard.com/enhanced/settlement/currencyrate/subscribed/summary-rates"
    signing_key = authenticationutils.load_signing_key(keyfilename, key_password)
elif args.source == "BoE":
    import requests
    import io
    url_endpoint = 'http://www.bankofengland.co.uk/boeapps/iadb/fromshowcolumns.asp?csv.x=yes'
    series_code_lookup={}
    series_code_lookup["USD"] = "XUDLUSS"
    series_code_lookup["EUR"] = "XUDLERS"
    series_code_lookup["JPY"] = "XUDLJYS"
    series_code_lookup["CAD"] = "XUDLCDS"
    series_code_lookup["AUD"] = "XUDLADS"

    series_codes = [series_code_lookup[c] for c in curs]
    
    sd = before_start.strftime("%d/%b/%Y")
    ed = end_date.strftime("%d/%b/%Y")

    payload = {
        'Datefrom'   : sd,
        'Dateto'     : ed,
        'SeriesCodes': ','.join(series_codes),
        'CSVF'       : 'TN',
        'UsingCodes' : 'Y',
        'VPD'        : 'Y',
        'VFD'        : 'N'
    }

    headers = {
        'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) '
                    'AppleWebKit/537.36 (KHTML, like Gecko) '
                    'Chrome/54.0.2840.90 '
                    'Safari/537.36'
    }

    response = requests.get(url_endpoint, params=payload, headers=headers)
    if response.status_code == 200:
        print(f"Sucessful fetch from: {response.url}")
    else:
        print(f"Failed fetching from: {response.url}")
        print(f"Status code: {response.status_code}")
        print(f"Reason: {response.reason}")
        sys.exit(-1)
    df = pandas.read_csv(io.BytesIO(response.content))
    df.index = pandas.to_datetime(df.DATE)
    df.drop(columns=['DATE'])
    reid = df.reindex(date_range_like, method='ffill')
elif args.source=="Visa":
    import requests
    url = "https://www.visa.com.my/cmsapi/fx/rates"
    when = datetime.date(2024,3,5)
    params = {
        "amount": 100,
        "fee": 0.0,
        "utcConvertedDate": when.strftime('%m/%d/%Y'),
        "exchangedate": when.strftime('%m/%d/%Y'),
        "fromCurr": "GBP",
        "toCurr": "CAD",
    }

    headers = {
        'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) '
                    'AppleWebKit/537.36 (KHTML, like Gecko) '
                    'Chrome/54.0.2840.90 '
                    'Safari/537.36'
    }

    response = requests.get(url, params=params, headers=headers)
    if response.status_code == 200:
        print(f"Sucessful fetch from: {response.url}")
        rate = float(response.json()["fxRateWithAdditionalFee"])
    else:
        print(f"Failed fetching from: {response.url}")
        print(f"Status code: {response.status_code}")
        print(f"Reason: {response.reason}")
        sys.exit(-1)

    print(f"{rate=}")
else:
    sys.exit(-2)

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
            row["Date"] = str(date.date())
            if args.source == "MC":
                datestring = date.strftime("%Y-%m-%d")
                uri = f"{uri_base}?rate_date={datestring}&trans_curr={cur}&trans_amt={ref_amt}&crdhld_bill_curr=GBP"
                authHeader = OAuth.get_authorization_header(uri, 'GET', None, consumer_key, signing_key)
                response = requests.get(uri, headers={'Authorization' : authHeader})
                rjdat = response.json()["data"]
                # print(rjdat)
                row["currency"] = rjdat["transCurr"]
                row[f"{rjdat['transCurr']} Rate"] = rjdat["effectiveConversionRate"]
                row["Diff"] = rjdat["pctDifferenceMastercardExclAllFeesAndEcb"]
                row["ECB Date"] = rjdat["ecb"]["ecbReferenceRateDate"]
                row["ECB Rate"] = rjdat["ecb"]["ecbReferenceRate"]
            elif args.source == "ECB":
                this_row = reid.loc[pandas.Timestamp(date)]
                rate = this_row.loc[:, cur, :, :,:][0]
                row[f"{cur} Rate"] = float(f"{1/rate:0.7f}")
            elif args.source == "BoE":
                this_row = reid.loc[pandas.Timestamp(date)]
                rate = this_row.loc[series_code_lookup[cur]]
                row[f"{cur} Rate"] = float(f"{1/rate:0.7f}")
            elif args.source == "Visa":
                pass

            print(row)
            df = pandas.concat([df, pandas.DataFrame([row])], ignore_index=True)
            date += datetime.timedelta(days=1)
            quarter = pandas.to_datetime(date).to_period("Q")
        df.to_excel(xls, sheet_name=f"{start_quarter.year}Q{start_quarter.quarter}-{cur}", index=False)
