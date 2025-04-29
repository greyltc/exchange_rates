#!/usr/bin/env python3

# needs `pacman -Syu python-openpyxl` for pandas.ExcelWriter(...engine="openpyxl")

import argparse
import pandas
import datetime
import pathlib
import sys

# defaults
quarter = "2025Q2"
tocur = "GBP"
curs = ["USD", "EUR", "JPY", "AUD", "CAD", "CHF"]
ref_amt = 100.0
fee_percent = 0.0
out_file_name = "master_exchange_rates.xlsx"

# get user arguments
parser = argparse.ArgumentParser(description="get exchange rates", formatter_class=argparse.ArgumentDefaultsHelpFormatter)
parser.add_argument("--benchmark-currency", "-b", default=tocur, help="benchmark currency")
parser.add_argument("--currencies", "-c", nargs='+', default=curs, help="exchange currencies")
parser.add_argument("--quarter", "-q", default=quarter, help="fiscal quarter")
parser.add_argument("--out-file", "-o", default=out_file_name, help="output file name")
parser.add_argument("--amount", "-a", type=float, default=ref_amt, help="exchange amount (matters only for CCs)")
parser.add_argument("--bank-fee-percent", "-f", type=float, default=fee_percent, help="bank fee [%%] (matters only for CCs)")
parser.add_argument("--source", choices=['ECB', 'BoE', 'MCAPI', 'Visa', 'MC'], default="Visa", help="exchange rate data source")
parser.prog = sys.argv[0]
args = parser.parse_args(sys.argv[1:])

out_file_name = args.out_file
curs = args.currencies
quarter = args.quarter
tocur = args.benchmark_currency
ref_amt = args.amount
fee_percent = args.bank_fee_percent

start_date = pandas.Period(quarter).start_time.to_pydatetime()
end_date   = pandas.Period(quarter).end_time  .to_pydatetime()
start_quarter = pandas.to_datetime(start_date).to_period("Q")

# 10 days before quarter start. ensures we can interpolate a rate on day 1
before_start = start_date.date()-datetime.timedelta(days=10)

# all the days of the quarter
date_range_like = pandas.date_range(start=pandas.Timestamp(start_date), end=pandas.Timestamp(end_date), freq='1D')

headers = {
    'User-Agent': 'Mozilla/5.0 (X11; Linux x86_64) '
                'AppleWebKit/537.36 (KHTML, like Gecko) '
                'Chrome/54.0.2840.90 '
                'Safari/537.36'
}

if args.source == "ECB":
    # european central bank
    import pandasdmx
    if "GBP" not in curs:
        curs = curs + ["GBP"]
    try:
        curs.remove("EUR")
    except:
        pass
    ecb = pandasdmx.Request('ECB')
    print("Warning ECB source always returns exchange rates to EUR")
    tocur = "EUR"
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
elif args.source == "MCAPI":
    # mastercard API
    import oauth1.authenticationutils as authenticationutils
    from oauth1.oauth import OAuth
    import requests
    # from mysecrets_example import keyfilename, key_password, consumer_key
    from mysecrets import keyfilename, key_password, consumer_key
    uri_base = "https://sandbox.api.mastercard.com/enhanced/settlement/currencyrate/subscribed/summary-rates"
    signing_key = authenticationutils.load_signing_key(keyfilename, key_password)
elif args.source == "BoE":
    # bank of england
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
    # visa
    import requests
    # discovered via https://www.visa.com.my/support/consumer/travel-support/exchange-rate-calculator.html
    # eg: https://www.visa.com.my/cmsapi/fx/rates?amount=101&fee=0&utcConvertedDate=05%2F02%2F2024&exchangedate=05%2F02%2F2024&fromCurr=GBP&toCurr=USD
    url = "https://www.visa.com.my/cmsapi/fx/rates"
    when = datetime.date(2024,3,5)
    params = {
        "amount": ref_amt,
        "fee": fee_percent,
        "utcConvertedDate": when.strftime('%m/%d/%Y'),
        "exchangedate": when.strftime('%m/%d/%Y'),
        "fromCurr": "GBP",
        "toCurr": "CAD",
    }
elif args.source == "MC":
    # mastercard
    import requests
    # discovered via https://www.mastercard.us/en-us/personal/get-support/convert-currency.html
    # eg: https://www.mastercard.us/settlement/currencyrate/conversion-rate?fxDate=2024-04-01&transCurr=CAD&crdhldBillCurr=GBP&bankFee=0&transAmt=150
    url = "https://www.mastercard.us/settlement/currencyrate/conversion-rate"
    when = when = datetime.date(2024,3,5)
    params = {
        "fxDate": when.strftime('%Y-%m-%d'),
        "transCurr": "CAD",
        "crdhldBillCurr": "GBP",
        "bankFee": fee_percent,
        "transAmt": ref_amt,
    }
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
        sheet_name=f"{args.source}-{start_quarter.year}Q{start_quarter.quarter}-{cur}to{tocur}"
        if sheet_name in xls.sheets:
            print(f"Sheet {sheet_name} is already in {out_file_name}. Skipping data fetch.")
        else:
            while quarter == start_quarter:
                row = {}
                row["Date"] = str(date.date())
                if args.source == "MCAPI":
                    datestring = date.strftime("%Y-%m-%d")
                    uri = f"{uri_base}?rate_date={datestring}&trans_curr={cur}&trans_amt={str(ref_amt)}&crdhld_bill_curr={tocur}"
                    authHeader = OAuth.get_authorization_header(uri, 'GET', None, consumer_key, signing_key)
                    response = requests.get(uri, headers={'Authorization' : authHeader})
                    rjdat = response.json()["data"]
                    # print(rjdat)
                    row["currency"] = rjdat["transCurr"]
                    row[f"{rjdat['transCurr']}/{tocur} Rate"] = rjdat["effectiveConversionRate"]
                    row["Diff"] = rjdat["pctDifferenceMastercardExclAllFeesAndEcb"]
                    row["ECB Date"] = rjdat["ecb"]["ecbReferenceRateDate"]
                    row["ECB Rate"] = rjdat["ecb"]["ecbReferenceRate"]
                elif args.source == "ECB":
                    this_row = reid.loc[pandas.Timestamp(date)]
                    rate = this_row.loc[:, cur, :, :,:][0]
                    row[f"{cur}/{tocur} Rate"] = float(f"{1/rate:0.7f}")
                elif args.source == "BoE":
                    this_row = reid.loc[pandas.Timestamp(date)]
                    rate = this_row.loc[series_code_lookup[cur]]
                    row[f"{cur}/{tocur} Rate"] = float(f"{1/rate:0.7f}")
                elif args.source == "Visa":
                    datestring = date.strftime('%m/%d/%Y')
                    params["utcConvertedDate"] = datestring
                    params["exchangedate"] = datestring
                    params["fromCurr"] = tocur
                    params["toCurr"] = cur
                    response = requests.get(url, params=params, headers=headers)
                    if response.status_code == 200:
                        rjdat = response.json()
                        rate = float(rjdat["fxRateWithAdditionalFee"])
                        row[f"{cur}/{tocur} Rate"] = rate
                    else:
                        print(f"Failed fetching from: {response.url}")
                        print(f"Status code: {response.status_code}")
                        print(f"Reason: {response.reason}")
                        sys.exit(-1)
                elif args.source == "MC":
                    datestring = date.strftime('%Y-%m-%d')
                    params["fxDate"] = datestring
                    params["crdhldBillCurr"] = tocur
                    params["transCurr"] = cur
                    response = requests.get(url, params=params, headers=headers)
                    if response.status_code == 200:
                        rjdat = response.json()
                        if rjdat["type"] == "error":
                            print(f"Error fetching from: {response.url}")
                            print(f"Error code: {rjdat['data']["errorCode"]}")
                            print(f"Error message: {rjdat['data']["errorMessage"]}")
                            sys.exit(-1)
                        else:
                            rate = rjdat["data"]["conversionRate"]
                            row[f"{cur}/{tocur} Rate"] = rate
                    else:
                        print(f"Failed fetching from: {response.url}")
                        print(f"Status code: {response.status_code}")
                        print(f"Reason: {response.reason}")
                        sys.exit(-1)
                else:
                    sys.exit(-2)

                print(row)
                if len(row) > 1:
                    df = pandas.concat([df, pandas.DataFrame([row])], ignore_index=True)
                date += datetime.timedelta(days=1)
                quarter = pandas.to_datetime(date).to_period("Q")
        if not df.empty:
            df.to_excel(xls, sheet_name=sheet_name, index=False)
