# exchange_rates
get currency exchange rates

## Usage
```
$ ./get-rates.py --help
usage: ./get-rates.py [-h] [--benchmark-currency BENCHMARK_CURRENCY] [--currencies CURRENCIES [CURRENCIES ...]] [--quarter QUARTER] [--out-file OUT_FILE] [--amount AMOUNT] [--bank-fee-percent BANK_FEE_PERCENT] [--source {ECB,BoE,MCAPI,Visa,MC}]

get exchange rates

options:
  -h, --help            show this help message and exit
  --benchmark-currency BENCHMARK_CURRENCY, -b BENCHMARK_CURRENCY
                        benchmark currency (default: GBP)
  --currencies CURRENCIES [CURRENCIES ...], -c CURRENCIES [CURRENCIES ...]
                        exchange currencies (default: ['USD', 'EUR', 'JPY', 'AUD', 'CAD'])
  --quarter QUARTER, -q QUARTER
                        fiscal quarter (default: 2024Q1)
  --out-file OUT_FILE, -o OUT_FILE
                        output file name (default: master_exchange_rates.xlsx)
  --amount AMOUNT, -a AMOUNT
                        exchange amount (matters only for CCs) (default: 100.0)
  --bank-fee-percent BANK_FEE_PERCENT, -f BANK_FEE_PERCENT
                        bank fee [%] (matters only for CCs) (default: 0.0)
  --source {ECB,BoE,MCAPI,Visa,MC}
                        exchange rate data source (default: Visa)
```

## Install
```
# get the script
wget https://raw.githubusercontent.com/greyltc/exchange_rates/main/get-rates.py -O get-rates.py
chmod +x get-rates.py

# install python-openpyxl
pacman -Syu python-openpyxl

# needed for BoE source
python3 -m venv .venv
source .venv/bin/activate # or `.venv\Scripts\activate`
pip install pandasdmx
# end BoE needs

# needed for MCAPI source
git clone https://github.com/Mastercard/oauth1-signer-python.git
cd oauth1-signer-python
git checkout 7205e45
wget https://raw.githubusercontent.com/greyltc/exchange_rates/main/mysecrets_example.py -O mysecrets.py

# insert your secrets. see https://developer.mastercard.com/platform/documentation/security-and-authentication/using-oauth-1a-to-access-mastercard-apis/
sed 's,^keyfilename = .*,keyfilename = "keyfile.p12",' -i mysecrets.py
sed 's,^key_password = .*,key_password = "yourpw",' -i mysecrets.py
sed 's,^consumer_key = .*,consumer_key = "bunchofupperlowercasealphanumerics!bunchofhex",' -i mysecrets.py

sed 's,^quarter = .*,quarter = "2023Q3",' -i get_rates.py  # update the year/quarter string
# end MCAPI needs

python get-rates.py
```
