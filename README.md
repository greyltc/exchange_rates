# exchange_rates
use mastercard api to get currency exchange rates

## Usage
```
python3 -m venv .venv
source .venv/bin/activate # or `.venv\Scripts\activate`
pip install pandasdmx
git clone https://github.com/Mastercard/oauth1-signer-python.git
cd oauth1-signer-python
git checkout 7205e45
wget https://raw.githubusercontent.com/greyltc/exchange_rates/main/get_rates.py -O get_rates.py
wget https://raw.githubusercontent.com/greyltc/exchange_rates/main/mysecrets_example.py -O mysecrets.py

# insert your secrets. see https://developer.mastercard.com/platform/documentation/security-and-authentication/using-oauth-1a-to-access-mastercard-apis/
sed 's,^keyfilename = .*,keyfilename = "keyfile.p12",' -i mysecrets.py
sed 's,^key_password = .*,key_password = "yourpw",' -i mysecrets.py
sed 's,^consumer_key = .*,consumer_key = "bunchofupperlowercasealphanumerics!bunchofhex",' -i mysecrets.py

sed 's,^quarter = .*,quarter = "2023Q3",' -i get_rates.py  # update the year/quarter string

python get_rates.py
```
