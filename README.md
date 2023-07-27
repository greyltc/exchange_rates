# exchange_rates
use mastercard api to get currency exchange rates

## Usage
```
git clone https://github.com/Mastercard/oauth1-signer-python.git
cd oauth1-signer-python
git checkout 7205e45
wget https://raw.githubusercontent.com/greyltc/exchange_rates/main/get_rates.py
wget -O mysecrets.py https://raw.githubusercontent.com/greyltc/exchange_rates/main/mysecrets_example.py
vim mysecrets.py  # insert your secrets. see https://developer.mastercard.com/platform/documentation/security-and-authentication/using-oauth-1a-to-access-mastercard-apis/
vim get_rates.py  # update the year/quarter string
python get_rates.py
```
