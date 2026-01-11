# CryptoSheet Oracle

![License](https://img.shields.io/github/license/your-username/CryptoSheet-Oracle)
![Stars](https://img.shields.io/github/stars/your-username/CryptoSheet-Oracle)
![Issues](https://img.shields.io/github/issues/your-username/CryptoSheet-Oracle)
![PRs Welcome](https://img.shields.io/badge/PRs-welcome-brightgreen)

# CryptoSheet Oracle

CryptoSheet Oracle is an open-source Google Apps Script tool that brings real-time cryptocurrency prices directly into Google Sheets using the CoinMarketCap API.

The project is designed for portfolio tracking, financial dashboards, and spreadsheet automation, with a strong focus on simplicity, performance, and secure API usage.

---

## Features

- Live cryptocurrency prices in USD
- Supports all cryptocurrencies listed on CoinMarketCap
- Works directly inside Google Sheets
- Formula-based usage in spreadsheet cells
- Bulk price updates for ranges
- Script Cache to reduce API calls
- Error recovery for temporary API failures
- Price lookup by symbol or CoinMarketCap ID

---

## Use Cases

- Crypto portfolio tracking
- Personal finance spreadsheets
- Investment dashboards
- Automated reporting
- Educational and open-source projects

---

## Screenshots

Screenshots are available in the `/screenshots` directory.

- `formula-example.png` – crypto price formula usage in Google Sheets
- `table-example.png` – automatic price updates in a table
- `crypto-list.png` – full cryptocurrency list imported from CoinMarketCap

---

## Installation

### Requirements

- Google account
- Google Sheets
- CoinMarketCap API key (free tier is sufficient)

---

### Step 1. Get a CoinMarketCap API Key

1. Create an account on CoinMarketCap  
   https://coinmarketcap.com/api/

2. Generate an API key in your dashboard

---

### Step 2. Create or Open a Google Sheet

1. Open Google Sheets  
   https://docs.google.com/spreadsheets/

2. Create a new spreadsheet or open an existing one

---

### Step 3. Open Google Apps Script

1. In Google Sheets, open:
   Extensions → Apps Script

2. A new Apps Script project will open in a separate tab

Official documentation:  
https://developers.google.com/apps-script

---

### Step 4. Add the Script

1. Remove any existing code in the editor
2. Copy the content of `src/cryptoSheetOracle.gs`
3. Paste it into the Apps Script editor
4. Save the project

---

### Step 5. Store API Key Securely

Do not hardcode API keys in the source code.

1. In Apps Script, open:
   Project Settings → Script Properties

2. Add a new Script Property:

Key: CMC_API_KEY  
Value: your_api_key_here

3. Save the changes

---

### Step 6. Authorize the Script

1. In Apps Script, run any function (for example `getCryptoPrice`)
2. Google will request permissions
3. Review and approve access

---

## Usage

### Get crypto price by symbol

Use directly in Google Sheets cells:

=getCryptoPrice("BTC")  
=getCryptoPrice("ETH")

---

### Get price by CoinMarketCap ID

=getCryptoPriceId("", 1)

---

### Bulk price update

Spreadsheet layout:

Column A: Symbol  
Column B: Price

Run the function:

updatePrices()

---

## Example Spreadsheet

An example spreadsheet is available in:

/examples/example-sheet.xlsx

Structure:

Symbol | Price (USD)  
BTC | =getCryptoPrice(A2)  
ETH | =getCryptoPrice(A3)  
SOL | =getCryptoPrice(A4)

---

## Caching Strategy

- Prices are cached for 5 minutes
- Reduces API usage
- Improves performance and stability
- Helps stay within free API limits

---

## Error Handling

- Temporary API issues return #ERROR!
- updateErrorCells() retries failed formulas
- Prevents broken sheets during API downtime

---

## Security

- API keys are stored in Script Properties
- No secrets are committed to GitHub
- Safe for public repositories

---

## Contributing

Contributions are welcome.

You can help by:
- Improving performance
- Adding new features
- Fixing bugs
- Improving documentation

### How to contribute

1. Fork the repository
2. Create a feature branch
3. Commit your changes
4. Open a pull request

---

## Roadmap

- Multi-currency support (EUR, UAH, GBP)
- Scheduled automatic refresh
- Percentage change tracking
- Historical prices
- Chart generation

---

## License

MIT License

This project is free to use, modify, and distribute.

