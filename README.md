# Stock Analysis for my Grandpa

A daily stock analysis workbook that builds itself.

Every weekday this repository downloads closing prices for a forty-name
watchlist, works out every moving average, Bollinger Band, MACD reading and
volume comparison, pulls the three most recent headlines for each ticker, and
writes the whole thing into an Excel file — laid out exactly the way it has
always been laid out by hand, with a great deal more in the columns to the right.

There is a website too. One button, one spreadsheet.

---

## The website

Open the GitHub Pages site for this repository. You get:

- **A single Generate button.** It downloads today's `.xlsx`. Nothing to install,
  nothing to configure.
- **A preview of today's numbers** underneath it — the same eleven questions,
  answered, in the same order they appear in the spreadsheet.
- **A full analysis terminal** further down the page: market breadth, a sector
  rotation heat map across six time windows, every indicator for every ticker in
  one sortable grid, a feed of the signals that fired today, and the headlines
  behind the biggest moves. Click any ticker for its complete read-out.
- **A banner at the top** that downloads `stock_analysis.py`, the single file
  that builds all of it, so anyone can run the same thing on their own machine.

The numbers refresh automatically three times each weekday.

---

## The spreadsheet

`Stock_Analysis_MM-DD.xlsx` contains:

| Tab | What is on it |
| --- | --- |
| **Dashboard** | The whole watchlist on one screen — advancers, decliners, breadth, sector ranking with a chart, strongest and weakest names, biggest gainers and losers. |
| **MM-DD** | The daily sheet. One tab per session, five sessions deep. |
| **Sector Rotation** | The eleven sector funds across six time windows, colour-graded. |
| **Signals** | Only the tickers where something actually happened today. |
| **How To Read This** | Every column explained in plain English. |

### The daily sheet

Columns A to M are the original layout, unchanged, down to the row numbers —
sectors on rows 7 to 17, the index funds on 19 to 22, the watchlist from 24 down:

| Column | Question |
| --- | --- |
| `%change` | The day's move, close against yesterday's close |
| `Prev Strength` | Rank of that move inside its group, 1 being strongest |
| `Price > Pre-Day` | Closed above yesterday? |
| `Yest Volum > 2Pre-Day` | More shares traded than the day before? |
| `MACD (Green > Red)` | MACD line above its signal line? |
| `Price > 8EMA` | Above the 8-day exponential average? |
| `> 21EMA` | Above the 21-day exponential average? |
| `Price > 50DMA` | Above the 50-day simple average? |
| `Price > Upper BB` | Above the top Bollinger Band? |
| `Price > Lower BB` | Above the bottom Bollinger Band? |
| `Price > Mid BB` | Above the middle Bollinger Band? |

Each answer is `Yes`, `No`, or `Same` when the two numbers are within a tenth of
a percent of each other and calling it either way would be dishonest. The cells
colour themselves — green, red, amber — and they recolour if you type over them.

From column N rightwards, everything is new:

- **Bull Score** — nine checks, one point each, drawn as a bar in the cell
- **Signal** — the score in words, from Very Strong to Very Weak
- **Trend** — where the price sits against all four averages at once
- **Full OHLCV** — open, high, low, volume, and volume against its 20-day average
- **Momentum** — RSI, Stochastic %K, ADX, and the three raw MACD numbers
- **Every moving average** — 8 and 21 EMA, 50, 100 and 200 DMA
- **Bollinger detail** — upper, middle and lower bands, %B, and band width
- **Range and volatility** — ATR, 52-week high and low, distance from the high,
  position inside the year's range
- **Returns** — five-day, one-month, three-month and year-to-date
- **Top News 1 to 3** — three clickable headlines per ticker

---

## Running it yourself

```bash
pip install -r requirements.txt
python stock_analysis.py
```

That writes `Stock_Analysis_MM-DD.xlsx` into the current folder. It takes about
five seconds.

### Options

```bash
python stock_analysis.py --history 10        # ten sessions as tabs instead of five
python stock_analysis.py --no-news           # skip headline fetching, runs faster
python stock_analysis.py --out report.xlsx   # choose the filename
python stock_analysis.py --json data.json    # also write the JSON the website reads
python stock_analysis.py --watchlist mine.json
```

### Changing the watchlist

Edit `watchlist.json`. Order is preserved in the spreadsheet, so the tickers
appear exactly where you put them.

```json
{
  "groups": [
    { "key": "sectors", "title": "SECTOR SPDRs", "numbered": true,
      "tickers": ["XLI", "XLV", "XLF"] }
  ]
}
```

The `settings` block in the same file controls the indicator periods — the EMA
and DMA lengths, the Bollinger settings, the MACD periods, and how close two
numbers have to be before the answer becomes `Same`.

---

## Where the data comes from

Yahoo Finance, through the free [`yfinance`](https://pypi.org/project/yfinance/)
package. No account, no API key, no subscription, no rate limit worth worrying
about at forty tickers. Headlines come from Yahoo Finance's public RSS feeds.

Indicator settings match the chart layout they were read off:

```
EMA 8, 21          DMA 50, 100, 200      Bollinger 20, 2
MACD 12, 26, 9     Volume MA 20          RSI 14      ATR 14
```

---

## How the automation works

`.github/workflows/daily.yml` runs the script three times each weekday, commits
the rebuilt `latest.xlsx` and `latest.json` into `docs/data/`, and GitHub Pages
serves the result. Nothing runs on anyone's computer, and there is nothing to
keep switched on.

To publish the site: **Settings → Pages → Source: Deploy from a branch →
`main` / `docs`**.

To rebuild by hand: **Actions → Build the daily workbook → Run workflow**.

---

## Not investment advice

This is a charting and bookkeeping tool. Every number in it is arithmetic on
past prices from a free public feed that may be delayed, adjusted or simply
wrong. Nothing here is a recommendation to buy or sell anything. Check prices
with a broker before acting on them.

## Licence

MIT. See [LICENSE](LICENSE).
