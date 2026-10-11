# Price scraper to Excel

A Python script that collects a product catalogue from a website and saves it as a formatted Excel report. Built as a portfolio project.

The demo target is [books.toscrape.com](https://books.toscrape.com), a sandbox site made for scraping practice.

## Features

- Walks through catalogue pages automatically, following the "next" link
- Collects title, price, rating, availability and link for every product
- Saves an Excel file with two sheets: the full product list (with filters and a frozen header) and a summary by rating
- Polite scraping: a pause between requests and a clear error if the site does not respond

## Tech

Python, requests, BeautifulSoup, openpyxl.

## Run

```
pip install requests beautifulsoup4 openpyxl
python scraper.py        # first 5 pages
python scraper.py 50     # all 50 pages
```

The result is saved as `books.xlsx` next to the script.
