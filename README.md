# SFDA Drug Scraper

A Python web scraper that extracts comprehensive drug information from the Saudi FDA (SFDA) drug list at https://www.sfda.gov.sa/ar/drugs-list

## Features

- **Multi-page scraping**: Supports scraping multiple pages from the SFDA website
- **Detailed extraction**: Clicks into each drug entry to extract comprehensive details
- **Excel export**: Automatically saves all extracted data to an Excel file with formatting
- **Comprehensive data**: Captures 27 fields including:
  - Trade Name & Scientific Name
  - Registration Number & Status
  - Concentration & Pharmaceutical Form
  - Route of Administration
  - Package Information
  - Manufacturer & Marketing Details
  - Price & Marketing Status
  - And more...

## Requirements

- Python 3.7+
- Selenium
- openpyxl
- webdriver-manager
- Chrome browser

## Installation

1. Clone the repository:
```bash
git clone https://github.com/MrSAYKA/webscrapper.git
cd webscrapper
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Usage

Run the scraper:
```bash
python webscrapper.py
```

### Configuration

To change the number of pages to scrape, modify line 73 in `webscrapper.py`:

Current (5 pages for testing):
```python
for page in range(1, 6):  # Pages 1-5
```

For all 873 pages:
```python
for page in range(1, 874):  # Pages 1-873
```

## Output

The script generates `sfda_drugs_test.xlsx` with:
- Formatted headers with blue background
- Properly sized columns
- All extracted drug information

## Features

- **Automatic waits**: Includes 20-second initial page load wait and 5-10 second waits between actions
- **Error handling**: Gracefully skips rows without details links
- **Progress tracking**: Displays real-time progress with drug names
- **Batch saving**: Saves Excel file after each page
- **Detailed logging**: Shows errors and skipped rows for debugging

## How It Works

1. **Navigate to page**: Opens the SFDA drugs list page
2. **Extract row data**: Finds all drug rows in the table
3. **Click details**: Clicks the "التفاصيل" (Details) link for each drug
4. **Extract info**: Parses all fields from the detail page
5. **Map to Excel**: Maps Arabic labels to English column headers
6. **Go back**: Returns to the list to process next drug
7. **Save**: Automatically saves to Excel file

## Notes

- The website is in Arabic; field labels are mapped to English headers
- The scraper waits adequately for page loads to ensure data accuracy
- Each drug detail page takes 5 seconds to load
- Returning to the list takes 10 seconds per drug

## License

MIT License - Feel free to use and modify as needed

## Author

MrSAYKA

---

**Note**: This scraper is for educational purposes. Please respect the SFDA website's terms of service and robots.txt file.

