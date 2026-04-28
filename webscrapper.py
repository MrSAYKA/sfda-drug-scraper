from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
import time
import re

def extract_detail_page_info(driver, wait):
    """Extract all information from a drug detail page"""
    try:
        # Wait for the detail page to load
        wait.until(EC.presence_of_all_elements_located((By.TAG_NAME, "tr")))
        time.sleep(5)
        
        # Find all rows with data
        rows = driver.find_elements(By.TAG_NAME, "tr")
        detail_data = {}
        
        for row in rows:
            cells = row.find_elements(By.TAG_NAME, "td")
            if len(cells) == 2:
                label = cells[0].text.strip()  # Arabic label (right side)
                value = cells[1].text.strip()  # Value (left side)
                if label:
                    detail_data[label] = value
        
        return detail_data
    except Exception as e:
        print(f"  Error extracting detail: {str(e)[:80]}")
        return {}

def scrape_sfda_5_pages():

    print("=" * 60)
    print("SFDA DRUG SCRAPER - TEST (5 Pages)")
    print("=" * 60)

    print("\nSetting up ChromeDriver...")
    options = Options()
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("--start-maximized")
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    options.add_experimental_option("useAutomationExtension", False)

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 20)
    print("✓ Ready!\n")

    # Create Excel
    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = "Drugs"

    # Headers - comprehensive list of possible fields
    headers = ["Trade Name", "Scientific Name", "Registration Number", "Concentration", 
               "Concentration Unit", "Pharmaceutical Form", "Route of Administration",
               "Size", "Size Unit", "Package Type", "Package Size", "Dispensing Method",
               "Monitoring", "Validity Period", "Storage Conditions", "Manufacturing Company",
               "Country of Manufacture", "Marketing Company", "Country of Marketing Company",
               "First Agent", "Second Agent", "Third Agent", "Classification", "Price",
               "Registration Status", "Marketing Status", "Report", "Pharmacological Code"]
    sheet.append(headers)

    # Style headers
    header_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF")
    for cell in sheet[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Set column widths
    for col_num, header in enumerate(headers, 1):
        sheet.column_dimensions[openpyxl.utils.get_column_letter(col_num)].width = 25

    website = "https://www.sfda.gov.sa/ar/drugs-list"
    total_drugs = 0
    errors = 0

    for page in range(1, 6):  # ← Change to range(1, 874) for full 873 pages

        print(f"\nPAGE {page}/5")
        print("-" * 40)

        try:
            url = f"{website}?pg={page}"
            driver.get(url)

            wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody tr")))
            time.sleep(20)

            rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
            print(f"Found {len(rows)} rows")
            
            row_index = 0
            while row_index < len(rows):
                try:
                    # Navigate back to the current page to reset rows
                    driver.get(f"{website}?pg={page}")
                    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "table tbody tr")))
                    time.sleep(10)
                    
                    # Re-fetch rows to avoid stale element references
                    rows = driver.find_elements(By.CSS_SELECTOR, "table tbody tr")
                    
                    if row_index >= len(rows):
                        break
                    
                    row = rows[row_index]
                    cells = row.find_elements(By.TAG_NAME, "td")
                    
                    if len(cells) >= 1:
                        try:
                            # Try to find the "التفاصيل" (Details) link - could be <a> or other element
                            details_link = None
                            
                            # Try finding by tag <a>
                            links = row.find_elements(By.TAG_NAME, "a")
                            if links:
                                details_link = links[0]
                            
                            # If no link found, try finding button
                            if not details_link:
                                buttons = row.find_elements(By.TAG_NAME, "button")
                                if buttons:
                                    details_link = buttons[0]
                            
                            if details_link:
                                # Click the details link
                                details_link.click()
                                time.sleep(5)
                                
                                # Extract all information from detail page
                                detail_info = extract_detail_page_info(driver, wait)
                                
                                # Map Arabic labels to English headers
                                row_data = [
                                    detail_info.get("الاسم التجاري", ""),           # Trade Name
                                    detail_info.get("الاسم العلمي", ""),           # Scientific Name
                                    detail_info.get("رقم التسجيل", ""),            # Registration Number
                                    detail_info.get("التركيز", ""),                # Concentration
                                    detail_info.get("وحدة التركيز", ""),           # Concentration Unit
                                    detail_info.get("الشكل الصيدلاني", ""),        # Pharmaceutical Form
                                    detail_info.get("طريقة الإستعمال", ""),        # Route of Administration
                                    detail_info.get("الحجم", ""),                  # Size
                                    detail_info.get("وحدة الحجم", ""),             # Size Unit
                                    detail_info.get("نوع العبوة", ""),              # Package Type
                                    detail_info.get("حجم العبوة", ""),              # Package Size
                                    detail_info.get("طريقة الصرف", ""),             # Dispensing Method
                                    detail_info.get("المراقبة", ""),               # Monitoring
                                    detail_info.get("مدة الصلاحية", ""),            # Validity Period
                                    detail_info.get("ظروف التخزين", ""),           # Storage Conditions
                                    detail_info.get("الشركة المصنعة", ""),         # Manufacturing Company
                                    detail_info.get("بلد التصنيع", ""),            # Country of Manufacture
                                    detail_info.get("الشركة المسوقة", ""),         # Marketing Company
                                    detail_info.get("بلد الشركة المسوقة", ""),     # Country of Marketing Company
                                    detail_info.get("الوكيل الأول", ""),           # First Agent
                                    detail_info.get("الوكيل الثاني", ""),          # Second Agent
                                    detail_info.get("الوكيل الثالث", ""),          # Third Agent
                                    detail_info.get("H05AA02", ""),               # Classification
                                    detail_info.get("السعر", ""),                  # Price
                                    detail_info.get("حالة التسجيل", ""),           # Registration Status
                                    detail_info.get("حالة التسويق", ""),           # Marketing Status
                                    detail_info.get("التقرير", ""),                # Report
                                    detail_info.get("رمز التكييف الدوائية", ""),   # Pharmacological Code
                                ]
                                
                                sheet.append(row_data)
                                total_drugs += 1
                                trade_name = row_data[0][:40] if row_data[0] else "Unknown"
                                print(f"  ✓ [{total_drugs}] {trade_name}")
                            else:
                                print(f"  ⊘ Row {row_index}: No details link found - skipping")
                        
                        except Exception as link_error:
                            print(f"  ⊘ Row {row_index}: {str(link_error)[:60]} - skipping")
                        
                        row_index += 1

                except Exception as e:
                    print(f"  ✗ Row {row_index} error: {str(e)[:80]}")
                    errors += 1
                    row_index += 1

            workbook.save('sfda_drugs_test.xlsx')
            print(f"✓ Page {page} saved!")

        except Exception as e:
            print(f"✗ Page {page} error: {str(e)[:80]}")
            errors += 1

    driver.quit()

    print("\n" + "=" * 60)
    print("DONE!")
    print(f"Total drugs scraped : {total_drugs}")
    print(f"Errors              : {errors}")
    print("File saved          : sfda_drugs_test.xlsx")
    print("=" * 60)

if __name__ == "__main__":
    scrape_sfda_5_pages()