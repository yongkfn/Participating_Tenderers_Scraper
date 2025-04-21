from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import pandas as pd
import time
import logging
import pytesseract
from PIL import Image
import platform

# Configure Tesseract path for Windows
if platform.system() == 'Windows':
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'
import io
import re
from datetime import datetime
import sys
import os
from typing import List, Tuple, Optional

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('ura_scraper.log'),
        logging.StreamHandler()
    ]
)

class XPaths:
    """XPath constants for URA website elements"""
    # Main search elements
    SEARCH_BOX = '//*[@id="us-s-txt"]'
    
    # Search results - Specific to your website structure
    SEARCH_RESULTS_CONTAINER = "//div[contains(@class, 'search-results-container')]"
    SEARCH_RESULT_ITEMS = "//div[contains(@class, 'search-results-container')]/div/a/div[2]"
    
    # Project details popup
    CLOSE_BUTTON = "//a[contains(@class, 'close-popup')] | //button[contains(@class, 'close')] | //img[@alt='Close']"
    
    # Tender results
    TENDER_RESULTS_LINK = "//a[contains(text(), 'Tender Results') or contains(text(), 'Award')]"
    SUCCESSFUL_TENDERER = "//strong[contains(text(), 'SUCCESSFUL TENDERER')]/following-sibling::text()"
    
    # PDF-specific elements
    PDF_FRAME = "//iframe[contains(@src, '.pdf')] | //embed[contains(@src, '.pdf')]"
    PDF_PAGE = "//canvas | //div[contains(@class, 'textLayer')]"


def check_dependencies():
    """Check if all required dependencies are installed"""
    missing_deps = []
    
    # Check Python packages
    try:
        import selenium
    except ImportError:
        missing_deps.append("selenium")
    
    try:
        import pandas
    except ImportError:
        missing_deps.append("pandas")
    
    try:
        import openpyxl
    except ImportError:
        missing_deps.append("openpyxl")
    
    try:
        import pytesseract
        # Check if Tesseract OCR is installed
        try:
            pytesseract.get_tesseract_version()
        except pytesseract.TesseractNotFoundError:
            missing_deps.append("Tesseract OCR (executable)")
    except ImportError:
        missing_deps.append("pytesseract")
    
    try:
        from PIL import Image
    except ImportError:
        missing_deps.append("Pillow")
    
    if missing_deps:
        logging.error("Missing dependencies: " + ", ".join(missing_deps))
        logging.error("Please install them using: pip install -r requirements.txt")
        logging.error("For Tesseract OCR, please visit: https://github.com/UB-Mannheim/tesseract/wiki")
        sys.exit(1)
    else:
        logging.info("All dependencies are properly installed")


def format_date_for_search(date_input) -> str:
    """Convert date from Excel format to expected search format"""
    if isinstance(date_input, datetime):
        # If date is already a datetime object
        return date_input.strftime("%d-%b-%y").lstrip("0")
    elif isinstance(date_input, str):
        try:
            # If date is a string, try to parse it
            if " " in date_input:  # For format like "2025-03-13 00:00:00"
                date_obj = datetime.strptime(date_input.split(" ")[0], "%Y-%m-%d")
            else:  # For format like "2025-03-13"
                date_obj = datetime.strptime(date_input, "%Y-%m-%d")
            return date_obj.strftime("%d-%b-%y").lstrip("0")
        except ValueError:
            # If above parsing fails, try other common formats
            try:
                date_obj = datetime.strptime(date_input, "%d-%b-%y")
                return date_input  # Already in correct format
            except:
                logging.warning(f"Unable to parse date: {date_input}")
                return date_input
    return str(date_input)


def setup_driver() -> webdriver.Chrome:
    """Set up Chrome driver with options"""
    from webdriver_manager.chrome import ChromeDriverManager
    from selenium.webdriver.chrome.service import Service
    
    chrome_options = Options()
    chrome_options.add_argument('--start-maximized')
    chrome_options.add_argument('--disable-notifications')
    chrome_options.add_argument('--disable-popup-blocking')
    
    # Add prefs for PDF handling
    chrome_options.add_experimental_option('prefs', {
        "download.default_directory": os.getcwd(),
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": False,
        "plugins.plugins_disabled": [],
    })
    
    try:
        # Use ChromeDriverManager to automatically download and manage the driver
        service = Service(ChromeDriverManager().install())
        driver = webdriver.Chrome(service=service, options=chrome_options)
    except Exception as e:
        logging.error(f"Error setting up Chrome driver: {str(e)}")
        logging.error("Please make sure Chrome browser is installed")
        sys.exit(1)
    
    driver.implicitly_wait(10)
    return driver


def extract_date_from_text(text: str) -> Optional[str]:
    """Extract date from text in various formats"""
    date_patterns = [
        r'DATE OF AWARD[^\d]*(\d{1,2})[-\s]+([A-Za-z]+)[-\s]+(\d{4})',
        r'(\d{1,2})[-\s]+([A-Za-z]+)[-\s]+(\d{4})',
        r'(\d{1,2})-([A-Za-z]{3})-(\d{2})'
    ]
    
    for pattern in date_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            try:
                day = match.group(1)
                month = match.group(2)
                year = match.group(3)
                
                # Handle 2-digit year
                if len(year) == 2:
                    year = f"20{year}"
                
                # Parse date
                date_str = f"{day} {month} {year}"
                parsed_date = datetime.strptime(date_str, "%d %B %Y") if len(month) > 3 else datetime.strptime(date_str, "%d %b %Y")
                
                # Return in format matching CSV: "7-Aug-24"
                return parsed_date.strftime("%d-%b-%y").lstrip("0")
            except Exception as e:
                logging.warning(f"Error parsing date: {str(e)}")
                continue
    return None


def verify_project_date(driver: webdriver.Chrome, expected_date: str) -> bool:
    """Verify that the project's award date matches what we expect"""
    try:
        # Wait for popup to be visible (various possible selectors)
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".popup-content, .info-window, .details-panel"))
        )
        
        # Get all text from the popup
        popup_text = driver.find_element(By.CSS_SELECTOR, ".popup-content, .info-window, .details-panel").text
        
        # Extract date from the text
        found_date = extract_date_from_text(popup_text)
        
        if found_date:
            logging.info(f"Found date: {found_date}, Expected: {expected_date}")
            # Compare with expected date
            return found_date == expected_date
            
    except Exception as e:
        logging.error(f"Error verifying project date: {str(e)}")
    
    return False


def search_for_project(driver: webdriver.Chrome, project_name: str, expected_date: str) -> bool:
    """Search for a project on the URA website with sequential search result checking"""
    try:
        # Navigate to URA GLS website
        driver.get("https://eservice.ura.gov.sg/maps/?service=GLSRELEASE")
        time.sleep(3)  # Wait for page to load completely
        
        # Find and clear search box
        search_box = WebDriverWait(driver, 15).until(
            EC.element_to_be_clickable((By.XPATH, XPaths.SEARCH_BOX))
        )
        search_box.clear()
        search_box.send_keys(project_name)
        search_box.send_keys(Keys.RETURN)
        time.sleep(3)  # Wait for search results to load
        
        # Find all search result items
        try:
            # Wait for search results to appear
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, XPaths.SEARCH_RESULTS_CONTAINER))
            )
            
            # Get all search result items
            result_items = driver.find_elements(By.XPATH, XPaths.SEARCH_RESULT_ITEMS)
            logging.info(f"Found {len(result_items)} search results for '{project_name}'")
            
            if not result_items:
                logging.warning(f"No search results found for '{project_name}'")
                return False
            
            # Try each search result one by one
            for i, result in enumerate(result_items):
                try:
                    logging.info(f"Trying search result {i+1} of {len(result_items)}: {result.text}")
                    
                    # Click on the search result
                    driver.execute_script("arguments[0].click();", result)
                    time.sleep(2)  # Wait for details to load
                    
                    # Verify if the date matches
                    if verify_project_date(driver, expected_date):
                        logging.info(f"Found matching project: '{project_name}' with date '{expected_date}'")
                        return True
                    
                    # Date doesn't match, go back and try next result
                    logging.info(f"Date didn't match for result {i+1}, trying next...")
                    
                    # Close the popup
                    try:
                        close_buttons = driver.find_elements(By.XPATH, XPaths.CLOSE_BUTTON)
                        if close_buttons:
                            driver.execute_script("arguments[0].click();", close_buttons[0])
                        else:
                            # Try pressing Escape key
                            webdriver.ActionChains(driver).send_keys(Keys.ESCAPE).perform()
                    except Exception as e:
                        logging.warning(f"Error closing popup: {str(e)}")
                        # Try to recover by going back to search results
                        driver.execute_script("history.back();")
                    
                    time.sleep(2)  # Wait before trying next result
                
                except Exception as e:
                    logging.error(f"Error processing search result {i+1}: {str(e)}")
                    # Try to continue with next result
                    continue
            
            logging.warning(f"None of the search results matched '{project_name}' with date '{expected_date}'")
            return False
            
        except Exception as e:
            logging.error(f"Error finding search results: {str(e)}")
            return False
        
    except Exception as e:
        logging.error(f"Error in search_for_project: {str(e)}")
        return False


def extract_text_from_pdf(driver: webdriver.Chrome, pdf_url: str) -> str:
    """Extract text from PDF using various methods"""
    try:
        # Open the PDF in a new tab
        original_window = driver.current_window_handle
        driver.execute_script(f"window.open('{pdf_url}', '_blank');")
        
        # Switch to new tab
        WebDriverWait(driver, 10).until(EC.number_of_windows_to_be(2))
        for window_handle in driver.window_handles:
            if window_handle != original_window:
                driver.switch_to.window(window_handle)
                break
        
        time.sleep(5)  # Wait for PDF to load
        
        # Try multiple methods to extract text
        text = ""
        
        # Method 1: Try to get text directly from PDF viewer
        try:
            pdf_frame = driver.find_element(By.XPATH, XPaths.PDF_FRAME)
            driver.switch_to.frame(pdf_frame)
            text_elements = driver.find_elements(By.XPATH, "//span[contains(@class, 'text')]")
            text = "\n".join([elem.text for elem in text_elements])
            driver.switch_to.default_content()
        except:
            pass
        
        # Method 2: Use OCR if direct text extraction fails
        if not text:
            try:
                # Check if Tesseract is available
                try:
                    pytesseract.get_tesseract_version()
                except pytesseract.TesseractNotFoundError:
                    logging.error("Tesseract is not installed or not found in PATH")
                    logging.error("Please install Tesseract OCR from https://github.com/UB-Mannheim/tesseract/wiki")
                    return ""
                
                # Take screenshot of the PDF
                screenshot = driver.get_screenshot_as_png()
                image = Image.open(io.BytesIO(screenshot))
                
                # Use OCR
                text = pytesseract.image_to_string(image)
            except Exception as e:
                logging.error(f"OCR failed: {str(e)}")
        
        # Close PDF tab and switch back
        driver.close()
        driver.switch_to.window(original_window)
        
        return text
        
    except Exception as e:
        logging.error(f"Error extracting PDF text: {str(e)}")
        # Make sure we switch back to original window
        try:
            driver.switch_to.window(original_window)
        except:
            pass
        return ""


def extract_tenderers_from_pdf_text(pdf_text: str) -> Tuple[str, List[str]]:
    """Extract tenderer information from PDF text"""
    successful_tenderer = ""
    other_tenderers = []
    
    lines = pdf_text.split('\n')
    found_table = False
    
    for i, line in enumerate(lines):
        # Look for table headers or successful tenderer section
        if "SUCCESSFUL TENDERER" in line.upper():
            # Extract successful tenderer from following lines
            if i + 1 < len(lines):
                successful_tenderer = lines[i + 1].strip()
        
        # Look for ranking table
        if "RANKING" in line.upper() and "NAME OF TENDERER" in line.upper():
            found_table = True
            continue
        
        if found_table:
            # Match ranking (1, 2, 3, etc.) followed by company name
            match = re.match(r'^\s*(\d+)\s+(.+)$', line.strip())
            if match:
                ranking = int(match.group(1))
                company_name = match.group(2).strip()
                
                if ranking == 1:
                    successful_tenderer = company_name
                else:
                    other_tenderers.append(company_name)
            elif not line.strip():
                # Empty line might indicate end of table
                found_table = False
    
    return successful_tenderer, other_tenderers


def extract_tenderers(driver: webdriver.Chrome) -> Tuple[str, List[str]]:
    """Extract tenderer information from the project page"""
    try:
        # Find tender results link
        try:
            tender_results_link = WebDriverWait(driver, 10).until(
                EC.element_to_be_clickable((By.XPATH, XPaths.TENDER_RESULTS_LINK))
            )
            pdf_url = tender_results_link.get_attribute('href')
            
            # Extract text from PDF
            pdf_text = extract_text_from_pdf(driver, pdf_url)
            
            # Parse tenderer information
            successful_tenderer, other_tenderers = extract_tenderers_from_pdf_text(pdf_text)
        except Exception as e:
            logging.error(f"Error finding tender results link: {str(e)}")
            successful_tenderer, other_tenderers = "", []
        
        # If we didn't get successful tenderer from PDF, try to get it from the page
        if not successful_tenderer:
            try:
                successful_element = driver.find_element(By.XPATH, XPaths.SUCCESSFUL_TENDERER)
                successful_tenderer = successful_element.text.strip()
            except:
                pass
        
        return successful_tenderer, other_tenderers
        
    except Exception as e:
        logging.error(f"Error extracting tenderers: {str(e)}")
        return "", []


def process_project(driver: webdriver.Chrome, project_info: dict) -> dict:
    """Process a single project and extract tenderer information"""
    project_name = project_info.get('Location', '')
    raw_date = project_info.get('Date of Award', '')
    num_bids = project_info.get('Number of Bids', 0)
    
    # Convert date to expected format
    expected_date = format_date_for_search(raw_date)
    
    logging.info(f"Processing project: {project_name} (Original date: {raw_date}, Formatted date: {expected_date}, {num_bids} bids)")
    
    if search_for_project(driver, project_name, expected_date):
        successful_tenderer, other_tenderers = extract_tenderers(driver)
        
        # Update project info
        project_info['Name of Successful Tenderer'] = successful_tenderer
        
        # Fill in other tenderers
        for i, tenderer in enumerate(other_tenderers[:4]):  # Limit to 4 other tenderers
            column_name = f'Name of Other Participating Tenderer - {i+1}'
            project_info[column_name] = tenderer
        
        logging.info(f"Successfully extracted {len(other_tenderers) + 1} tenderers for {project_name}")
        
        # Close project popup
        try:
            close_button = driver.find_element(By.XPATH, XPaths.CLOSE_BUTTON)
            close_button.click()
            time.sleep(1)
        except:
            pass
    
    return project_info


def main():
    """Main function to process all projects needing tenderer information"""
    # Check all dependencies first
    check_dependencies()
    
    try:
        # Load Excel file
        excel_file = r'C:\Users\nikki\Documents\GitHub\Participating_Tenderers_Scraper\data\ParticipatingTenderers.xlsx'  # Use raw string to handle backslashes
        if not os.path.exists(excel_file):
            logging.error(f"Excel file not found: {excel_file}")
            logging.error("Please ensure the Excel file exists or update the path")
            sys.exit(1)
        
        # Use read_excel for xlsx files
        df = pd.read_excel(excel_file)
        
        # Identify projects needing tenderer information
        projects_to_process = []
        for _, row in df.iterrows():
            num_bids = row.get('Number of Bids', 0)
            if num_bids > 1 and pd.isna(row.get('Name of Other Participating Tenderer - 1', None)):
                projects_to_process.append(row.to_dict())
        
        logging.info(f"Found {len(projects_to_process)} projects needing tenderer information")
        
        # Set up driver
        driver = setup_driver()
        
        # Process each project
        updated_projects = []
        for project in projects_to_process:
            updated_project = process_project(driver, project)
            updated_projects.append(updated_project)
            time.sleep(2)  # Add delay between projects to avoid overloading the server
        
        # Update DataFrame
        for updated_project in updated_projects:
            # Find and update the row in the original DataFrame
            location = updated_project['Location']
            date = updated_project['Date of Award']
            mask = (df['Location'] == location) & (df['Date of Award'] == date)
            
            if mask.any():
                for column, value in updated_project.items():
                    df.loc[mask, column] = value
        
        # Save updated Excel file
        output_file = r'"C:\Users\nikki\Documents\GitHub\Participating_Tenderers_Scraper\ura_sites_interim_results.xlsx"'
        df.to_excel(output_file, index=False)
        logging.info(f"Saved updated data to {output_file}")
        
        driver.quit()
        
    except Exception as e:
        logging.error(f"Error in main function: {str(e)}")
        if 'driver' in locals():
            driver.quit()


if __name__ == "__main__":
    main()