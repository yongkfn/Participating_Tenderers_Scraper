// URA Vacant Sites Web Scraper
const puppeteer = require('puppeteer');
const XLSX = require('xlsx');
const fs = require('fs');

// Helper function to format company names consistently
function formatCompanyName(name) {
  if (!name) return '';
  
  // Convert to string in case it's not
  const companyName = String(name);
  
  // Remove excess whitespace
  let formatted = companyName.replace(/\s+/g, ' ').trim();
  
  // Fix common abbreviations spacing
  formatted = formatted.replace(/(\w+)(\s+)(Pte)(\s*)(Ltd)/gi, '$1 $3. $5.');
  formatted = formatted.replace(/(\w+)(\s+)(Pte)(\s*)(\.?)(\s*)(Ltd)/gi, '$1 $3.$6$7');
  
  // Fix spacing after commas
  formatted = formatted.replace(/,(\S)/g, ', $1');
  
  // Remove multiple periods
  formatted = formatted.replace(/\.+/g, '.');
  
  return formatted;
}

// Main function to run the scraper
async function runScraper(options = {}) {
  const {
    headless = false,
    debug = false,
    startIndex = 0,
    endIndex = Infinity,
    retries = 3
  } = options;
  console.log('Starting URA vacant sites scraper...');
  
  // Read the Excel file
  const workbook = XLSX.readFile('./data/ura-vacant-sites_CNQC.xlsx');
  
  // Get locations from the sheet (appears to be just one sheet now)
  const sheetName = workbook.SheetNames[0]; // "CNQC Part 2024"
  const worksheet = workbook.Sheets[sheetName];
  const sheetData = XLSX.utils.sheet_to_json(worksheet);
  
  console.log(`Using data from sheet: ${sheetName}`);
  console.log(`Found ${sheetData.length} entries in the Excel file`);
  
  // Extract locations to search
  const locations = sheetData.map(row => row.Location);
  console.log(`Found ${locations.length} locations to search:`, locations);
  
  // Create a deep copy of the sheet data for our output
  const outputData = JSON.parse(JSON.stringify(sheetData));
  
  // Launch browser
  const browser = await puppeteer.launch({
    headless: headless,
    defaultViewport: null,
    args: ['--start-maximized', '--disable-web-security', '--no-sandbox']
  });
  
  // Create a new page
  const page = await browser.newPage();
  
  // Set timeout to 60 seconds
  page.setDefaultNavigationTimeout(60000);
  
  try {
    // Filter locations based on start and end indices
    const locationsToProcess = locations.slice(
      Math.min(startIndex, locations.length),
      Math.min(endIndex, locations.length)
    );
    
    console.log(`Processing ${locationsToProcess.length} locations (from index ${startIndex} to ${Math.min(endIndex, locations.length) - 1})`);
    
    // Process each location
    for (let i = 0; i < locationsToProcess.length; i++) {
      const location = locationsToProcess[i];
      const originalIndex = startIndex + i;
      console.log(`Searching for location (${i+1}/${locationsToProcess.length}, original index: ${originalIndex}): ${location}`);
      
      let success = false;
      let attemptCount = 0;
      
      // Retry loop
      while (!success && attemptCount < retries) {
        attemptCount++;
        if (attemptCount > 1) {
          console.log(`Retry attempt ${attemptCount} for location: ${location}`);
        }
      
        try {
          // Navigate to URA site
          await page.goto('https://eservice.ura.gov.sg/maps/?service=GLSRELEASE', {
            waitUntil: 'networkidle2',
            timeout: 60000
          });
          
          // Take screenshot if in debug mode
          if (debug) {
            await takeScreenshot(page, `${originalIndex}_initial_page`);
            await savePageHtml(page, `${originalIndex}_initial_page`);
          }
          

          // Wait for the page to load completely
        await page.evaluate(() => new Promise(resolve => setTimeout(resolve, 5000))); // Give it more time to fully load

        console.log('Using XPath to find search input field...');

        try {
          // Use the provided XPath to find the search input element
          const searchInputXPath = '//*[@id="us-s-txt"]';
          
          // Wait for the element to be available
          await page.waitForXPath(searchInputXPath, { timeout: 10000 });
          
          // Get the element
          const [searchInput] = await page.$x(searchInputXPath);
          
          if (searchInput) {
            console.log('Search input found, clicking and typing...');
            
            // Click the input to focus it
            await searchInput.click();
            
            // Clear any existing text
            await searchInput.evaluate(el => el.value = '');
            
            // Type the location
            await searchInput.type(location, { delay: 100 }); // Slight delay between keypresses
            
            console.log(`Typed "${location}" into search field, pressing Enter...`);
            
            // Press Enter to submit the search
            await page.keyboard.press('Enter');
            
            // Wait for search results to load
            await page.evaluate(() => new Promise(resolve => setTimeout(resolve, 5000)));
            
            console.log('Search submitted, waiting for results...');
          } else {
            console.error('Search input element found by XPath but could not be accessed');
            throw new Error('Could not access search input element');
          }
        } catch (error) {
          console.error(`Error using search input: ${error.message}`);
          
          // Fallback to URL-based search if XPath fails
          console.log('Falling back to URL-based search...');
          await page.evaluate((searchLocation) => {
            window.location.href = `https://eservice.ura.gov.sg/maps/?service=GLSRELEASE&site=1045&search=${encodeURIComponent(searchLocation)}`;
          }, location);
          await page.evaluate(() => new Promise(resolve => setTimeout(resolve, 5000)));
        }

// Take a screenshot after search
await page.screenshot({ path: `screenshot_after_search_${originalIndex}.png` });
        
        // Wait for search results to appear (might be in a dropdown, list, or other UI element)
        await page.waitForTimeout(3000); // Wait for search results
        
        // Look for search results using various possible selectors
        const resultSelectors = [
          '.search-result-item', 
          '.search-results li', 
          '.result-item',
          'div[role="listitem"]',
          '[data-search-result]',
          // Add more selectors if needed
        ];
        
        let resultFound = false;
        for (const selector of resultSelectors) {
          const results = await page.$(selector);
          if (results.length > 0) {
            // Click the first result
            await results[0].click();
            resultFound = true;
            console.log(`Clicked search result using selector: ${selector}`);
            break;
          }
        }
        
        if (!resultFound) {
          console.log(`No search results found for location: ${location}`);
          continue; // Skip to the next location
        }
        
          // Wait for the project details to load
          console.log('Waiting for project details to load...');
          await page.evaluate(() => new Promise(resolve => setTimeout(resolve, 5000)));
          
          if (debug) {
            await takeScreenshot(page, `${originalIndex}_project_details_page`);
            await savePageHtml(page, `${originalIndex}_project_details_page`);
          }
          
          console.log('Looking for "Tender Results" tab...');
          
          // From the screenshot, we need to click on "Tender Results" in the left panel
          // First, gather information about all potential "Tender Results" elements
          const tenderResultsElements = await page.evaluate(() => {
            const TEXT = 'Tender Results';
            
            // Get all elements that might be the tender results link/button
            const allElements = Array.from(document.querySelectorAll('a, button, div, span, li, td'));
            
            // Find elements containing exactly "Tender Results" text
            const exactMatches = allElements.filter(el => {
              const text = el.textContent.trim();
              return text === TEXT;
            });
            
            // Find elements containing "Tender Results" as part of their text
            const partialMatches = allElements.filter(el => {
              const text = el.textContent.trim();
              return text !== TEXT && text.includes(TEXT);
            });
            
            // Find elements with similar classes or ids
            const attributeMatches = allElements.filter(el => {
              const id = el.id || '';
              const className = el.className || '';
              
              return (id.toLowerCase().includes('tender') || 
                     id.toLowerCase().includes('result') ||
                     className.toLowerCase().includes('tender') ||
                     className.toLowerCase().includes('result')) && 
                     !exactMatches.includes(el) && 
                     !partialMatches.includes(el);
            });
            
            // Capture element details for debugging
            const captureElementDetails = (elements) => {
              return elements.map(el => ({
                tag: el.tagName,
                id: el.id || '',
                className: el.className || '',
                text: el.textContent.trim(),
                isVisible: el.offsetParent !== null,
                rect: el.getBoundingClientRect()
              }));
            };
            
            return {
              exactMatches: captureElementDetails(exactMatches),
              partialMatches: captureElementDetails(partialMatches),
              attributeMatches: captureElementDetails(attributeMatches)
            };
          });
          
          console.log('Potential Tender Results elements:', JSON.stringify(tenderResultsElements, null, 2));
          
          // Use the gathered information to decide if "Tender Results" exists
          const tenderResultsExists = 
            tenderResultsElements.exactMatches.length > 0 || 
            tenderResultsElements.partialMatches.length > 0;
        
          if (tenderResultsExists) {
            console.log('Attempting to click on "Tender Results" tab...');
            
            let tenderResultsClicked = false;
            
            // Method 1: Try clicking exact matches first
            for (const element of tenderResultsElements.exactMatches) {
              try {
                if (!element.isVisible) continue;
                
                // Try to create a precise selector
                let selector = '';
                if (element.id) {
                  selector = `#${element.id}`;
                } else if (element.className && typeof element.className === 'string' && element.className.trim()) {
                  // Create a class-based selector
                  const firstClass = element.className.split(' ')[0].trim();
                  if (firstClass) {
                    selector = `${element.tag.toLowerCase()}.${firstClass}`;
                  }
                }
                
                if (selector) {
                  await page.click(selector, { timeout: 5000 });
                  console.log(`Clicked Tender Results using selector: ${selector}`);
                } else {
                  // Use XPath as fallback
                  await page.evaluate((tag, text) => {
                    const elements = document.querySelectorAll(tag);
                    for (const el of elements) {
                      if (el.textContent.trim() === text) {
                        el.click();
                        return true;
                      }
                    }
                    return false;
                  }, element.tag, 'Tender Results');
                  console.log('Clicked Tender Results using text-based evaluation');
                }
                
                await page.waitForTimeout(3000);
                
                // Check if clicking worked by looking for tender-specific content
                const tenderResultsLoaded = await page.evaluate(() => {
                  const content = document.body.textContent;
                  return content.includes('Successful Tenderer') || 
                         content.includes('Tender Price') ||
                         content.includes('Successful Tender');
                });
                
                if (tenderResultsLoaded) {
                  tenderResultsClicked = true;
                  console.log('Successfully navigated to tender results');
                  
                  if (debug) {
                    await takeScreenshot(page, `${originalIndex}_tender_results_page`);
                    await savePageHtml(page, `${originalIndex}_tender_results_page`);
                  }
                  break;
                }
              } catch (err) {
                console.log(`Failed to click exact match: ${err.message}`);
              }
            }
            
            // Method 2: Try partial matches if exact matches didn't work
            if (!tenderResultsClicked && tenderResultsElements.partialMatches.length > 0) {
              for (const element of tenderResultsElements.partialMatches) {
                try {
                  if (!element.isVisible) continue;
                  
                  // Use XPath to find element by text
                  await page.evaluate((tag, partialText) => {
                    const elements = document.querySelectorAll(tag);
                    for (const el of elements) {
                      if (el.textContent.includes(partialText)) {
                        el.click();
                        return true;
                      }
                    }
                    return false;
                  }, element.tag, 'Tender Results');
                  
                  await page.waitForTimeout(3000);
                  
                  // Check if clicking worked
                  const tenderResultsLoaded = await page.evaluate(() => {
                    const content = document.body.textContent;
                    return content.includes('Successful Tenderer') || 
                           content.includes('Tender Price') ||
                           content.includes('Successful Tender');
                  });
                  
                  if (tenderResultsLoaded) {
                    tenderResultsClicked = true;
                    console.log('Successfully navigated to tender results via partial match');
                    
                    if (debug) {
                      await takeScreenshot(page, `${originalIndex}_tender_results_page_partial`);
                    }
                    break;
                  }
                } catch (err) {
                  console.log(`Failed to click partial match: ${err.message}`);
                }
              }
            }
            
            // Method 3: Try attribute matches or any element with relevant class/id
            if (!tenderResultsClicked && tenderResultsElements.attributeMatches.length > 0) {
              for (const element of tenderResultsElements.attributeMatches) {
                try {
                  if (!element.isVisible) continue;
                  
                  if (element.id) {
                    await page.click(`#${element.id}`);
                  } else if (element.className && typeof element.className === 'string') {
                    const classes = element.className.split(' ')
                      .filter(c => c.trim().length > 0)
                      .map(c => `.${c}`)
                      .join('');
                    
                    if (classes) {
                      await page.click(`${element.tag.toLowerCase()}${classes}`);
                    }
                  }
                  
                  await page.waitForTimeout(3000);
                  
                  // Check if clicking worked
                  const tenderResultsLoaded = await page.evaluate(() => {
                    const content = document.body.textContent;
                    return content.includes('Successful Tenderer') || 
                           content.includes('Tender Price') ||
                           content.includes('Successful Tender');
                  });
                  
                  if (tenderResultsLoaded) {
                    tenderResultsClicked = true;
                    console.log('Successfully navigated to tender results via attribute match');
                    
                    if (debug) {
                      await takeScreenshot(page, `${originalIndex}_tender_results_page_attribute`);
                    }
                    break;
                  }
                } catch (err) {
                  console.log(`Failed to click attribute match: ${err.message}`);
                }
              }
            }
            
            // Method 4: Last resort - look for any tabs or navigation elements
            if (!tenderResultsClicked) {
              console.log('Using fallback method to find tender results tab...');
              
              try {
                // Look for tab/navigation elements
                const tabElements = await page.evaluate(() => {
                  // Common tab containers
                  const tabContainers = document.querySelectorAll('.tabs, .tab-container, .nav, .nav-tabs, ul.navigation, .sidebar');
                  
                  if (tabContainers.length === 0) return null;
                  
                  // Find all clickable items inside tab containers
                  const clickableItems = [];
                  tabContainers.forEach(container => {
                    const items = container.querySelectorAll('a, button, li, div[role="tab"], span[role="button"]');
                    items.forEach(item => {
                      if (item.offsetParent !== null) { // Check if visible
                        clickableItems.push({
                          text: item.textContent.trim(),
                          index: clickableItems.length
                        });
                      }
                    });
                  });
                  
                  return clickableItems;
                });
                
                if (tabElements && tabElements.length > 0) {
                  console.log('Found possible tab elements:', tabElements);
                  
                  // Try clicking each tab element one by one
                  for (let i = 0; i < tabElements.length; i++) {
                    try {
                      await page.evaluate((index) => {
                        // Common tab containers
                        const tabContainers = document.querySelectorAll('.tabs, .tab-container, .nav, .nav-tabs, ul.navigation, .sidebar');
                        
                        let clickableItems = [];
                        tabContainers.forEach(container => {
                          const items = container.querySelectorAll('a, button, li, div[role="tab"], span[role="button"]');
                          items.forEach(item => {
                            if (item.offsetParent !== null) { // Check if visible
                              clickableItems.push(item);
                            }
                          });
                        });
                        
                        if (index < clickableItems.length) {
                          clickableItems[index].click();
                          return true;
                        }
                        return false;
                      }, i);
                      
                      await page.waitForTimeout(2000);
                      
                      // Check if we found tender results
                      const tenderResultsLoaded = await page.evaluate(() => {
                        const content = document.body.textContent;
                        return content.includes('Successful Tenderer') || 
                               content.includes('Tender Price') ||
                               content.includes('Successful Tender');
                      });
                      
                      if (tenderResultsLoaded) {
                        tenderResultsClicked = true;
                        console.log(`Successfully navigated to tender results via tab element ${i}`);
                        
                        if (debug) {
                          await takeScreenshot(page, `${originalIndex}_tender_results_page_tab`);
                        }
                        break;
                      }
                    } catch (err) {
                      console.log(`Failed to click tab element ${i}: ${err.message}`);
                    }
                  }
                }
              } catch (err) {
                console.log(`Error in fallback tab finding: ${err.message}`);
              }
            }
            
            // Now extract the participating tenderers if we successfully navigated to tender results
            if (tenderResultsClicked) {
              console.log('Extracting tenderer information...');
              
              // Wait to ensure the tenderer data is loaded
              await page.waitForTimeout(3000);
              
              if (debug) {
                await takeScreenshot(page, `${originalIndex}_before_tenderer_extraction`);
                await savePageHtml(page, `${originalIndex}_before_tenderer_extraction`);
              }
              
              // Extract tenderer information using multiple strategies
              const tendererData = await page.evaluate(() => {
                // Strategy 1: Look for specific elements containing tenderer information
                function findTenderersBySelectors() {
                  const selectors = [
                    '.tenderer-info', 
                    '.bidder-info', 
                    '.tenderer-name',
                    '.participant-info',
                    '.tender-participant',
                    'tr td:contains("Tenderer")',
                    'div.tenderer',
                    '.tender-details',
                    '.bidder-details'
                  ];
                  
                  for (const selector of selectors) {
                    try {
                      const elements = document.querySelectorAll(selector);
                      if (elements.length > 0) {
                        return Array.from(elements).map(el => el.textContent.trim());
                      }
                    } catch (e) {
                      // Ignore errors with individual selectors
                    }
                  }
                  return null;
                }
                
                // Strategy 2: Look for table rows containing tenderer information
                function findTenderersInTables() {
                  const tables = document.querySelectorAll('table');
                  const tenderersFound = [];
                  
                  tables.forEach(table => {
                    const rows = table.querySelectorAll('tr');
                    rows.forEach(row => {
                      const text = row.textContent.trim();
                      if (text.includes('Tenderer') || 
                          text.includes('Bidder') || 
                          text.includes('Participant')) {
                        tenderersFound.push(text);
                      }
                    });
                  });
                  
                  return tenderersFound.length > 0 ? tenderersFound : null;
                }
                
                // Strategy 3: Look for list items that might contain tenderer information
                function findTenderersInLists() {
                  const lists = document.querySelectorAll('ul, ol');
                  const tenderersFound = [];
                  
                  lists.forEach(list => {
                    const items = list.querySelectorAll('li');
                    items.forEach(item => {
                      const text = item.textContent.trim();
                      if (text.includes('Tenderer') || 
                          text.includes('Bidder') || 
                          text.includes('Participant') ||
                          text.includes('Ltd') ||
                          text.includes('Pte')) {
                        tenderersFound.push(text);
                      }
                    });
                  });
                  
                  return tenderersFound.length > 0 ? tenderersFound : null;
                }
                
                // Strategy 4: Look for specific patterns in the page text
                function findTenderersInPageText() {
                  const pageText = document.body.innerText;
                  const lines = pageText.split('\n').map(line => line.trim()).filter(line => line.length > 0);
                  
                  // Find lines containing keywords
                  const tendererLines = lines.filter(line => 
                    line.includes('Tenderer') || 
                    line.includes('Bidder') || 
                    line.includes('Participant')
                  );
                  
                  // Also look for company names (typical patterns in the context)
                  const companyLines = lines.filter(line => 
                    (line.includes('Pte') || line.includes('Ltd') || line.includes('Limited')) &&
                    !tendererLines.includes(line)
                  );
                  
                  return [...tendererLines, ...companyLines];
                }
                
                // Strategy 5: Extract from structured data if available
                function findTenderersInStructuredData() {
                  // Look for specific divs with tenderer data based on the screenshot patterns
                  const successfulTendererElement = Array.from(document.querySelectorAll('div')).find(div => 
                    div.textContent.includes('SUCCESSFUL TENDERER') || 
                    div.textContent.includes('Successful Tenderer')
                  );
                  
                  if (successfulTendererElement) {
                    const parentElement = successfulTendererElement.parentElement;
                    if (parentElement) {
                      const allText = parentElement.innerText;
                      return [allText];
                    }
                  }
                  
                  return null;
                }
                
                // Try all strategies
                return {
                  bySelectors: findTenderersBySelectors(),
                  inTables: findTenderersInTables(),
                  inLists: findTenderersInLists(),
                  inPageText: findTenderersInPageText(),
                  inStructuredData: findTenderersInStructuredData()
                };
              });
              
              console.log('Tenderer extraction results:', JSON.stringify(tendererData, null, 2));
              
              // Process the extracted tenderer data
              let tenderers = [];
              
              // Prioritize the most structured data sources
              if (tendererData.bySelectors && tendererData.bySelectors.length > 0) {
                tenderers = tendererData.bySelectors;
              } else if (tendererData.inTables && tendererData.inTables.length > 0) {
                tenderers = tendererData.inTables;
              } else if (tendererData.inLists && tendererData.inLists.length > 0) {
                tenderers = tendererData.inLists;
              } else if (tendererData.inStructuredData && tendererData.inStructuredData.length > 0) {
                tenderers = tendererData.inStructuredData;
              } else if (tendererData.inPageText && tendererData.inPageText.length > 0) {
                tenderers = tendererData.inPageText;
              }
              
              // Clean up tenderer data
              const cleanedTenderers = tenderers.map(text => {
                // Remove common prefixes/labels
                let cleaned = text.replace(/^(Tenderer|Bidder|Participant|Successful Tenderer)[\s\-:]+/i, '')
                                  .trim();
                
                // Additional cleaning for known patterns from the specific site
                cleaned = cleaned.replace(/^\d+[\.\)]\s*/, ''); // Remove numeric prefixes like "1. " or "2) "
                
                return cleaned;
              }).filter(text => text.length > 0);
              
              // Deduplicate tenderers (sometimes the same company appears multiple times)
              const uniqueTenderers = [...new Set(cleanedTenderers)];
              
              console.log(`Processed tenderers for ${location}:`, uniqueTenderers);
              
              // Update the output data with tenderers
              const rowIndex = outputData.findIndex(row => row.Location === location);
              if (rowIndex !== -1) {
                // Clear existing "Participating Tenderer" fields if they exist
                for (const key in outputData[rowIndex]) {
                  if (key.includes('Participating Tenderer') || key.includes('Other Particpating Tenderer')) {
                    delete outputData[rowIndex][key];
                  }
                }
                
                // Add the successful tenderer first (if found)
                const successfulTendererIndex = uniqueTenderers.findIndex(t => 
                  t.toLowerCase().includes('successful') || 
                  t.toLowerCase().includes('winner') || 
                  t.toLowerCase().includes('awarded')
                );
                
                if (successfulTendererIndex !== -1) {
                  outputData[rowIndex]['Name of Successful Tenderer'] = formatCompanyName(uniqueTenderers[successfulTendererIndex]);
                  uniqueTenderers.splice(successfulTendererIndex, 1); // Remove from the list
                }
                
                // Add remaining tenderers
                uniqueTenderers.forEach((tenderer, idx) => {
                  outputData[rowIndex][`Name of Other Particpating Tenderer ${idx+1}`] = formatCompanyName(tenderer);
                });
                
                // Mark this location as successfully processed
                success = true;
              }
            } else {
              console.log(`Failed to navigate to tender results for ${location}`);
              
              if (debug) {
                await takeScreenshot(page, `${originalIndex}_failed_tender_results`);
              }
            }
          } else {
            console.log(`No Tender Results tab found for ${location}`);
            
            // We should still mark this as a success, just no tender results available
            success = true;
          }
        
        } catch (error) {
          console.error(`Error processing location ${location} (attempt ${attemptCount}):`, error.message);
          
          if (debug) {
            await takeScreenshot(page, `${originalIndex}_error_attempt_${attemptCount}`);
            await savePageHtml(page, `${originalIndex}_error_attempt_${attemptCount}`);
          }
          
          // If we've tried enough times, log and continue
          if (attemptCount >= retries) {
            console.log(`Failed to process ${location} after ${retries} attempts, moving on...`);
          } else {
            // Brief pause before retry
            console.log(`Waiting before retry attempt ${attemptCount + 1}...`);
            await page.evaluate(() => new Promise(resolve => setTimeout(resolve, 5000)));
          }
        }
      }
      
      // Update processing status in spreadsheet
      if (!success) {
        // If all attempts failed, note this in the output data
        const rowIndex = outputData.findIndex(row => row.Location === location);
        if (rowIndex !== -1) {
          outputData[rowIndex]['Processing Status'] = `Failed after ${retries} attempts`;
        }
      } else {
        // Mark as success in the output data
        const rowIndex = outputData.findIndex(row => row.Location === location);
        if (rowIndex !== -1) {
          outputData[rowIndex]['Processing Status'] = 'Success';
        }
      }
      
      // Save interim results after each location
      try {
        const interimWorkbook = XLSX.utils.book_new();
        const interimWorksheet = XLSX.utils.json_to_sheet(outputData);
        XLSX.utils.book_append_sheet(interimWorkbook, interimWorksheet, 'Scraping Results');
        XLSX.writeFile(interimWorkbook, 'ura_sites_interim_results.xlsx');
        console.log('Interim results saved');
      } catch (err) {
        console.error('Error saving interim results:', err.message);
      }
    }
    
    // Write the final updated data back to a new Excel file
    const newWorkbook = XLSX.utils.book_new();
    const newWorksheet = XLSX.utils.json_to_sheet(outputData);
    XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, 'Scraping Results');
    
    // Add timestamp to the filename
    const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
    const outputFilename = `ura_sites_with_tenderers_${timestamp}.xlsx`;
    
    XLSX.writeFile(newWorkbook, outputFilename);
    
    console.log(`Scraping completed. Results saved to ${outputFilename}`);
    
  } catch (error) {
    console.error('Fatal error during scraping:', error);
  } finally {
    // Close the browser
    await browser.close();
  }
}

// Helper function to take screenshots for debugging
async function takeScreenshot(page, name) {
  try {
    const screenshotPath = `./screenshots/${name}.png`;
    // Make sure the directory exists
    if (!fs.existsSync('./screenshots')) {
      fs.mkdirSync('./screenshots', { recursive: true });
    }
    await page.screenshot({ path: screenshotPath, fullPage: true });
    console.log(`Screenshot saved to ${screenshotPath}`);
  } catch (error) {
    console.error('Error taking screenshot:', error);
  }
}

// Helper function to get page HTML for debugging
async function savePageHtml(page, name) {
  try {
    const htmlPath = `./html/${name}.html`;
    // Make sure the directory exists
    if (!fs.existsSync('./html')) {
      fs.mkdirSync('./html', { recursive: true });
    }
    const html = await page.content();
    fs.writeFileSync(htmlPath, html);
    console.log(`HTML saved to ${htmlPath}`);
  } catch (error) {
    console.error('Error saving HTML:', error);
  }
}

// Run the scraper with command line arguments
async function main() {
  // Parse command line arguments
  const args = process.argv.slice(2);
  const options = {
    headless: args.includes('--headless'),
    debug: args.includes('--debug'),
    startIndex: args.includes('--start') ? parseInt(args[args.indexOf('--start') + 1]) : 0,
    endIndex: args.includes('--end') ? parseInt(args[args.indexOf('--end') + 1]) : Infinity,
    retries: args.includes('--retries') ? parseInt(args[args.indexOf('--retries') + 1]) : 3
  };
  
  console.log('Running with options:', options);
  
  try {
    await runScraper(options);
    console.log('Scraper completed successfully');
  } catch (error) {
    console.error('Fatal error in scraper:', error);
    process.exit(1);
  }
}

// Run the main function
main();