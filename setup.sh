#!/bin/bash
# URA Web Scraper Setup Script

# Create project directory
mkdir -p ura-scraper
cd ura-scraper

# Create package.json file
cat > package.json << 'EOL'
{
  "name": "ura-scraper",
  "version": "1.0.0",
  "description": "Web scraper for URA vacant sites",
  "main": "scraper.js",
  "scripts": {
    "start": "node scraper.js",
    "debug": "node scraper.js --debug",
    "headless": "node scraper.js --headless"
  },
  "dependencies": {
    "puppeteer": "^21.3.8",
    "xlsx": "^0.18.5"
  }
}
EOL

# Install dependencies
echo "Installing dependencies..."
npm install

# Create directories for debug output
mkdir -p screenshots html

echo "Setup completed successfully!"
echo "To run the scraper:"
echo "  npm start            - Run in visible browser mode"
echo "  npm run debug        - Run with debugging (screenshots)"
echo "  npm run headless     - Run in headless mode (no visible browser)"
echo ""
echo "Additional options:"
echo "  node scraper.js --start 2 --end 5  - Process only locations 2-5"
echo "  node scraper.js --retries 5        - Try up to 5 times per location"