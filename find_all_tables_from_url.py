# -*- coding: utf-8 -*-
"""
Created on Fri Jan 30 09:49:35 2026

@author: Krishna
"""

from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time

def fetch_html(url):
    """Fetch HTML content using Selenium (for JavaScript-rendered pages)."""
    try:
        
        # Set up Chrome options
        chrome_options = Options()
        chrome_options.add_argument('--headless')  # Run in background
        chrome_options.add_argument('--no-sandbox')
        chrome_options.add_argument('--disable-dev-shm-usage')
        chrome_options.add_argument('--disable-gpu')
        chrome_options.add_argument('--user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36')
        
        # Initialize driver
        driver = webdriver.Chrome(options=chrome_options)
        
        print("Loading page with Selenium (this may take a moment)...")
        driver.get(url)
        
        # Wait for tables to load (adjust wait time as needed)
        time.sleep(3)  # Give JavaScript time to render
        
        # Optionally wait for specific element
        try:
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.TAG_NAME, "table"))
            )
        except:
            print("Warning: No tables found after waiting")
        
        # Get the page source after JavaScript execution
        html_content = driver.page_source
        driver.quit()
        
        return html_content
        
    except ImportError:
        print("\n Selenium not installed. Install it with:")
        print("   pip install selenium")
        print("   You may also need to install ChromeDriver")
        return None
    except Exception as e:
        print(f"Error fetching URL with Selenium: {e}")
        return None