# -*- coding: utf-8 -*-
"""
Created on Mon Jan 26 11:45:33 2026

@author: Krishna
"""

from bs4 import BeautifulSoup
import os
import json
import re
import requests as req
from find_all_tables_from_url import fetch_html


def find_tables(html_content, path):
    soup = BeautifulSoup(html_content, 'html.parser')
    tables = soup.find_all('table')
    path = os.path.normpath(path)

    if not tables:
        return 0
    os.makedirs(path, exist_ok=True)
    for i, table in enumerate(tables):
        filename = os.path.join(path, f"table{i}.html")
        with open(filename, "w", encoding="utf-8") as file:
            file.write(table.prettify())

    return len(tables)




def load_config(config_file="config.json"):
    """Load configuration from JSON file (supports multiple formats)."""
    try:
        with open(config_file, "r", encoding="utf-8") as file:
            config = json.load(file)
        return config
    except FileNotFoundError:
        print(f"Error: Config file '{config_file}' not found.")
        return None
    except json.JSONDecodeError as e:
        print(f"Error: Invalid JSON in config file: {e}")
        return None

def url_to_dirname(url):
    """Convert URL to a safe directory name."""
    # Extract domain and path
    match = re.search(r'://([^/]+)(/[^?]*)?', url)
    if match:
        domain = match.group(1).replace('www.', '')
        path = match.group(2) or ''
        
        # Clean up
        name = domain + path
        name = re.sub(r'[^\w\-_]', '_', name)
        name = re.sub(r'_+', '_', name)
        name = name.strip('_')
        
        return name[:50]  # Limit length
    return "unknown"

def normalize_config(config):
    """Convert various config formats to standard format."""
    sources = []
    defaults = config.get("default_settings", {})
    
    # Format 1: Simple URL list
    if "urls" in config:
        base_dir = config.get("output_base_dir", "tables")
        
        for i, url in enumerate(config["urls"]):
            dirname = url_to_dirname(url)
            sources.append({
                "name": f"Source {i+1}",
                "url": url,
                "output_dir": os.path.join(base_dir, dirname)
            })
    
    # Format 2: Sources array (current format)
    elif "sources" in config:
        sources = config["sources"]
    
    # Format 3: Single URL (backward compatibility)
    elif "input_url" in config:
        sources = [{
            "name": "Single Source",
            "url": config["input_url"],
            "output_dir": config.get("output_dir", "tables/output"),
        }]
    
    return sources, defaults

def convert_tables_to_excel(input_dir, api_url="http://127.0.0.1:8000/convert-table-to-excel"):
    """Convert all HTML tables in a directory to Excel files using the API endpoint."""
    output_dir = f"{input_dir}_excel"
    os.makedirs(output_dir, exist_ok=True)
    
    if not os.path.exists(input_dir):
        return
    
    html_files = [f for f in os.listdir(input_dir) if f.endswith('.html')]
    
    for html_file in html_files:
        html_path = os.path.join(input_dir, html_file)
        with open(html_path, 'r', encoding='utf-8') as f:
            html_content = f.read()
        
        table_id = os.path.splitext(html_file)[0]
        
        data_dict = {
            "html": html_content,
            "table_id": table_id,
            "alternate_colors": True,
            "hyperlink": "https://www.google.com/"            
        }        
        
        try:
            
            
            response = req.get(api_url, json = data_dict, timeout=30.0)
            response.raise_for_status()
            
            excel_path = os.path.join(output_dir, f"{table_id}.xlsx")
            with open(excel_path, 'wb') as f:
                f.write(response.content)
        except Exception as e:
            print(f"Error converting {html_file}: {e}")
            pass


def main():
    config = load_config("config.json")
    if config is None:
        return
    
    sources, defaults = normalize_config(config)
    
    if not sources:
        return
    
    for source in sources:
        url = source.get("url")
        output_dir = source.get("output_dir")
        
        if not url or not output_dir:
            continue
        
        html_content = fetch_html(url)
        
        if html_content is None:
            continue
        
        find_tables(html_content, output_dir)
        convert_tables_to_excel(output_dir)

if __name__ == "__main__":
    main()