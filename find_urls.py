import requests
from bs4 import BeautifulSoup
import re

# Fetch the download page
url = "https://sco.ca.gov/upd_download_property_records.html"
print(f"Fetching {url}...")

headers = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}

try:
    response = requests.get(url, headers=headers, timeout=30)
    response.raise_for_status()
    
    # Parse the HTML
    soup = BeautifulSoup(response.text, 'html.parser')
    
    # Find all links that might be download links
    print("\nFound links containing 'Properties reported':")
    print("-" * 50)
    
    # Look for links with text containing "Properties reported"
    for link in soup.find_all('a'):
        link_text = link.get_text(strip=True)
        href = link.get('href', '')
        
        if 'Properties reported' in link_text or 'All properties' in link_text:
            # Make absolute URL if relative
            if href and not href.startswith('http'):
                if href.startswith('/'):
                    href = f"https://sco.ca.gov{href}"
                else:
                    href = f"https://sco.ca.gov/{href}"
            
            print(f"\nText: {link_text}")
            print(f"URL:  {href}")
    
    # Also look for any .zip or .csv links
    print("\n\nAll .zip or .csv links found on the page:")
    print("-" * 50)
    
    for link in soup.find_all('a', href=True):
        href = link['href']
        if '.zip' in href.lower() or '.csv' in href.lower():
            link_text = link.get_text(strip=True)
            
            # Make absolute URL if relative
            if not href.startswith('http'):
                if href.startswith('/'):
                    href = f"https://sco.ca.gov{href}"
                else:
                    href = f"https://sco.ca.gov/{href}"
            
            print(f"\nText: {link_text}")
            print(f"URL:  {href}")

except requests.exceptions.RequestException as e:
    print(f"Error fetching the page: {e}")
except Exception as e:
    print(f"Error parsing the page: {e}")
