import requests
from bs4 import BeautifulSoup
import urllib
import os


# Get the current directory
current_directory = os.getcwd()

# List all files in the current directory
files = os.listdir(current_directory)

# Iterate over each file and delete .ods files
for file in files:
    if file.endswith('.ods'):
        try:
            os.remove(file)
            print(f"Deleted: {file}")
        except Exception as e:
            print(f"Failed to delete {file}: {e}")
# URL of the page to scrape
url = 'https://www.ireland.ie/en/india/newdelhi/services/visas/processing-times-and-decisions/'

# Headers to mimic a browser request
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
}

# Make a request to the website
response = requests.get(url, headers=headers)

# Check if the request was successful
if response.status_code == 200:
    # Parse the page content
    soup = BeautifulSoup(response.content, 'lxml')

    # Find all links on the page
    links = soup.find_all('a')

    # Initialize a variable to store the ods link
    ods_link = None

    # Iterate over the links to find the one with .ods file
    for link in links:
        href = link.get('href', '')
        if href.endswith('.ods'):
            ods_link = href
            break

    # If the ods link was found
    if ods_link:
        # If the link is relative, construct the full URL
        if not ods_link.startswith('http'):
            ods_link = urllib.parse.urljoin(url, ods_link)
        
        # Download the file
        file_name = ods_link.split('/')[-1]
        
        # Make a request to download the ods file
        ods_response = requests.get(ods_link, headers=headers)
        
        # Check if the download request was successful
        if ods_response.status_code == 200:
            # Write the content to a file
            with open(file_name, 'wb') as file:
                file.write(ods_response.content)
            print(f'Successfully downloaded {file_name}')
        else:
            print(f'Failed to download the file. Status code: {ods_response.status_code}')
    else:
        print('No .ods file found on the page')
else:
    print(f'Failed to retrieve the page. Status code: {response.status_code}')
