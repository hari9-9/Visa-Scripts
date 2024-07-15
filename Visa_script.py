import requests
from bs4 import BeautifulSoup
import urllib
import os
import pandas as pd

def cleanup_workspace():
    """
    Cleans up the current working directory by deleting any .ods files present.
    """
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

def download_ods():
    """
    Scrapes the provided URL for a link to an .ods file, downloads it, and saves it locally.
    Returns:
    - 1 if successful download
    - -1 if download fails or no .ods file found
    """
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
                return 1
            else:
                print(f'Failed to download the file. Status code: {ods_response.status_code}')
                return -1
        else:
            print('No .ods file found on the page')
            return -1
    else:
        print(f'Failed to retrieve the page. Status code: {response.status_code}')
        return -1

def get_ods_path():
    """
    Retrieves the path of the .ods file in the current directory.
    Returns:
    - File path if .ods file found
    - None if no .ods file found
    """
    # Get the current directory
    current_directory = os.getcwd()

    # List all files in the current directory
    files = os.listdir(current_directory)

    # Iterate over each file and collect .ods file names
    for file in files:
        if file.endswith('.ods'):
            return file
    
    return None

def parse_ods(file_path):
    """
    Parses the .ods file into a pandas DataFrame, filters columns, and cleans up data.
    Returns:
    - Filtered DataFrame
    """
    # Read the .ods file into a pandas DataFrame
    df = pd.read_excel(file_path, engine="odf")
    
    # Locate the row containing the desired information
    row_index = df['Unnamed: 1'].str.contains('Decisions for period', na=False)

    # Extract the value from the DataFrame
    value = df.loc[row_index, 'Unnamed: 1'].values[0]
    print(value)
    # Filter relevant columns
    filtered_df = df.iloc[:, [2, 3]]  # Assuming columns 2 and 3 are relevant (adjust indices as needed)

    # Drop rows with NaN values
    filtered_df = filtered_df.dropna()

    # Rename columns based on the first row and drop the first row
    filtered_df = filtered_df.rename(columns=filtered_df.iloc[0]).drop(filtered_df.index[0])

    # Reset index for clean output
    filtered_df.reset_index(drop=True, inplace=True)

    # Print the cleaned DataFrame
    #print(filtered_df)

    return filtered_df

def check_application(df, applications):
    """
    Checks for specific Application Numbers in the DataFrame and prints their decisions.
    """
    for app in applications:
        if app in df['Application Number'].values:
            decision = df.loc[df['Application Number'] == app, 'Decision'].iloc[0]
            print(f"Decision for Application Number {app}: {decision}")
        else:
            print(f"Application Number {app} not found.")

# Clean up any existing .ods files in the directory
cleanup_workspace()

# Attempt to download the .ods file
result = download_ods()

# If download was successful, proceed with parsing and checking applications
if result:
    # Get the path of the downloaded .ods file
    path = get_ods_path()
    
    # Parse the .ods file into a DataFrame
    if path:
        df = parse_ods(path)
        
        # Example applications to check
        applications = [68049422, 6804942]
        
        # Check applications in the parsed DataFrame
        check_application(df, applications)
    else:
        print("Something went wrong while locating the ODS file")