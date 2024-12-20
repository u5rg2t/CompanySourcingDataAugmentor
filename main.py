import pandas as pd
import requests
from datetime import datetime
import os
from time import sleep
import sys
from dotenv import load_dotenv
import time  # Add this to your imports at the top

# Load environment variables
load_dotenv()

# Companies House API configuration
API_KEY = os.getenv('COMPANIES_HOUSE_API_KEY')
if not API_KEY:
    raise ValueError("COMPANIES_HOUSE_API_KEY not found in environment variables")

BASE_URL = 'https://api.company-information.service.gov.uk'

MAX_REQUESTS = 600 # As per API docs 600 per 5 mins..
MAX_REQUESTS_PER = 300 # 5 mins * 60 = 300 seconds
SAFE_BUFFER = 1.5  # 50% more time as a buffer

# Calculate delay needed for API rate limiting 
DELAY_SECONDS = (MAX_REQUESTS_PER / MAX_REQUESTS) * SAFE_BUFFER

def calculate_age(dob_str):
    """Calculate age from DOB string in format 'YYYY-MM'"""
    try:
        dob = datetime.strptime(dob_str, '%Y-%m')
        today = datetime.now()
        age = today.year - dob.year
        if today.month < dob.month:
            age -= 1
        return age
    except:
        return None

def get_company_info(company_number, max_retries=3):
    """Fetch company status and officers from Companies House API"""

    for attempt in range(max_retries):
        try:
            # Get company profile
            profile_response = requests.get(f"{BASE_URL}/company/{company_number}", auth=(API_KEY, ''))
            
            if profile_response.status_code == 429:
                print(f"Rate limit hit, waiting for {DELAY_SECONDS * 2} seconds...")
                sleep(DELAY_SECONDS * 2)  # Double the delay on rate limit
                continue
                
            if profile_response.status_code != 200:
                print(f"Failed to get company profile for {company_number}. Status code: {profile_response.status_code}")
                print(f"Error Response: {profile_response.text}")
                return None, None, []
            
            company_data = profile_response.json()
            company_status = company_data.get('company_status', '')
            
            # Get status change date - use date_of_cessation for all status types
            status_date = company_data.get('date_of_cessation')
            
            # If status is active or date is empty/null, set to None
            if company_status == 'active' or not status_date:
                status_date = None
            
            # Add delay between requests
            sleep(DELAY_SECONDS)
            
            # Get officers
            officers_response = requests.get(f"{BASE_URL}/company/{company_number}/officers", auth=(API_KEY, ''))
            
            if officers_response.status_code == 429:
                print(f"Rate limit hit, waiting for {DELAY_SECONDS * 2} seconds...")
                sleep(DELAY_SECONDS * 2)  # Double the delay on rate limit
                continue
                
            if officers_response.status_code != 200:
                print(f"Failed to get officers for {company_number}. Status code: {officers_response.status_code}")
                print(f"Error Response: {officers_response.text}")
                return company_status, status_date, []
            
            # Process officers
            officers_data = officers_response.json()
            
            officers = []
            for officer in officers_data.get('items', []):
                if officer.get('resigned_on'):  # Skip resigned officers
                    continue
                    
                officer_info = {
                    'name': officer.get('name'),
                    'dob': None,
                    'age': None
                }
                
                if officer.get('date_of_birth'):
                    dob_year = officer['date_of_birth'].get('year')
                    dob_month = officer['date_of_birth'].get('month')
                    if dob_year and dob_month:
                        dob_str = f"{dob_year}-{dob_month:02d}"
                        officer_info['dob'] = dob_str
                        officer_info['age'] = calculate_age(dob_str)
                
                officers.append(officer_info)
            
            return company_status, status_date, officers
            
        except Exception as e:
            print(f"Error processing company {company_number}: {str(e)}")
            if attempt < max_retries - 1:  # if not the last attempt
                sleep(DELAY_SECONDS * 2)  # wait before retrying
                continue
            return None, None, []
    
    return None, None, []  # if all retries failed

def process_excel():
    # Check if command line arguments are provided
    if len(sys.argv) < 3:
        print("Please provide [the input file path], [company number column name] and [optional: record limit] as command line arguments")
        print("Usage: python main.py 'path/to/your/excel/file.xlsx' 'Company Number Column Name' [optional: record limit]")
        print("Example: python main.py 'companies.xlsx' 'Cro Nbr' 100")
        return
        
    input_file = sys.argv[1]
    company_number_column = sys.argv[2]
    
    # Get optional record limit
    record_limit = None
    if len(sys.argv) > 3:
        try:
            record_limit = int(sys.argv[3])
            print(f"Processing first {record_limit} records...")
        except ValueError:
            print("Invalid record limit provided. Processing all records...")
    
    # Read the input file
    try:
        df = pd.read_excel(input_file)
        if company_number_column not in df.columns:
            print(f"Error: Column '{company_number_column}' not found in the Excel file")
            print(f"Available columns: {', '.join(df.columns)}")
            return
            
        # Limit records if specified
        if record_limit:
            df = df.head(record_limit)
            
    except Exception as e:
        print(f"Error reading input file: {str(e)}")
        return
    
    print(f"Processing {len(df)} companies...")
    start_time = time.time()
    
    # Add new columns
    df['Company_Status'] = None
    df['Status_Change_Date'] = None
    df['Active_Directors'] = None
    df['Directors_Ages'] = None
    
    # Process each row
    for index, row in df.iterrows():
        company_start_time = time.time()
        
        company_number = str(row[company_number_column]).strip()
        company_number = company_number.zfill(8)
        current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        print(f"[{current_time}] Processing company {company_number} ({index + 1}/{len(df)})...")
        
        # Use calculated delay
        sleep(DELAY_SECONDS)
        
        status, status_date, officers = get_company_info(company_number)
        
        # Update DataFrame
        df.at[index, 'Company_Status'] = status
        df.at[index, 'Status_Change_Date'] = status_date
        if officers:
            df.at[index, 'Active_Directors'] = '; '.join([o['name'] for o in officers])
            df.at[index, 'Directors_Ages'] = '; '.join([str(o['age']) for o in officers if 'age' in o])
        
        # Print time taken for this company
        company_time = time.time() - company_start_time
    
    # Save to new file
    try:
        output_filename = 'output_' + datetime.now().strftime('%Y%m%d_%H%M%S') + '.xlsx'
        df.to_excel(output_filename, index=False)
        total_time = time.time() - start_time
        print(f"\nTotal processing time: {total_time:.2f} seconds ({total_time/60:.2f} minutes)")
        print(f"Average time per company: {total_time/len(df):.2f} seconds")
        print(f"Results saved to {output_filename}")
    except Exception as e:
        print(f"Error saving output file: {str(e)}")

if __name__ == "__main__":
    process_excel()
