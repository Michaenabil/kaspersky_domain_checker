import pandas as pd
import requests
import sys
from time import sleep

def get_domain_info(domain, api_token):
    """Query Kaspersky's API for domain information"""
    url = f'https://opentip.kaspersky.com/api/v1/search/domain?request={domain}'
    
    headers = {
        'x-api-key': api_token,
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/136.0.0.0 Safari/537.36 Edg/136.0.0.0'
    }

    try:
        response = requests.get(url, headers=headers, timeout=40 , verify=False )
        response.raise_for_status()
        
        # Add delay to avoid rate limiting
        sleep(1)
        
        data = response.json()
        
        # Extract information based on API response structure
        zone = data.get('Zone', 'Unknown')
        categories = [cat.get('Name', 'Unknown') for cat in data.get('Categories', [])]
        
        return {
            'Zone': zone,
            'Categories': ', '.join(categories) if categories else 'None',
            'Status': 'Success'
        }
        
    except requests.exceptions.RequestException as e:
        return {
            'Zone': 'Error',
            'Categories': 'Error',
            'Status': f'Request failed: {str(e)}'
        }
    except Exception as e:
        return {
            'Zone': 'Error',
            'Categories': 'Error',
            'Status': f'Processing error: {str(e)}'
        }

def process_domains(input_file, output_file, api_token):
    """Process Excel file with domains and save results"""
    try:
        # Read input Excel
        df = pd.read_excel(input_file)
        
        # Verify required column exists
        if 'Domain' not in df.columns:
            raise ValueError("Input file must contain 'domain' column")
            
        total = len(df)
        if total == 0:
            print("No domains found in the input file.")
            return False
            
        processed = 0
        
        print(f"Starting processing of {total} domains...")
        
        # Process domains
        results = []
        for domain in df['Domain']:
            if pd.notna(domain) and domain.strip():
                result = get_domain_info(domain.strip(), api_token)
            else:
                result = {
                    'Zone': 'Invalid',
                    'Categories': 'Invalid',
                    'Status': 'Empty domain value'
                }
            results.append(result)
            processed += 1
            
            # Calculate progress
            progress = (processed / total) * 100
            print(f"\rProcessed {processed}/{total} ({progress:.1f}%)", end='', flush=True)
        
        # Add results to dataframe
        df['Kaspersky_Zone'] = [res['Zone'] for res in results]
        df['Kaspersky_Categories'] = [res['Categories'] for res in results]
        df['API_Status'] = [res['Status'] for res in results]
        
        # Save to Excel
        df.to_excel(output_file, index=False)
        print(f"\nSuccessfully processed {processed} domains. Results saved to {output_file}")
        
        return True
        
    except Exception as e:
        print(f"\nError processing file: {str(e)}")
        return False

if __name__ == "__main__":
    if len(sys.argv) != 4:
        print("Usage: python domain_checker.py <input.xlsx> <output.xlsx> <api_token>")
        sys.exit(1)
    
    input_path = sys.argv[1]
    output_path = sys.argv[2]
    token = sys.argv[3]
    
    process_domains(input_path, output_path, token)