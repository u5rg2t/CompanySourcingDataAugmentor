# Company Data Augmentor

A Python script that enriches company data using the Companies House API. This tool fetches additional information such as company status, status change dates, and director details for a list of companies provided in an Excel file.

## Features

- Fetches company status and status change dates
- Retrieves current directors and their ages
- Handles API rate limiting automatically
- Processes companies in bulk from Excel files
- Supports optional record limit for testing
- Includes progress tracking with timestamps
- Provides detailed timing statistics

## Prerequisites

- Python 3.6+
- Companies House API key ([Get one here](https://developer.company-information.service.gov.uk/))
- Input Excel file containing company numbers

## Installation

1. Clone the repository: 

```bash
git clone https://github.com/yourusername/CompanySourcingDataAugmentor.git
cd CompanySourcingDataAugmentor

2. Create and activate a virtual environment:
```bash
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate
```

3. Install required packages:
```bash
pip install -r requirements.txt
```

4. Create a `.env` file in the project root and add your API key:
```
COMPANIES_HOUSE_API_KEY=your_api_key_here
```

## Usage

Run the script with the following command:
```bash
python main.py "path/to/input.xlsx" "Company Number Column" [optional: record_limit]
```

### Arguments:
- `path/to/input.xlsx`: Path to your Excel file containing company numbers
- `Company Number Column`: Name of the column containing company numbers
- `record_limit` (optional): Number of records to process (for testing)

### Example:
```bash
python main.py "companies.xlsx" "CRO Number" 10
```

## Output

The script creates a new Excel file with the following additional columns:
- `Company_Status`: Current status of the company
- `Status_Change_Date`: Date when the status changed (for non-active companies)
- `Active_Directors`: List of current directors
- `Directors_Ages`: Ages of current directors

Output files are named with timestamp: `output_YYYYMMDD_HHMMSS.xlsx`

## Rate Limiting

The script respects Companies House API rate limits:
- Maximum 600 requests per 5 minutes
- Built-in delays between requests
- Automatic retry on rate limit errors

## Error Handling

- Validates input file and column names
- Handles API errors gracefully
- Retries on temporary failures
- Provides detailed error messages

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License

This project is licensed under the MIT License - see the LICENSE file for details.

## Acknowledgments

- [Companies House API](https://developer.company-information.service.gov.uk/) for providing the data
- Contributors and maintainers

## Support

For support, please open an issue in the GitHub repository.