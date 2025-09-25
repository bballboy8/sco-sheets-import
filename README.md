# SCO Sheets Import

A Python script that downloads property records from the California State Controller's Office (SCO) and imports them into a Google Sheets document. The script applies business logic filtering, categorizes records as Business or Individual, and adds required metadata columns.

## Features

- **Data Filtering**: Applies comprehensive filtering rules based on dollar value and data completeness
- **Business Classification**: Automatically categorizes records as "Business" or "Individual" using a comprehensive business terms list
- **Metadata Addition**: Adds required columns (Upload Date, Type, Confidence Level, Stage)
- **Retry Logic**: Robust error handling with exponential backoff for API calls
- **Batch Processing**: Efficient 1000-row batches with configurable rate limiting
- **Progress Tracking**: Detailed logging of download progress and import statistics
- **Deduplication**: Automatic detection and prevention of duplicate records

## Prerequisites

- Python 3.7 or higher
- Google Cloud Project with Google Sheets API and Google Drive API enabled
- Google Service Account with appropriate permissions

## Setup

### 1. Clone the Repository

```bash
git clone <repository-url>
cd sco-sheets-import
```

### 2. Create Virtual Environment

```bash
python -m venv .venv
source .venv/bin/activate  # On Windows: .venv\Scripts\activate
```

### 3. Install Dependencies

```bash
pip install -r requirements.txt
```

### 4. Google Cloud Setup

1. Go to the [Google Cloud Console](https://console.cloud.google.com/)
2. Create a new project or select an existing one
3. Enable the following APIs:
   - Google Sheets API
   - Google Drive API
4. Create a Service Account:
   - Go to IAM & Admin > Service Accounts
   - Click "Create Service Account"
   - Give it a name and description
   - Grant the following roles:
     - Editor (or more specific roles for Sheets and Drive)
5. Create and download a JSON key for the service account
6. Share your Google Sheets document with the service account email

### 5. Environment Configuration

1. Copy the sample environment file:
   ```bash
   cp env.sample .env
   ```

2. Edit `.env` and fill in your values:
   - `SHEET_ID`: The ID of your Google Sheets document (found in the URL)
   - `GOOGLE_CREDENTIALS`: The entire JSON content from your service account key file

## Usage

### Basic Usage

```bash
python sco_import.py
```

### Configuration Options

The script includes several configuration options in `sco_import.py`:

- `USE_ALL_ZIPS`: Set to `True` to process all ZIP files, `False` for testing with just the first one
- `APPEND_BATCH_SIZE`: Number of rows to append in each batch (default: 1000)
- `API_PAUSE_SEC`: Pause between API calls in seconds (default: 1.5)
- `MAX_RECORD_ROWS`: Maximum number of rows to import (default: 500,000)

### Data Sources

The script downloads data from the following SCO URLs:
- 01_From_0_To_Below_10.zip
- 02_From_10_To_Below_100.zip  
- 03_From_100_To_Below_500.zip
- 04_From_500_To_Beyond.zip

## Data Processing Logic

### Filtering Rules

The script applies the following filtering criteria to ensure data quality:

1. **Dollar Value Filter**: Only includes records with `CURRENT_CASH_BALANCE` over $6,000
2. **Owner Count Filters**:
   - Records with 2 owners must have value ≥ $11,999
   - Records with 3+ owners must have value ≥ $17,999
3. **Data Completeness**: Excludes records with:
   - Blank, "Unknown", or missing owner names
   - Blank, "Unknown", or missing addresses

### Business Classification

Records are automatically classified as "Business" or "Individual" based on the presence of business terms in the owner name. The script uses a comprehensive list of 100+ business indicators including:

- Legal entities: LLC, INC, CORP, CORPORATION, LLP, etc.
- Business types: COMPANY, FIRM, ASSOCIATES, HOSPITAL, etc.
- Industries: BANK, INSURANCE, MORTGAGE, REAL ESTATE, etc.
- Government: COUNTY, CITY, STATE, DEPARTMENT, etc.
- And many more...

### Metadata Columns

Each processed record includes:
- **CREATED_BY_DATE**: Date when the data was imported (YYYY-MM-DD format)
- **TYPE**: "Business" or "Individual" classification
- **CONFIDENCE_LEVEL**: Set to "100%" for all records
- **STAGE**: Set to "Leads" for all records

## How It Works

1. **Authentication**: Uses Google Service Account credentials to authenticate with Google Sheets API
2. **Download**: Downloads ZIP files from the SCO website containing CSV property records (with retry logic)
3. **Processing**: Extracts and processes CSV files from the ZIP archives
4. **Filtering**: Applies business rules to filter records based on:
   - Dollar value (must be over $6,000)
   - Owner count thresholds (2+ owners: $11,999+, 3+ owners: $17,999+)
   - Data completeness (no blank/unknown owner names or addresses)
5. **Classification**: Categorizes records as "Business" or "Individual" using comprehensive business terms
6. **Metadata Addition**: Adds required columns (Upload Date, Type, Confidence Level, Stage)
7. **Deduplication**: Identifies existing records using the "Property ID" column to avoid duplicates
8. **Import**: Appends new records to the Google Sheets document in 1000-row batches with retry logic
9. **Progress Tracking**: Provides detailed logging of download progress and import statistics

## Output

The script creates or updates a "Records" worksheet in your Google Sheets document with:
- **Filtered property records** that meet all business criteria
- **Business/Individual classification** for each record
- **Required metadata columns**:
  - `CREATED_BY_DATE`: Date when the data was processed
  - `TYPE`: "Business" or "Individual" classification
  - `CONFIDENCE_LEVEL`: Set to "100%" for all records
  - `STAGE`: Set to "Leads" for all records
- Automatic header row creation
- Deduplication based on Property ID
- Batch processing for efficient API usage

## Troubleshooting

### Common Issues

1. **Authentication Errors**: Ensure your service account has the correct permissions and the JSON credentials are properly formatted
2. **Sheet Access**: Make sure the service account email has edit access to your Google Sheets document
3. **API Quotas**: The script includes rate limiting and retry logic, but you may need to adjust `API_PAUSE_SEC` if you hit quota limits
4. **Memory Issues**: For very large datasets, consider reducing `APPEND_BATCH_SIZE` or `MAX_RECORD_ROWS`
5. **Network Issues**: The script includes retry logic with exponential backoff for both downloads and API calls
6. **Google API Errors**: Internal server errors are automatically retried with increasing delays

### Logs

The script provides detailed logging including:
- Download progress with file sizes and percentage completion
- Batch processing updates (1000 rows per batch)
- Filtering statistics (number of records excluded)
- Business/Individual classification results
- Deduplication statistics
- Retry attempts and delays
- Final import counts

## License

[Add your license information here]

## Contributing

[Add contribution guidelines here]
