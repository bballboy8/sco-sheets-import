# SCO Sheets Import

A Python script that downloads property records from the California State Controller's Office (SCO) and imports them into a Google Sheets document. The script handles deduplication, batch processing, and provides progress tracking during the import process.

## Features

- Downloads property records from multiple ZIP files from the SCO website
- Imports CSV data into Google Sheets with automatic deduplication
- Batch processing for efficient API usage
- Progress tracking and logging
- Configurable batch sizes and API rate limiting
- Automatic header detection and column mapping

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
- `API_PAUSE_SEC`: Pause between API calls in seconds (default: 0.8)
- `MAX_RECORD_ROWS`: Maximum number of rows to import (default: 180,000)

### Data Sources

The script downloads data from the following SCO URLs:
- 01_From_0_To_Below_10.zip
- 02_From_10_To_Below_100.zip  
- 03_From_100_To_Below_500.zip
- 04_From_500_To_Beyond.zip

## How It Works

1. **Authentication**: Uses Google Service Account credentials to authenticate with Google Sheets API
2. **Download**: Downloads ZIP files from the SCO website containing CSV property records
3. **Processing**: Extracts and processes CSV files from the ZIP archives
4. **Deduplication**: Identifies existing records using the "Property ID" column to avoid duplicates
5. **Import**: Appends new records to the Google Sheets document in batches
6. **Progress Tracking**: Provides detailed logging of download progress and import statistics

## Output

The script creates or updates a "Records" worksheet in your Google Sheets document with:
- Property records from the SCO data
- Automatic header row creation
- Deduplication based on Property ID
- Batch processing for efficient API usage

## Troubleshooting

### Common Issues

1. **Authentication Errors**: Ensure your service account has the correct permissions and the JSON credentials are properly formatted
2. **Sheet Access**: Make sure the service account email has edit access to your Google Sheets document
3. **API Quotas**: The script includes rate limiting, but you may need to adjust `API_PAUSE_SEC` if you hit quota limits
4. **Memory Issues**: For very large datasets, consider reducing `APPEND_BATCH_SIZE` or `MAX_RECORD_ROWS`

### Logs

The script provides detailed logging including:
- Download progress with file sizes
- Batch processing updates
- Deduplication statistics
- Final import counts

## License

[Add your license information here]

## Contributing

[Add contribution guidelines here]
