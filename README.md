# Bitcoin Address Analyzer

A comprehensive Bitcoin address analysis tool that fetches transaction data from the Blockchain.info API and generates detailed reports with 39 calculated parameters as per assignment requirements.

## Features

- **Multi-API Support**: Uses Blockchain.info API with rate limiting and retry mechanisms
- **Comprehensive Analysis**: Calculates 39 parameters for each Bitcoin address
- **Professional Excel Output**: Generates styled reports with multiple sheets
- **Batch Processing**: Analyze multiple addresses in sequence
- **Error Handling**: Robust error handling with fallback options

## How This Program Works

### Architecture Overview
The Bitcoin Address Analyzer follows a modular, multi-layered architecture designed for reliability and maintainability:

1. **API Layer**: Interfaces with Blockchain.info API using rate limiting and retry mechanisms
2. **Data Processing Layer**: Processes raw transaction data and separates inputs/outputs
3. **Calculation Layer**: Computes 39 statistical parameters using mathematical algorithms
4. **Output Layer**: Generates professional Excel reports with multiple formatted sheets

### Data Flow Process
```
Bitcoin Address → API Request → Transaction Data → Data Processing → Parameter Calculation → Excel Report
```

### Step-by-Step Workflow

#### 1. **Address Input & Validation**
- Accepts Bitcoin addresses from user input, file, or predefined list
- Validates address format and API connectivity
- Implements rate limiting to respect API constraints

#### 2. **Data Fetching & Processing**
- **API Integration**: Makes HTTP requests to Blockchain.info API with authentication
- **Transaction Retrieval**: Fetches up to 1000 transactions per address
- **Data Parsing**: Converts JSON responses to structured Python objects
- **Error Handling**: Implements exponential backoff for failed requests

#### 3. **Transaction Analysis**
- **Input/Output Classification**: Separates transactions where the address is sender vs receiver
- **Address Extraction**: Identifies all input and output addresses in each transaction
- **Amount Calculation**: Computes BTC amounts transferred (excluding change addresses)
- **Statistical Processing**: Handles edge cases like empty transaction lists

#### 4. **Parameter Calculation**
The program calculates 39 parameters across two main categories:

**Input Transaction Parameters (1-20):**
- Transaction counts and recipient analysis
- Sender address statistics and distributions
- Coin transfer amounts with statistical measures (mean, std dev, min/max)

**Output Transaction Parameters (21-39):**
- Transaction counts and sender analysis
- Receiver address statistics and distributions
- Coin reception amounts with statistical measures

#### 5. **Report Generation**
- **Multi-Sheet Excel**: Creates Summary, Detailed Analysis, and Parameter Guide sheets
- **Professional Styling**: Applies color coding, borders, and formatting
- **Data Organization**: Optimizes column widths and freezes panes for navigation
- **Fallback Handling**: Gracefully degrades to basic Excel if styling fails

### Technical Implementation Details

#### **Rate Limiting Strategy**
```python
# Configurable delays between requests
'base_delay': 1,           # Base delay between requests (seconds)
'inter_address_delay': 3,  # Delay between different addresses
'max_retries': 10,         # Maximum retry attempts
'timeout': 30              # Request timeout
```

#### **Data Structure Processing**
- **Input Transactions**: Analyzes `vin` (inputs) to identify when address sends coins
- **Output Transactions**: Analyzes `vout` (outputs) to identify when address receives coins
- **Change Detection**: Automatically excludes self-transfers (change addresses)
- **Statistical Calculations**: Uses NumPy for mathematical operations and statistical measures

#### **Error Handling Mechanisms**
- **Network Failures**: Automatic retry with exponential backoff
- **API Rate Limits**: Intelligent delay management
- **Data Validation**: Checks for missing or malformed data
- **Graceful Degradation**: Falls back to basic Excel if advanced features fail

#### **Memory Management**
- **Streaming Processing**: Processes transactions one at a time to minimize memory usage
- **Efficient Data Structures**: Uses sets for unique address counting
- **Batch Processing**: Handles multiple addresses sequentially to avoid overwhelming APIs

### Performance Optimizations

1. **Parallel Processing**: Sequential processing to respect API rate limits
2. **Caching**: Stores API responses to avoid redundant requests
3. **Efficient Algorithms**: Optimized statistical calculations using NumPy
4. **Memory Efficiency**: Processes data in chunks rather than loading everything at once

### Scalability Features

- **Configurable Limits**: Adjustable transaction limits and delays
- **Batch Processing**: Can handle multiple addresses in sequence
- **API Key Management**: Supports multiple API keys for higher rate limits
- **Modular Design**: Easy to extend with additional analysis parameters

## Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd Assignment_2
```

2. Install required dependencies:
```bash
pip install -r requirements.txt
```

3. Set up your Blockchain.info API key in `main.py`:
```python
API_KEY = "your-api-key-here"
```

## Usage

Run the main script:
```bash
python3 main.py
```

Choose from the following options:
1. Analyze all default addresses (10 predefined addresses)
2. Use addresses from `addresses.txt` file
3. Process just one address for testing
4. Exit


## Project Structure

```
Assignment_2/
├── main.py                          # Main analysis script
├── addresses.txt                    # Sample Bitcoin addresses
├── requirements.txt                 # Python dependencies
├── README.md                       # This file
├── screenshots/                    # Screenshot directory
│   ├── main_interface.png
│   ├── analysis_progress.png
│   ├── summary_sheet.png
│   ├── detailed_analysis.png
│   └── parameter_guide.png
└── output/                         # Generated Excel reports
    └── bitcoin_analysis_*.xlsx
```

## API Configuration

The tool uses the Blockchain.info API with the following configuration:
- **Base Delay**: 1 second between requests
- **Inter-Address Delay**: 3 seconds between different addresses
- **Max Retries**: 10 attempts per request
- **Timeout**: 30 seconds per request

## Output Format

The tool generates Excel files with three sheets:

1. **Summary Sheet**: Overall statistics and address overview
2. **Detailed Analysis**: All 39 calculated parameters
3. **Parameter Guide**: Complete parameter definitions

## Parameters Analyzed

### Input Transaction Parameters (1-20)
- Transaction counts and recipient analysis
- Sender address statistics
- Coin transfer amounts and distributions

### Output Transaction Parameters (21-39)
- Transaction counts and sender analysis
- Receiver address statistics
- Coin reception amounts and distributions

## Error Handling

- **Network Errors**: Automatic retry with exponential backoff
- **Rate Limiting**: Built-in delays and retry mechanisms
- **API Failures**: Graceful fallback to basic Excel export
- **File Locking**: Automatic filename generation for locked files

## Dependencies

- `requests`: HTTP requests to Blockchain.info API
- `pandas`: Data manipulation and analysis
- `numpy`: Mathematical calculations
- `openpyxl`: Excel file generation and styling
