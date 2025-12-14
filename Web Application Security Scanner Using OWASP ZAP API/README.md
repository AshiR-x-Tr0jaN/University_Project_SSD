# 🛡️ Web Application Security Scanner

Automated vulnerability detection system using OWASP ZAP

## Features
- Automated web application scanning
- Multiple scan types (Quick, Standard, Deep)
- Detailed vulnerability reports
- HTML report generation
- Historical scan data
- Risk categorization

## Installation

1. Install OWASP ZAP from https://www.zaproxy.org/download/
2. Install Python dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Usage

1. Start ZAP on localhost:8080
2. Run the scanner:
   ```bash
   python scanner.py
   ```

## Project Structure
```
SecurityScanner/
│
├── scanner.py              # Main scanner script
├── config.py               # Configuration settings
├── requirements.txt        # Python dependencies
├── scan_results.db         # SQLite database
├── reports/                # Generated reports
└── README.md              # This file
```

## Test Sites
- http://testphp.vulnweb.com
- http://testhtml5.vulnweb.com

## Warning
⚠️ Only scan websites you own or have permission to test!

## License
Educational use only
