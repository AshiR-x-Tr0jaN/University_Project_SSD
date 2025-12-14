ZAP_CONFIG = {
    'api_key': 'changeme',
    'proxy_url': 'http://127.0.0.1:8080',
    'timeout': 300,
    'max_depth': 5
}

DATABASE_CONFIG = {
    'db_name': 'scan_results.db'
}

SCAN_TYPES = {
    'quick': {
        'name': 'Quick Scan',
        'duration': '5-10 minutes',
        'includes': ['Spider', 'Passive Scan']
    },
    'standard': {
        'name': 'Standard Scan', 
        'duration': '15-30 minutes',
        'includes': ['Spider', 'Passive Scan', 'Active Scan']
    },
    'deep': {
        'name': 'Deep Scan',
        'duration': '1-2 hours',
        'includes': ['Spider', 'Passive Scan', 'Active Scan', 'Full Policy']
    }
}


# ==========================================
# .gitignore
# ==========================================

# Database
*.db
*.sqlite3

# Python
__pycache__/
*.py[cod]
*$py.class
*.so
.Python
env/
venv/
ENV/

# Reports
*.html
*.pdf
reports/

# IDE
.vscode/
.idea/
*.swp
*.swo

# OS
.DS_Store
Thumbs.db

# Logs
*.log

