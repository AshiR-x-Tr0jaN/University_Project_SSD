from scanner import SecurityScanner
import json

# List of websites to scan
websites = [
    'http://testphp.vulnweb.com',
    'http://testhtml5.vulnweb.com',
    'http://testasp.vulnweb.com'
]

scanner = SecurityScanner()

results = []

for url in websites:
    print(f"\n{'='*60}")
    print(f"Scanning: {url}")
    print('='*60)
    
    scan_id = scanner.start_scan(url, 'quick')
    
    if scan_id:
        scanner.print_summary(scan_id)
        scanner.generate_report(scan_id, f'report_{scan_id}.html')
        results.append({
            'url': url,
            'scan_id': scan_id,
            'status': 'completed'
        })
    else:
        results.append({
            'url': url,
            'scan_id': None,
            'status': 'failed'
        })

# Save batch results
with open('batch_results.json', 'w') as f:
    json.dump(results, indent=2, fp=f)

print("\n✅ Batch scan completed!")
print(f"Total sites scanned: {len(results)}")
print(f"Check batch_results.json for summary")
