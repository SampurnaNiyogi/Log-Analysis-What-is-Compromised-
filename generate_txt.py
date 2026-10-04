import json
import os

JSON_PATH = "compromise_summary.json"
TXT_PATH = "compromise_report.txt"

def generate_txt():
    if not os.path.exists(JSON_PATH):
        print(f"Error: {JSON_PATH} not found. Run log_analyzer.py first.")
        return

    with open(JSON_PATH, 'r', encoding='utf-8') as f:
        data = json.load(f)

    with open(TXT_PATH, 'w', encoding='utf-8') as f:
        f.write("=========================================================\n")
        f.write("      COMPROMISED ENDPOINTS SUMMARY REPORT\n")
        f.write("=========================================================\n\n")
        f.write(f"Total Compromised Endpoints Flagged: {len(data)}\n\n")
        
        # Sort endpoints alphabetically
        for endpoint, alerts in sorted(data.items()):
            f.write(f"Endpoint ID: [{endpoint}]\n")
            for alert in alerts:
                f.write(f"  [!] {alert}\n")
            f.write("-" * 60 + "\n")

    print(f"Success! Readable text summary generated at: {TXT_PATH}")

if __name__ == "__main__":
    generate_txt()