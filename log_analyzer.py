import pandas as pd
from collections import Counter, defaultdict
import json
import os

# Set paths relative to your D:\log analysis folder
DATA_DIR = "data"
OUTPUT_JSON = "compromise_summary.json"

def run_analysis():
    print("Starting Log Analysis...")
    compromised = {}

    # 1. ANALYZE USB EVENTS
    print("\n--- STEP 2: ANALYZING USB LOGS ---")
    usb = pd.read_csv(f"{DATA_DIR}/usb_events.csv")
    print(f"[*] Loaded {len(usb)} USB event records.")
    print("[*] Sample raw data (head):")
    print(usb.head(2))
    print("\n[*] Searching for suspicious vendor ID 'abcd:1234'...")
    for _, r in usb.iterrows():
        # Hunting for the known suspicious USB signature
        if 'abcd:1234' in str(r['message']):
            ep = r['endpoint_id']
            compromised.setdefault(ep, []).append(f"USB: Suspicious mass storage 'abcd:1234' inserted at {r['event_time']}")
            print(f"[!] Suspicious USB event flagged:")
            print(f"    Endpoint ID: {ep}")
            print(f"    Event Time:  {r['event_time']}")
            print(f"    Message:     {r['message']}")
    input("\n[Pause for Screenshot - Step 2] Press Enter to continue...")

    # 2. ANALYZE ENDPOINT SECURITY
    print("\n--- STEP 3: ANALYZING ENDPOINT SECURITY LOGS ---")
    sec = pd.read_csv(f"{DATA_DIR}/endpoint_security.csv")
    print(f"[*] Loaded {len(sec):,} endpoint security logs.")
    print("[*] Sample raw data (head):")
    print(sec.head(2))
    print("\n[*] Filtering for 'FileType: Malicious' and deduplicating...")
    
    # Find quarantined files marked explicitly as 'Malicious'
    malicious_q = sec[sec['message'].str.contains('FileType: Malicious', na=False)].copy()
    malicious_q['file_name'] = malicious_q['message'].str.extract(r'File: (.+?) \|', expand=False)
    malicious_q['file_sha'] = malicious_q['message'].str.extract(r'SHA256: ([a-f0-9]+)', expand=False)
    
    # Deduplicate so we don't report the same file 100 times for one endpoint
    malicious_q_dedup = malicious_q.drop_duplicates(subset=['endpoint_id','file_sha'])
    
    av_count = 0
    print("[*] Sample findings (first 5 endpoints):")
    for ep in malicious_q_dedup['endpoint_id'].unique():
        files = malicious_q_dedup[malicious_q_dedup['endpoint_id']==ep]['file_name'].tolist()
        unique_files = set(files)
        compromised.setdefault(ep, []).append(f"AV: Malicious file(s) quarantined: {', '.join(unique_files)}")
        if av_count < 5:
            print(f"    - {ep} - Files: {', '.join(unique_files)}")
        av_count += 1
        
    if av_count > 5:
        print(f"    ... and {av_count - 5} more endpoints.")
    print(f"\n[+] Total distinct endpoints flagged by AV: {av_count}")
    input("\n[Pause for Screenshot - Step 3] Press Enter to continue...")

    # 3. ANALYZE WEB ACTIVITY (Processed in chunks to save memory)
    print("\n--- STEP 4: ANALYZING WEB ACTIVITY ---")
    print("Analyzing Web Activity for C2 Beaconing (This may take a moment)...")
    example_hits = defaultdict(list)
    
    total_web_rows = 0
    # Read the massive CSV in chunks of 500k rows
    chunk_num = 1
    for chunk in pd.read_csv(f"{DATA_DIR}/web_activity.csv", chunksize=500_000):
        total_web_rows += len(chunk)
        if chunk_num == 1:
            print("[*] Sample raw data (head of chunk 1):")
            print(chunk.head(2))
            print("\n[*] Scanning for '.example' C2 domains...")
            
        print(f"    Processing chunk {chunk_num} ({len(chunk):,} rows)...")
        # Look for simulated C2 domains ending in .example
        example = chunk[chunk['message'].str.endswith('.example', na=False)]
        for _, r in example.iterrows():
            example_hits[r['endpoint_id']].append((r['event_time'], r['message']))
        chunk_num += 1

    print(f"\n[*] Total web activity logs processed: {total_web_rows:,}")
    print(f"[+] Final count: {len(example_hits)} endpoints contacted at least one .example C2 domain.")
    
    print("\n[*] Sample C2 beaconing findings (first 3 endpoints):")
    sample_eps = list(example_hits.keys())[:3]
    for ep in sample_eps:
        sample_domains = list(set([h[1] for h in example_hits[ep]]))[:2]
        print(f"    - {ep} connected to: {', '.join(sample_domains)}")
        
    input("\n[Pause for Screenshot - Step 4] Press Enter to continue...")

    for ep, hits in example_hits.items():
        # Grab up to 3 unique domains they connected to
        domains = list(set([h[1] for h in hits]))
        compromised.setdefault(ep, []).append(f"C2 Beacon: Connected to suspicious domains {domains[:3]}")

    # 4. EVASION DETECTION
    # If an endpoint is in the USB logs but missing from Security logs, the AV was disabled.
    sec_eps = set(sec['endpoint_id'].unique())
    usb_eps = set(usb['endpoint_id'].unique())
    missing_from_sec = usb_eps - sec_eps
    for ep in missing_from_sec:
        compromised.setdefault(ep, []).append("EVASION: Endpoint AV/Logging agent was forcibly disabled.")

    # 5. SAVE RESULTS
    with open(OUTPUT_JSON, 'w', encoding='utf-8') as f:
        json.dump(compromised, f, indent=2, ensure_ascii=False)
    
    print(f"\nAnalysis complete! Flagged {len(compromised)} endpoints.")
    print(f"Raw data saved to {OUTPUT_JSON}")

if __name__ == "__main__":
    run_analysis()