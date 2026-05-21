import os

log_path = r"C:\Users\mvz5cyq\.gemini\antigravity\scratch\outlook-sharepoint-sync\dist\sync.log"
output_path = r"C:\Users\mvz5cyq\.gemini\antigravity\scratch\outlook-sharepoint-sync\log_check_results.txt"

results = []
if os.path.exists(log_path):
    with open(log_path, "r", encoding="utf-8", errors="ignore") as f:
        lines = f.readlines()
    results.append(f"Total lines in log: {len(lines)}")
    
    # Check for warnings, errors, exceptions
    err_lines = []
    for idx, line in enumerate(lines):
        if any(k in line.lower() for k in ["warning", "error", "failed", "exception"]):
            err_lines.append(f"Line {idx+1}: {line.strip()}")
            
    results.append(f"Found {len(err_lines)} warning/error lines:")
    results.extend(err_lines[-50:]) # last 50 error/warning lines
else:
    results.append("Log file does not exist.")

with open(output_path, "w", encoding="utf-8") as f:
    f.write("\n".join(results))

print(f"Results written to {output_path}")
