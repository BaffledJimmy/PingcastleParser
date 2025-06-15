import os
import json
import openai
import time
import subprocess
import requests
import re
from datetime import datetime

# --- Configuration ---
SECRETS_PATH = "C:\\Scripts\\OpenAISecrets.sec"
REPORT_DIR   = r"C:\scripts\PingCastle\reports"
LOG_DIR      = r"C:\Scripts\Logs"

# --- Ensure log directory exists ---
os.makedirs(LOG_DIR, exist_ok=True)

# --- Load OpenAI secrets from DPAPI-encrypted file using PowerShell ---
cmd = [
    "powershell", "-NoProfile", "-NonInteractive", "-Command",
    f"""
    $secure = Get-Content '{SECRETS_PATH}' | ConvertTo-SecureString;
    $json = [Runtime.InteropServices.Marshal]::PtrToStringAuto(
        [Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    );
    Write-Output $json
    """
]

try:
    decrypted_json = subprocess.check_output(cmd, stderr=subprocess.DEVNULL, text=True).strip()
    secrets = json.loads(decrypted_json)
except Exception as e:
    print(f"Failed to load secrets: {e}")
    exit(1)

openai.api_key = secrets["OPENAI_API_KEY"]
assistant_id   = secrets["ASSISTANT_ID"]
thread_id      = secrets["THREAD_ID"]
webhook_url    = secrets["ANALYSIS_WEBHOOK"]

# --- Find the 4 most recent HTML reports ---
html_files = sorted(
    [f for f in os.listdir(REPORT_DIR) if f.endswith(".html")],
    key=lambda f: os.path.getmtime(os.path.join(REPORT_DIR, f)),
    reverse=True
)

if not html_files:
    print("No PingCastle HTML reports found.")
    exit(1)

if len(html_files) < 4:
    print(f"Warning: Only {len(html_files)} reports found. Forecasting may be limited.")

recent_files = html_files[:4]
file_ids = []

for filename in recent_files:
    full_path = os.path.join(REPORT_DIR, filename)
    print(f"Uploading: {full_path}")
    try:
        with open(full_path, "rb") as f:
            uploaded = openai.files.create(file=f, purpose="assistants")
            file_ids.append(uploaded.id)
    except Exception as e:
        print(f"Failed to upload {filename}: {e}")
        exit(1)

# --- Create message with context ---
try:
    openai.beta.threads.messages.create(
        thread_id=thread_id,
        role="user",
        content=(
            "Here are the 4 most recent PingCastle HTML reports. "
            "Please compare them to identify changes in risk posture and forecast what might impact the score in upcoming weeks. "
            "Filenames may not reflect exact weeks — multiple reports may be from the same day. "
            "The security team already receives deltas separately; focus on long-term hygiene trends and risk predictions."
        ),
        attachments=[{
            "file_id": fid,
            "tools": [{"type": "file_search"}]
        } for fid in file_ids]
    )
except Exception as e:
    print(f"Failed to create message: {e}")
    exit(1)

# --- Start and poll assistant run ---
try:
    run = openai.beta.threads.runs.create(thread_id=thread_id, assistant_id=assistant_id)
except Exception as e:
    print(f"Run start failed: {e}")
    exit(1)

print("Waiting for assistant to respond...")
while True:
    status = openai.beta.threads.runs.retrieve(thread_id=thread_id, run_id=run.id)
    if status.status in ["completed", "failed", "cancelled"]:
        break
    time.sleep(1)

if status.status != "completed":
    print(f"Assistant run failed: {status.status}")
    exit(1)

# --- Get assistant response ---
messages = openai.beta.threads.messages.list(thread_id=thread_id)
reply = next((m.content[0].text.value for m in reversed(messages.data) if m.role == "assistant"), None)

def format_for_slack(raw):
    # Remove OpenAI-style citations
    cleaned = re.sub(r"【\d+:\d+†source】", "", raw)

    # Convert markdown-style bold (**text**) to Slack-style (*text*)
    cleaned = re.sub(r"\*\*(.+?)\*\*", r"*\1*", cleaned)

    # Convert leading hyphens to bullets
    cleaned = re.sub(r"(?m)^- ", "• ", cleaned)
    return cleaned.strip()

if reply:
    formatted_reply = format_for_slack(reply)
    payload = {"text": f"*PingCastle AI Analysis*\n\n{formatted_reply}"}

    # --- Send to webhook ---
    try:
        r = requests.post(webhook_url, json=payload)
        r.raise_for_status()
        print("Webhook sent.")
    except Exception as e:
        print(f"Failed to send webhook: {e}")

    # --- Write to .log file ---
    date_str = datetime.now().strftime("%Y%m%d-%H%M%S")
    log_path = os.path.join(LOG_DIR, f"PingCastleAI-{date_str}.log")
    with open(log_path, "w", encoding="utf-8") as f:
        f.write(formatted_reply)
    print(f"Saved response to: {log_path}")

else:
    print("No assistant reply found.")
