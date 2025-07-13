import subprocess
import sys
import asyncio
from datetime import datetime
from pathlib import Path
import os
import socket


# Auto install required packages
def install(package):
    subprocess.check_call([sys.executable, "-m", "pip", "install", package])

# Ensure required modules are installed
try:
    from dotenv import load_dotenv
except ImportError:
    install("python-dotenv")
    from dotenv import load_dotenv

try:
    import pandas as pd
except ImportError:
    install("pandas")
    import pandas as pd

try:
    from playwright.async_api import async_playwright
except ImportError:
    install("playwright")
    from playwright.async_api import async_playwright
    subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=True)

try:
    import openpyxl
except ImportError:
    install("openpyxl")
    import openpyxl

try:
    import requests
except ImportError:
    install("requests")
    import requests

try:
    import platform
    import psutil
except ImportError:
    install("psutil")
    import platform
    import psutil


from pandas import ExcelWriter

# Constants
TARGET_TEXT = "安否確認へのご回答ありがとうございます。"

async def run_test(context, url):
    result = "Failed"
    message = ""
    duration = 0
    console_errors = []

    try:
        page = await context.new_page()
        page.on("console", lambda msg: console_errors.append(msg.text) if msg.type == "error" else None)

        start_time = datetime.now()
        await page.goto(url, timeout=300_000)

        await page.click('input[name="SafetySituation"][value="1"]')
        await page.click('(//input[@type="radio"])[4]')
        await page.fill('//*[@name="Comment"]', ' test comment')
        await page.click('button[type="submit"]')

        try:
            await page.wait_for_selector(f"text={TARGET_TEXT}", timeout=300_000)
            result = "Success"
            message = TARGET_TEXT
        except:
            message = "Success text not found."

        duration = round((datetime.now() - start_time).total_seconds(), 2)
        await page.close()

    except Exception as e:
        message = f"{type(e).__name__}: {str(e)[:100]}"

    return {
        "url": url,
        "duration_sec": duration,
        "result": result,
        "message": message,
        "console_errors": "; ".join(console_errors)[:200]
    }


def send_summary_to_slack_webhook(summary_df, webhook_url):
    # Gather PC info
    pc_info = {
        "Hostname": socket.gethostname(),
        "Processor": platform.processor() or "N/A",
        "CPU Count": psutil.cpu_count(logical=True),
        "RAM (GB)": round(psutil.virtual_memory().total / (1024 ** 3), 2),
        "OS": f"{platform.system()} {platform.release()}",
    }

    # Format PC info text
    pc_info_text = "*🖥️ PC Details:*\n" + "\n".join([f"*{k}:* {v}" for k, v in pc_info.items()])

    # Format summary text (first row only)
    summary_text = "*📊 Test Summary*\n\n"
    summary_text += "\n".join([f"*{col}:* {val}" for col, val in summary_df.iloc[0].items()])

    # Combine summary + PC info
    full_message = f"{summary_text}\n\n{pc_info_text}"

    payload = {
        "text": full_message
    }

    response = requests.post(webhook_url, json=payload)

    if response.ok:
        print("✅ Summary + PC details sent to Slack via webhook!")
    else:
        print("❌ Failed to send summary:", response.text)

async def main():
    # Automatically pick the first CSV file from URL folder
    url_folder = Path("URL")
    csv_files = sorted(url_folder.glob("*.csv"))

    if not csv_files:
        print("❌ No CSV files found in 'URL' folder. Please add at least one CSV file.")
        return

    input_file = csv_files[0]
    print(f"📂 Using input CSV file: {input_file}")

    try:
        max_tabs = int(input("Enter max number of concurrent tabs: "))
    except ValueError:
        print("❌ Invalid number. Please enter a valid integer.")
        return

    start_time = datetime.now()
    results_dir = Path("Results")
    results_dir.mkdir(exist_ok=True)

    df = pd.read_csv(input_file)
    urls = df["URL"].dropna().tolist()
    total = len(urls)
    completed = 0
    lock = asyncio.Lock()

    # Build meaningful filename
    timestamp = start_time.strftime("%Y-%m-%d__%H-%M-%S")
    file_name = f"Results__tabs-{max_tabs}__urls-{total}__{timestamp}.xlsx"
    excel_output_file = results_dir / file_name

    print(f"🚀 Test started at: {start_time.strftime('%Y-%m-%d %H:%M:%S')} | Total URLs: {total}")

    async with async_playwright() as pw:
        browser = await pw.chromium.launch(headless=True)
        context = await browser.new_context(viewport={"width": 1920, "height": 1080})
        sem = asyncio.Semaphore(max_tabs)

        async def bounded_run(url):
            nonlocal completed
            async with sem:
                result = await run_test(context, url)
                async with lock:
                    completed += 1
                    print(f"✅ [{completed}/{total}] {url}")
                return result

        results = await asyncio.gather(*(bounded_run(url) for url in urls))
        await context.close()
        await browser.close()

    end_time = datetime.now()
    total_exec_time = round(((end_time - start_time).total_seconds())/60, 2)

    results_df = pd.DataFrame(results)

    # Summary DataFrame with execution time
    summary_data = {
        "Total URLs": [total],
        "Number of Tabs Used": [max_tabs],
        "Success Count": [results_df["result"].value_counts().get("Success", 0)],
        "Failure Count": [results_df["result"].value_counts().get("Failed", 0)],
        "Max Duration (sec)": [results_df["duration_sec"].max()],
        "Min Duration (sec)": [results_df[results_df["duration_sec"] > 0]["duration_sec"].min()],
        "Average Duration (sec)": [results_df["duration_sec"].mean()],
        "Total Execution Time (sec)": [total_exec_time]
    }
    summary_df = pd.DataFrame(summary_data)

    # Save both result and summary in Excel
    with ExcelWriter(excel_output_file, engine='openpyxl') as writer:
        results_df.to_excel(writer, sheet_name="Test Results", index=False)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

    print(f"\n✅ Results saved to: {excel_output_file}")

    print("\n======= 📊 Summary =========")
    print(summary_df.to_string(index=False))

    print("\n🕒 Started at:", start_time.strftime("%Y-%m-%d %H:%M:%S"))
    print("✅ Finished at:", end_time.strftime("%Y-%m-%d %H:%M:%S"))
    print(f"⏱ Total time: {total_exec_time / 60:.2f} minutes")

    # sending the details  to slack
    # Load .env variables
    load_dotenv()

    # Get the webhook URL
    webhook_url = f"https://hooks.slack.com/services/{os.getenv("PART_1")}/{os.getenv("PART_2")}/{os.getenv("PART_3")}"
    send_summary_to_slack_webhook(summary_df, webhook_url)


if __name__ == "__main__":
    asyncio.run(main())
