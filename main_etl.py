import subprocess
import time
import os


def run_script(script_path):
    """
    Helper to run another Python script and print its stdout/stderr.
    """
    result = subprocess.run(
        ["python", script_path], capture_output=True, text=True
    )
    print(result.stdout)
    print(result.stderr)


def main():
    # Example: open Outlook so COM operations work reliably
    outlook_path = r"C:\Program Files\Microsoft Office\root\Office16\OUTLOOK.EXE"
    if os.path.exists(outlook_path):
        outlook_proc = subprocess.Popen(
            f'start /min "" "{outlook_path}"', shell=True
        )
        print("Opening Outlook...")
        time.sleep(10)
    else:
        outlook_proc = None
        print("Outlook path not found; proceeding without forcing Outlook open.")


    # Adjust names/paths if you use a 'scripts/' subfolder
    email_script          = r"Python Email ETL\email_downloader.py"
    file_monitor_script   = r"Python Email ETL\file_count_monitor.py"
    bed_etl_script        = r"Python Email ETL\bed_report_etl.py"

    print("Running email downloader...")
    run_script(email_script)

    print("Checking file counts and triggering updates if needed...")
    run_script(file_monitor_script)

    print("Running bed report ETL...")
    run_script(bed_etl_script)

    if outlook_proc:
        print("Closing Outlook...")
        outlook_proc.kill()

    print("ETL workflow completed.")


if __name__ == "__main__":
    main()
