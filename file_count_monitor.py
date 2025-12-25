import json
import glob
import subprocess
from datetime import datetime
import os


def load_json_data(filepath):
    """
    Load historical log data from a JSON file.
    Returns an empty list if no file exists.
    """
    try:
        with open(filepath, "r") as f:
            return json.load(f)
    except FileNotFoundError:
        return []  # No previous logs yet


def save_json_data(filepath, data):
    """
    Save JSON log data to disk, creating directories if needed.
    """
    os.makedirs(os.path.dirname(filepath), exist_ok=True)
    with open(filepath, "w") as f:
        json.dump(data, f, indent=4)


def log_and_compare(tracked_folder, pattern="*.xlsx", json_file_path="logs/file_counts.json", trigger_script=None, verbose=True, timestamp_logs=True):
    """
    Track changes in the number of files inside a folder.

    This can be used as part of an automated ETL workflow where the
    arrival of new files should trigger downstream processing.

    Behaviour:
    - Counts files matching a pattern in a folder
    - Compares against the last recorded count
    - If the count changes, an optional script is executed
    - The new count is appended to a JSON audit log

    Parameters
    ----------
    tracked_folder : str
        Directory to monitor.
    pattern : str, optional
        Glob pattern (default "*.xlsx").
    json_file_path : str
        Location of JSON audit log.
    trigger_script : str, optional
        Path to a Python script to run when file count changes.
    verbose : bool
        If True, prints detailed status messages.
    timestamp_logs : bool
        If True, stores today's date in the log.
    """

    # --- Determine today's record ---
    today = datetime.now().strftime("%Y-%m-%d") if timestamp_logs else None
    current_file_count = len(glob.glob(os.path.join(tracked_folder, pattern)))

    if verbose:
        print(f"Scanning folder: {tracked_folder}")
        print(f"Pattern: {pattern}")
        print(f"Files found: {current_file_count}")

    # --- Load previous audit trail ---
    data = load_json_data(json_file_path)
    previous_record = data[-1] if data else None
    previous_file_count = previous_record["file_count"] if previous_record else None

    # --- Compare current count with last logged count ---
    if previous_file_count is None:
        if verbose:
            print("No previous log found. Creating initial record...")
    else:
        if current_file_count == previous_file_count:
            if verbose:
                print(f"No change detected (count = {current_file_count}).")
        else:
            if verbose:
                print(
                    f"File count changed: {previous_file_count} → {current_file_count}"
                )
            # Trigger downstream script if supplied
            if trigger_script:
                if verbose:
                    print(f"Executing trigger script: {trigger_script}")
                subprocess.run(["python", trigger_script], check=True)

    # --- Append new audit entry ---
    entry = {"file_count": current_file_count}
    if timestamp_logs:
        entry["date"] = today

    data.append(entry)
    save_json_data(json_file_path, data)

    if verbose:
        if timestamp_logs:
            print(f"Logged: {today}, file_count = {current_file_count}")
        else:
            print(f"Logged file_count = {current_file_count}")

    return entry  # useful for tests or downstream logic
