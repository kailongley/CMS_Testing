import os
import subprocess
import time

def run_scripts_in_folder(folder_path, scripts_to_run):
    processes = []
    for script in scripts_to_run:
        script_path = os.path.join(folder_path, f"{script}.py")
        print(f"Running script: {script_path}")
        process = subprocess.Popen(['python', script_path])
        processes.append(process)

        # Add a delay after running tsv_excel
        if script == 'tsv_excel':
            time.sleep(2)  # Adjust the delay time as needed

        # Add a delay after running site_daily_alarms
        elif script == 'site_daily_alarms':
            time.sleep(2)  # Adjust the delay time as needed

    # Wait for all processes to finish
    for process in processes:
        process.wait()

if __name__ == "__main__":
    folder_path = r"C:\Users\klongley\Documents\CMS_Testing\CMS\scripts"
    scripts_to_run = ['tsv_excel', 'excel_master', 'site_daily_alarms', 'pre_graph_creator', 'post_graph_creator']
    
    run_scripts_in_folder(folder_path, scripts_to_run)
