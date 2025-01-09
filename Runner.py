import subprocess
import os

current_dir = os.getcwd()

script_name = "pdf_extractor.py"  

if not os.path.exists(os.path.join(current_dir, script_name)):
    print(f"Error: {script_name} not found in the current directory!")
else:
    print(f"Running {script_name}...")
    subprocess.run(["python", script_name], shell=True)