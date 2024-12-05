import os
import subprocess
from pathlib import Path
import nicegui
import shutil

"""
This file can be used to create an executable version of the code by just running 

py build.py

Along with the executable there are several folders that will need to be packaged alongside it to function properly:

static                                      - stores the css file for nicegui
supplementary_files                         - the intended location for the nltk_data folder and any matrices needed, you will need to provide these
temp                                        - stores a temporary copy of any uploaded documents

"""

# Paths for the folders to be created or copied
current_directory = Path(__file__).parent
static_src = current_directory / 'static'
supplementary_files_src = current_directory / 'supplementary_files'
temp_dir = current_directory / 'dist' / 'temp'

# Create the dist/temp directory
if not temp_dir.exists():
    temp_dir.mkdir(parents=True)

# Copy static and supplementary_files directories
shutil.copytree(static_src, current_directory / 'dist/static', dirs_exist_ok=True)
shutil.copytree(supplementary_files_src, current_directory / 'dist/supplementary_files', dirs_exist_ok=True)

# PyInstaller command
cmd = [
    'python',
    '-m', 'PyInstaller',
    'main.py',
    '--name', 'AI_Contract_Scanner',
    '--onefile',
    # '--windowed',  # prevents console from appearing
    '--add-data', f'{Path(nicegui.__file__).parent}{os.pathsep}nicegui'
]
subprocess.call(cmd)
