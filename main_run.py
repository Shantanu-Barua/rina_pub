### This is the main curve fitting code file ###

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
import re
from io import BytesIO
import win32com.client
import requests
from urllib.parse import urlparse, unquote

### Folders
input_folder = 'source'
output_folder = 'final'
tags_folder = 'tags'

## Seperate source folder for soring raw files
src1 = 'source/single'
src2 = 'source/multiple'
src3 = 'source/multi_plot_pos'

## Seperate outputs into different folders
otpt1 = 'final/single'
otpt2 = 'final/multiple'
otpt3 = 'final/multi_plot_pos'

### Make required folders
os.makedirs(output_folder, exist_ok=True)
os.makedirs(input_folder, exist_ok=True)
os.makedirs(tags_folder, exist_ok=True)
os.makedirs(src1, exist_ok=True)
os.makedirs(src2, exist_ok=True)
os.makedirs(src3, exist_ok=True)
os.makedirs(otpt1, exist_ok=True)
os.makedirs(otpt2, exist_ok=True)
os.makedirs(otpt3, exist_ok=True)

### Download curve fitting files from gihub
## Define functions
def get_filename_from_url(url):
    """Extracts filename from a GitHub raw URL, ignoring query parameters."""
    parsed_url = urlparse(url)
    filename = os.path.basename(parsed_url.path)
    return unquote(filename)

def download_file(url, save_dir):
    """Downloads file from a URL and saves it in the specified directory."""
    filename = get_filename_from_url(url)
    file_path = os.path.join(save_dir, filename)

    try:
        response = requests.get(url)
        response.raise_for_status()  # raise HTTPError for bad responses

        with open(file_path, 'wb') as f:
            f.write(response.content)
        print(f"⬇️ Downloaded: {filename}")
        return file_path

    except requests.exceptions.RequestException as e:
        print(f"❌ Failed to download {filename}: {e}")
        return None

def download_if_missing(url_list, save_dir='.'):
    """Checks for and downloads files that don't already exist."""
    if not os.path.exists(save_dir):
        os.makedirs(save_dir)

    for url in url_list:
        filename = get_filename_from_url(url)
        file_path = os.path.join(save_dir, filename)

        if os.path.exists(file_path):
            print(f"✅ Already exists: {filename}")
        else:
            download_file(url, save_dir)

### Required curve fitting code files
github_raw_urls = [
    "https://raw.githubusercontent.com/Shantanu-Barua/rina/refs/heads/main/curve_fit_single_plot_v1.py?token=GHSAT0AAAAAADGSGUD2KBRGDXAAE4PE22R62DZTMSQ",
    "https://raw.githubusercontent.com/Shantanu-Barua/rina/refs/heads/main/curve_fit_single_plot_std_temp_v1.py?token=GHSAT0AAAAAADGSGUD2YOGVDNNKFZFVO7GI2DZTNCQ",
    "https://raw.githubusercontent.com/Shantanu-Barua/rina/refs/heads/main/curve_fit_multi_plot_v1.py?token=GHSAT0AAAAAADGSGUD2WG5NTTHHPNOVWFCC2DZTNSQ",
    "https://raw.githubusercontent.com/Shantanu-Barua/rina/refs/heads/main/curve_fit_multi_plot_std_temp_v1.py?token=GHSAT0AAAAAADGSGUD3JQVG5GEIRPZBO4522DZTOSQ",
    "https://raw.githubusercontent.com/Shantanu-Barua/rina/refs/heads/main/curve_fit_multi_plot_multi_pos_v1.py?token=GHSAT0AAAAAADGSGUD3AOR2HW4DUODS4QGQ2DZTPAA",
    "https://raw.githubusercontent.com/Shantanu-Barua/rina/refs/heads/main/curve_fit_multi_plot_multi_pos_std_temp_v1.py?token=GHSAT0AAAAAADGSGUD3OHAPVM6V3QYBB3M22DZTPJQ"
    
]

### Download the files if missing
download_if_missing(github_raw_urls)