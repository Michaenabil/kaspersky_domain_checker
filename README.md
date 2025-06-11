# Kaspersky Domain Checker

## Overview

This script allows you to analyze domains using the [Kaspersky OpenTIP API](https://opentip.kaspersky.com) by querying security reputation and category information. It processes domains from an Excel input file and saves the results to a new Excel file.

## Features

- Queries Kaspersky OpenTIP API for each domain
- Extracts category labels and reputation zone
- Handles missing/empty/invalid domain entries
- Tracks and prints real-time processing progress
- Saves results in Excel format with status indicators

## Requirements

- Python 3.x
- `requests`, `pandas` libraries (`pip install requests pandas openpyxl`)
- Kaspersky OpenTIP API Key (you must request it [from here](https://opentip.kaspersky.com))

## Usage

### Command-Line Format

```bash
python kaspersky_domain_checker.py input_file.xlsx output_file.xlsx <API_TOKEN>
