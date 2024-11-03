
<h1 align="center">
  <br>
  Passport Entry Automation for Travel Agencies
  <br>
</h1>

<h4 align="center">A simple automation programme that checks duplicates for 2 excel files</h4>

This program was created to help a travel agency client automate the process of verifying passport entries between two Excel files. The client previously documented each passport number manually when booking tickets, then checked entries against a system file to ensure no errors. This program automates that process, eliminating the need for manual comparison.

## Features

1. **Missing Entries Detection**  
   Checks for passport numbers in the first Excel file that do not appear in the second. This helps the client verify if any passport numbers were missed when entering them into the system.

2. **Erroneous Entries Detection**  
   Identifies passport numbers in the second Excel file that do not exist in the first, allowing the client to detect any mistakenly entered passports.

3. **Duplicate Detection**  
   Scans both files for duplicate passport numbers, helping to avoid accidental double bookings and financial loss.

## Usage

1. Run the program.
2. Enter the path to the first Excel file.
3. Enter the path to the second Excel file.
4. The program will output:
   - A list of missing passport entries (present in the first file but not in the second).
   - A list of erroneous entries (present in the second file but not in the first).
   - Any duplicate passport entries found in either file.

This automation streamlines the verification process, allowing the client to focus on other essential tasks.

## Requirements

- Python 3.x
- Libraries:
  - `pandas`
  - `openpyxl`

## Installation

1. Clone this repository.
   ```bash
   git clone https://github.com/your-username/passport-entry-automation.git

