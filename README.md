# Financial Data Mapping Tools

This repository contains Python scripts for mapping financial data and converting templates.

### Acquisition_Reporting_tabs.py

- **Purpose**: Maps Proforma financial data from the Reporting tab.
- **Functionality**: Processes 12 months of income statement data.

### PriorTemplate_toNew.py

- **Purpose**: Converts old financial templates to a new format.
- **Functionality**: Requires user input to change paths and structure.

### Unacquired_Loads.py
- **Purpose**: Populate the database with 1000's of deals that PACS did not close on
- **Functionality**: An algorithm to match deal id's to files for the upload process

#### Installation

1. Clone the repository:

    ```bash
    git clone https://github.com/yourusername/financial-data-mapping-tools.git
    ```

2. Install the required libraries:

    ```
    pip install -r requirements.txt
    ```

## Usage

### Acquisition_Reporting_tabs.py

1. Run the script:
    ```
    python Acquisition_Reporting_tabs.py
    ```

### PriorTemplate_toNew.py

1. Run the script and follow the prompts:
    ```
    python PriorTemplate_toNew.py
    ```

### Unacquired_Loads.py

1. Run the script and monitor the matching:
    ```
    python Unacquired_Loads.py
    ```
2. Load the json file of the schema connection
