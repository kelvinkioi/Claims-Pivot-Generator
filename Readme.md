git add .## Claims Pivot Generator

The Claims Pivot Generator is a Python-based tool designed to process insurance claims data and generate pivot tables for analysis. It reads input data from a CSV file, processes it to calculate key metrics, and outputs a summarized pivot table in Excel format.

### Features
- Reads claims data from a CSV file.
- Generates pivot tables based on specified columns.
- Outputs results in a user-friendly Excel file.
- Supports customization of pivot table fields.

### Usage
1. Place your claims data in a CSV file.
2. Run the `app.py` script.
3. Specify the input file path and desired output file name when prompted.
4. Open the generated Excel file to view the pivot table.

### Requirements
- Python 3.x
- pandas
- openpyxl
- streamlit

### Example
```bash
streamlit run app.py --input claims_data.csv --output summary_pivot.xlsx
```

### License
This project is licensed under the MIT License.## Claims Pivot Generator