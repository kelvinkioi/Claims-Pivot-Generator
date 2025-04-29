# Claims Pivot Generator

The Claims Pivot Generator is a Python-based tool designed to process insurance claims data and generate pivot tables for analysis. It reads input data from a CSV file, processes it to calculate key metrics, and outputs a summarized pivot table in Excel format. Additionally, you can now access the tool via a web interface.

### Features
- Reads claims data from a CSV file.
- Generates pivot tables based on specified columns.
- Outputs results in a user-friendly Excel file.
- Supports customization of pivot table fields.
- Accessible via a web-based interface.

### Usage
1. Place your claims data in a CSV file.
2. Run the `app.py` script locally or use the web interface.
3. For local usage:
    - Specify the input file path and desired output file name when prompted.
    - Open the generated Excel file to view the pivot table.
4. For web usage:
    - Visit the [Claims Pivot Generator Web App](https://claims-pivot-generator-mjkp6vy9vjm35jpwyzpu5d.streamlit.app/).
    - Upload your CSV file and download the generated pivot table.

### Requirements
- Python 3.x (for local usage)
- pandas
- openpyxl
- streamlit

### Example (Local Usage)
```bash
streamlit run app.py 
```

### License
This project is licensed under the MIT License.# Claims Pivot Generator