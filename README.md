# Excel File Processor - Streamlit App

A simple and intuitive Streamlit application that allows you to upload Excel files, process them, and download the processed results.

## Features

- 📁 **File Upload**: Upload Excel files (.xlsx or .xls)
- ⚙️ **Data Processing**: Automatically processes your data with customizable logic
- 📊 **Data Visualization**: View data statistics, column information, and previews
- 💾 **Download Results**: Download processed Excel files with multiple sheets
- 🔍 **Data Analysis**: Built-in tools for missing data analysis and statistics

## What the App Does

The app processes your Excel data by:
1. Adding a timestamp column (`Processed_At`)
2. Adding a row number column (`Row_Number`)
3. Adding calculated columns for numeric data (doubles the first numeric column)
4. Creating a summary sheet with processing information

## Installation

### 🚀 Quick Setup (Recommended)

**For macOS/Linux:**
```bash
./setup.sh
```

**For Windows:**
```batch
setup.bat
```

These scripts will automatically:
- Create a virtual environment
- Install all dependencies
- Set up everything you need

### 📋 Manual Setup (Alternative)

1. **Clone or download this repository**

2. **Create a virtual environment:**
   ```bash
   python -m venv venv
   ```

3. **Activate the virtual environment:**
   
   **On macOS/Linux:**
   ```bash
   source venv/bin/activate
   ```
   
   **On Windows:**
   ```bash
   venv\Scripts\activate
   ```

4. **Install the required dependencies:**
   ```bash
   pip install -r requirements.txt
   ```

## Usage

### 🚀 Quick Start

**For macOS/Linux:**
```bash
./run_app.sh
```

**For Windows:**
```batch
run_app.bat
```

### 📋 Manual Run

1. **Activate the virtual environment** (if not already active):
   
   **On macOS/Linux:**
   ```bash
   source venv/bin/activate
   ```
   
   **On Windows:**
   ```bash
   venv\Scripts\activate
   ```

2. **Run the Streamlit app:**
   ```bash
   streamlit run app.py
   ```

3. **Open your browser** and navigate to the URL shown in the terminal (usually `http://localhost:8501`)

4. **Upload an Excel file** using the sidebar file uploader

5. **View the processed data** and download the result

## Customization

You can customize the data processing logic by modifying the `process_excel_data()` function in `app.py`. This function currently:
- Adds timestamp and row number columns
- Creates calculated columns for numeric data
- You can add your own processing logic here

## File Structure

```
PPA/
├── app.py              # Main Streamlit application
├── requirements.txt    # Python dependencies
├── README.md          # This file
└── [your Excel files] # Your input files
```

## Dependencies

- **streamlit**: Web application framework
- **pandas**: Data manipulation and analysis
- **openpyxl**: Excel file reading/writing
- **xlrd**: Legacy Excel file support

## Example

1. Upload an Excel file with columns like: Name, Age, Salary, Department
2. The app will add: Processed_At, Row_Number, and Salary_Doubled columns
3. Download the processed file with both the processed data and a summary sheet

## Troubleshooting

- **File upload issues**: Make sure your Excel file is not corrupted and is in .xlsx or .xls format
- **Memory issues**: For very large files, consider processing smaller chunks
- **Dependencies**: Make sure all required packages are installed using `pip install -r requirements.txt`

## License

This project is open source and available under the MIT License.
