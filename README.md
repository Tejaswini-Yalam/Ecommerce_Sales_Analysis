# E-Commerce Sales & Customer Behavior Analysis

## 1. Project Overview

This project is an interactive Streamlit dashboard for analyzing e-commerce sales and customer behavior. It generates a synthetic sales dataset, intentionally introduces common data-quality issues, cleans the data, performs analysis, displays interactive visualizations, and can generate a PowerPoint report.

## 2. Objectives

- Generate a realistic e-commerce sales dataset.
- Demonstrate data cleaning and preprocessing.
- Analyze revenue by product category and region.
- Analyze customer-segment performance.
- Study monthly revenue trends.
- Examine delivery performance.
- Identify the top products by revenue.
- Provide interactive filtering through Streamlit.
- Generate a downloadable PowerPoint summary report.

## 3. Technologies and Libraries

- Python
- Pandas
- NumPy
- Streamlit
- Plotly
- Matplotlib
- python-pptx

The complete dependency list is provided in `requirements.txt`.

## 4. Dataset

The dataset used in this project was generated and prepared by me specifically for the E-Commerce Sales & Customer Behavior Analysis project. The application creates the following datasets automatically:

- `data/ecommerce_sales_raw.csv` — the generated raw dataset.
- `data/ecommerce_sales_clean.csv` — the cleaned and processed dataset used for analysis.

An external dataset is not required to run the application.

### Optional Reference Dataset

For users who wish to explore a real-world e-commerce dataset, the project can also be extended or compared with the **Brazilian E-Commerce Public Dataset by Olist**, available on Kaggle:

**Dataset Link:**  
https://www.kaggle.com/datasets/olistbr/brazilian-ecommerce

The Olist dataset contains approximately 100,000 orders from Brazilian marketplaces and includes information related to orders, products, customers, payments, sellers, and reviews. It is provided by Olist and hosted on Kaggle. Download the dataset before using it. 

> **Note:** The Olist dataset is an optional external reference dataset and is not required for the current version of this project.

The generated dataset contains 600 original rows plus deliberately added duplicate rows and data-quality issues for demonstrating cleaning techniques.

## 5. Project Workflow

1. Data generation
2. Data cleaning
3. KPI calculation
4. Category, region and segment analysis
5. Monthly revenue analysis
6. Delivery-performance analysis
7. Interactive Streamlit dashboard
8. PowerPoint report generation

## 6. Dashboard Features

The dashboard provides filters for:
- Region
- Product Category
- Customer Segment
- Order Date Range

It displays:
- Total Orders
- Total Revenue
- Average Order Value
- Average Customer Rating
- Cancelled Rate
- Revenue by Product Category
- Revenue by Region
- Monthly Revenue Trend
- Customer Segment Performance
- Top 10 Products by Revenue
- Delivery Performance
- Filtered raw data

## 7. How to Run

### Step 1: Open the Project Folder

Open the project folder in VS Code/BobIDE and open the integrated terminal.

Make sure the terminal is inside the project folder.

Example:

```powershell
cd "C:\Users\yalam\New folder"
```

### Step 2: Create a Virtual Environment

Create a Python virtual environment named `venv`:

```powershell
python -m venv venv
```

This creates a `venv/` folder inside the project.


### Step 3: Activate the Virtual Environment

On Windows PowerShell:

```powershell
.\venv\Scripts\Activate.ps1
```

After successful activation, the terminal should show something similar to:

```text
(venv) PS C:\Users\yalam\New folder>
```

> **If PowerShell shows an error saying that script execution is disabled**, activation can be skipped. Use the virtual environment's Python directly as shown in Step 4.

### Step 4: Install Project Dependencies

If the virtual environment is activated:

```powershell
pip install -r requirements.txt
```

If the virtual environment cannot be activated, use:

```powershell
.\venv\Scripts\python.exe -m pip install -r requirements.txt
```

The `requirements.txt` file contains all the Python packages required to run this project.

### Step 5: Run the Streamlit Application

If the virtual environment is activated:

```powershell
streamlit run Tejaswini_Ecommerce_Sales_Analysis.py
```

If the virtual environment is not activated:

```powershell
.\venv\Scripts\python.exe -m streamlit run Tejaswini_Ecommerce_Sales_Analysis.py
```

### Step 6: Open the Application

After running the Streamlit application, the terminal will display a local URL, usually:

```text
http://localhost:8501
```

Open this URL in a web browser to view the E-Commerce Sales & Customer Behavior Analysis dashboard.

### Important Notes

- Python 3.12 or a compatible Python version is required.
- `requirements.txt` contains all required project dependencies.
- The main application file is `Tejaswini_Ecommerce_Sales_Analysis.py`.
- When the project is downloaded to a new computer, create a new `venv/` using Step 2 and install the dependencies using Step 4.

## 8. Project Files

- `Tejaswini_Ecommerce_Sales_Analysis.py` — complete Python source code.
- `Tejaswini_Ecommerce_Sales_Analysis.ipynb` — notebook version of the code.
- `requirements.txt` — required Python libraries.
- `Tejaswini_Ecommerce_Sales_Analysis_ProjectReport.docx` — project documentation.
- `README.md` — project overview and setup instructions.

## 9. Expected Output

The application produces an interactive e-commerce analytics dashboard and allows the user to generate and download a PowerPoint report containing key findings and visualizations.

## 10. Author

**Name:** Tejaswini Yalam  
**Project:** E-Commerce Sales & Customer Behavior Analysis
