# 📊 AI Excel Analyst

**AI Excel Analyst** is an interactive Excel data analysis application built with **Python, Pandas, and Streamlit**.

Upload an Excel file and analyze your data using natural-language questions. The application is designed to work with different types of Excel datasets and automatically identify useful columns, statistics, patterns, relationships, and trends.

Instead of manually writing formulas or creating pivot tables, you can simply ask questions about your data and get meaningful answers.

## ✨ What It Can Do

AI Excel Analyst can help you explore and understand Excel data by performing tasks such as:

* Analyze rows and columns
* Identify numeric and categorical data
* Calculate totals and averages
* Find minimum and maximum values
* Count records and unique values
* Analyze categories and groups
* Compare different groups
* Identify highest and lowest performing categories
* Analyze trends over time
* Calculate percentages and rates
* Analyze missing values
* Identify duplicate records
* Generate calculated columns
* Perform grouped calculations
* Analyze relationships between columns
* Create charts and visualizations
* Answer natural-language questions
* Generate data-driven insights
* Summarize important patterns in the dataset

## 💬 Ask Questions in Natural Language

You don't need to know Python, Pandas, SQL, or Excel formulas.

For example:

```text
What is the total sales?
```

```text
Which category has the highest revenue?
```

```text
What is the average value?
```

```text
Which product was sold the most?
```

```text
Show me the top 10 customers.
```

```text
Which month had the highest sales?
```

```text
Are there any missing values?
```

```text
Find duplicate records.
```

```text
Show a chart of sales by category.
```

You can ask questions based on the columns and information available in your own Excel file.

## 🧮 Calculated Columns

The application can also work with calculated fields.

For example:

```text
Create a calculated column Revenue = Quantity × Price
```

This allows additional metrics to be created from existing columns and then used for further analysis.

## 📊 Data Analysis

The application can analyze different types of Excel information, including:

### Numeric Data

* Sum
* Average
* Minimum
* Maximum
* Count
* Distribution
* Comparisons

### Categorical Data

* Category counts
* Most common values
* Group comparisons
* Aggregations by category

### Date and Time Data

* Daily trends
* Monthly trends
* Yearly trends
* Highest and lowest periods
* Time-based comparisons

### Data Quality

* Missing values
* Duplicate rows
* Empty columns
* Basic dataset structure

## 📈 Visualizations

The application can create visual representations of the data when appropriate.

Examples include:

* Bar charts
* Horizontal bar charts
* Line charts
* Category comparisons
* Time-series charts

For example:

```text
Create a horizontal bar chart of revenue by category.
```

## 📁 Excel File Support

The application is designed to analyze Excel datasets with different structures.

It does not depend on one specific dataset or one specific business category.

You can upload an Excel workbook containing your own data and ask questions based on the available columns.

## 🚀 How to Use

1. Open the application.
2. Upload an Excel file.
3. Let the application analyze the dataset.
4. Ask a question in natural language.
5. Review the calculated result, analysis, or visualization.

## 🛠️ Technology

Built with:

* **Python**
* **Streamlit**
* **Pandas**
* **NumPy**
* **OpenPyXL**
* **Matplotlib**

## 💻 Run Locally

Clone the repository:

```bash
git clone https://github.com/Pythonsnake2024/AI-Excel-Analyst.git
```

Open the project directory:

```bash
cd AI-Excel-Analyst
```

Install the required packages:

```bash
pip install -r requirements.txt
```

Run the application:

```bash
streamlit run app.py
```

The application will open in your web browser.

## 📂 Project Structure

```text
AI-Excel-Analyst/
│
├── app.py
├── requirements.txt
└── README.md
```

## ☁️ Deployment

The application can be deployed using **Streamlit Community Cloud**.

After deployment, users can access the application through a web browser without installing Python or any additional software.

## 🔐 Data

Do not upload confidential or sensitive Excel files to the public GitHub repository.

Excel files should be uploaded through the application for analysis rather than committed to the repository.

## 🎯 Goal

The goal of AI Excel Analyst is to make Excel data analysis easier and more accessible.

**Upload your data → Ask a question → Get an answer.**

---

### 🔗 Repository

https://github.com/Pythonsnake2024/AI-Excel-Analyst
