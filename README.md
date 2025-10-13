# 🎈 Streamlit App

🔄 T2 WS Transformation App

A Streamlit web app that automates the transformation of raw and unstructured sales Excel files into a standardized reporting format used across multiple outlets.
The app supports 40+ customer-specific templates, each with unique file structures, sheet names, and data layouts.

It was built to eliminate manual Excel reformatting, reduce data errors, and accelerate monthly reporting for regional sales teams.

🧩 Overview

Sales teams often receive Excel files from different outlets — each with its own structure, headers, and formatting.
This app automatically recognizes the correct outlet format, applies the right transformation logic, and outputs a clean, Power BI-ready dataset.

🔹 Core Functions

Upload a raw Excel sales file (any outlet format)

Select the correct transformation template (from 40+ customer-specific rules)

Upload an optional mapping file for SKU and size conversion

Generate a standardized output with consistent columns and data types

Download the transformed Excel ready for integration into reporting systems

🖥️ How It Works

File Upload:
Users drag and drop raw outlet data and (optionally) a mapping file into the app.

Template Selection:
A dropdown allows users to select the outlet code (e.g. 30010008 酒甕, 30010013 酒田, 30010315 圳程, etc.).
Each code triggers a custom transformation pipeline written in Python.

Transformation Engine:
The app runs pre-defined logic per outlet, including:

Sheet selection

Header detection and row trimming

SKU name normalization and size conversion

Date and volume formatting (9L conversion using size lookup)

Column renaming to a unified schema

Output Generation:
Once processed, the cleaned dataset is displayed in-app and can be downloaded as an Excel file ready for Power BI or further analysis.

🧰 Tech Stack
Component	Description
Python	Core data transformation logic
Pandas / OpenPyXL	Excel processing and data cleaning
Streamlit	Interactive web interface for upload, selection, and output
Excel / Power BI	Final data destination and visualization layer

<img width="1074" height="590" alt="Screenshot 2025-10-13 at 11 06 08" src="https://github.com/user-attachments/assets/d3e571f5-d7fd-4b13-9d7d-d781127b4bd7" />
