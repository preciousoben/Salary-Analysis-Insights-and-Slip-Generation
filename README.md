# **Salary Insights: In-Depth Exploratory Data Analysis and Payslip Generation**

This project dives into a **company's salary dataset for December 2016**, analyzing employee compensation structures and uncovering trends in salary distributions, gender representation, and designations. The dataset consists of key variables such as employee **names, gender, base salary, net salary, total deductions, and additions to salary**, which give a comprehensive view of how the final net salary is calculated.  

Through **Exploratory Data Analysis (EDA)**, we explore various relationships, including **gender vs net salary** and **gender vs designation**, providing insights into how compensation differs across different roles and genders. Additionally, this project offers tools for **automated payslip generation**—both an **interactive payslip using Power BI** and a **printable payslip generated through Python in a Jupyter Notebook**.

---

## **Project Structure & Highlights**

- **Technologies Used:** Python, Jupyter Notebook, KNIME Analytics, Power BI  
- **Dataset:** Employee salary data, including base and net salaries, deductions, and other compensation variables  
- **Skills Demonstrated:**  
  - Data cleaning and preparation  
  - EDA to uncover trends in salary structures  
  - Correlation analysis between compensation variables  
  - Workflow automation using KNIME for reproducibility 
  - Payslip generation (Power BI and Python)  

---

## **How to Explore the Project**

1. **Interactive Jupyter Notebook**:  
   Click [here](https://colab.research.google.com/drive/120W5rmAZKVUSCKUArmuSvpsb9OkYjV5j?usp=sharing) to access the Jupyter Notebook on **Google Colab**.  

2. **Power BI Dashboard**:  
   Explore salary insights and generate **interactive payslips** with Power BI [here](https://app.powerbi.com/reportEmbed?reportId=95eb6265-e87d-41be-9876-a74a0b5765f9&autoAuth=true&ctid=23d0348e-2962-4f09-9203-398c135660be).  

3. **Python Script for Payslip Generation**:  
   The complete Python code, including **printable payslip generation**, can be found in the `Salary and worker analysis and slip generation.py` file.  

4. **KNIME Workflow**:  
   Access the **KNIME workflow** in this repository to explore automated data processing.  

---

## **Project Workflow**  

1. **Data Cleaning & Transformation**  
   - Addressed missing or inconsistent data
   - Removed duplicate records
   - Converted text columns like gender and designation into numerical values for analysis  
   - Ensured data consistency to facilitate seamless correlation analysis  

2. **Exploratory Data Analysis (EDA)**  
   - **Analyzed gender representation** across different designations  
   - **Compared net salary distribution by gender** to identify trends  
   - Used **column charts** to visualize gender vs designation  
   - Created **box plots** to compare base salaries across designations  

3. **Correlation Analysis**  
   - Converted categorical variables into numeric representations  
   - Generated a **correlation matrix** to investigate relationships among base salary, net salary, and other variables  
   - Visualized correlations using **heatmaps** to identify significant patterns  

4. **Payslip Generation**  
   - **Power BI:** Designed an interactive payslip template within the dashboard  
   - **Python:** Automated **printable payslip generation** in Jupyter Notebook, making it easy to export and share  

5. **Workflow Automation with KNIME**  
   - Built an automated workflow to streamline data processing and ensure efficient analysis for future updates  

---

## **Repository Content**  

| File/Link                            | Description                                      |
|--------------------------------------|--------------------------------------------------|
| `Salary and worker analysis and slip generation.py`         | Python script with EDA and payslip generation    |
| [Jupyter Notebook](https://colab.research.google.com/drive/120W5rmAZKVUSCKUArmuSvpsb9OkYjV5j?usp=sharing)                | Viewable notebook on Google Colab                |
| [Power BI Dashboard](https://app.powerbi.com/reportEmbed?reportId=95eb6265-e87d-41be-9876-a74a0b5765f9&autoAuth=true&ctid=23d0348e-2962-4f09-9203-398c135660be)              | Interactive dashboard with payslip generation    |
| `Salary_worker_analysis.knwf`    | KNIME workflow for automated data processing     |

---

## **How to Run the Code Locally**  

1. **Clone this repository**:  
   ```bash  
   git clone https://github.com/preciousoben/Salary-Analysis-Insights-and-Slip-Generation.git  
   cd Salary-Analysis-Insights-and-Slip-Generation 
   ```  

2. **Install the required libraries**:  
   ```bash  
   pip install -r requirements.txt  
   ```  

3. **Run the Python script**:  
   ```bash  
   python Salary_Insights_Project.py  
   ```  

---

