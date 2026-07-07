Operational Stats Automation

Project Overview
This project automates the generation of operational dashboards using Python, Excel, and PowerPoint.
It reads KPI data from an Excel file, dynamically updates a PowerPoint template, and generates visual dashboards including gauge charts for performance metrics.

Author
Name: Ramyashree D G
Role: Intern
Company: Systech Solutions

Objectives
- Automate manual reporting process
- Reduce human effort and errors
- Generate consistent dashboards
- Visualize KPI metrics effectively

Technologies Used
- Python
- Pandas
- python-pptx
- Matplotlib
- Excel
- PowerPoint

Project Structure
Operational_Stats_Automation/
│
├── test.py                # Main script
├── job_gauge.py           # Job KPI gauge chart
├── cycle_gauge.py         # Cycle KPI gauge chart
│
├── clean_data.xlsx             # Input data
├── template.pptx              # PPT template
│
├── Output_dashboard.pptx      # Generated dashboard
├── Job gauge.png              # Job gauge output
├── Cycle gauge.png            # Cycle gauge output

How to Run the Project

Step 1: Install Dependencies
pip install pandas python-pptx matplotlib openpyxl

Step 2: Run the Script
python test.py

Features
- Automatic data extraction from Excel
- Dynamic PowerPoint text replacement
- KPI visualization using gauge charts
- Image insertion into slides
- Fully automated dashboard generation

Output
The script generates:
- PowerPoint Dashboard (Output_dashboard.pptx)
- Job KPI Gauge Chart (Job gauge.png)
- Cycle KPI Gauge Chart (Cycle gauge.png)

Testing & Validation
- Verified with sample input data
- Ensured correct KPI mapping
- Validated chart generation
- Checked PowerPoint output formatting

Challenges Faced
- Handling text replacement in PPT
- Designing gauge charts dynamically
- Aligning data with template placeholders

Solutions Implemented
- Used structured key matching for PPT updates
- Created reusable gauge chart scripts
- Ensed proper data formatting before processing

Key Learnings
- Automation using Python
- Working with PowerPoint via code
- Data visualization techniques
- Debugging real-time issues

Future Improvements
- Add more chart types
- Improve UI design of dashboards
- Integrate real-time data sources
- Convert script into web application
