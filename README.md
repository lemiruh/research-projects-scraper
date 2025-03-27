# Research Projects Scraper

This project contains Python scripts for automating the extraction of academic research project data using Selenium. The tool is designed to navigate a university-funded research portal, perform login authentication, and extract structured data including project details, funding information, and reviewer comments.

## 🔧 Features
- Automated login and multi-page navigation
- Dynamic selection of project year
- Extraction of key project attributes (PI, funding, duration, scores, etc.)
- Collection of qualitative feedback from external reviewers and panels
- Export of results into structured Excel workbooks using `openpyxl` and `xlsxwriter`

## 🧰 Tech Stack
- **Python**
- **Selenium WebDriver**
- **openpyxl** and **xlsxwriter** (for Excel file generation)
- **Git/GitHub** (version control)

## 📁 Project Structure
- `research_projects_crawler.py`: A crawler that scrapes general project data for each year and outputs multiple structured Excel sheets (objectives, outputs, impacts).
- `web_scraping_login.py`: A more advanced crawler that handles login authentication, navigates through multiple browser windows, and extracts reviewer panel comments and evaluation scores.

## 📌 Use Cases
- Academic research portfolio analysis
- Internal review and funding trend analysis
- Educational demos in data engineering and web automation

## 🚀 How to Run
Make sure you have the following installed:
```bash
pip install selenium openpyxl xlsxwriter
