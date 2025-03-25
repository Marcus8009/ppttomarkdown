# PowerPoint to Markdown Converter

This repository contains a simple [Streamlit](https://streamlit.io/) application that converts PowerPoint files (`.ppt` and `.pptx`) into Markdown format. The app scans a specified folder for any PowerPoint files, converts their slides into Markdown, and saves the result in a newly created `markdown_output` folder.

---

## Features

- Convert multiple PowerPoint files (`.ppt` or `.pptx`) at once.
- Extracts all text from each slide and formats it in Markdown.
- Saves Markdown files in a separate output folder for easy organization.

---

## Requirements

- Python 3.7 or above
- Required Python libraries:
  - [streamlit](https://pypi.org/project/streamlit/)
  - [python-pptx](https://pypi.org/project/python-pptx/)

---

## Installation

1. **Clone the repository** or download the `app.py` file.
2. **Create a virtual environment** (recommended):
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Linux or macOS
   .\venv\Scripts\activate   # On Windows
