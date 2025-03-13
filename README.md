# Requirement File Generator

Requirement File Generator is a Streamlit-based web application that processes multiple Excel work order files, validates their data, and generates a combined Excel output along with a detailed list of validation errors. The app provides a preview of the results and download options for both the Excel output and a ZIP file containing the output and error logs.

## Features

- **Multi-file Upload:** Upload multiple Excel files (work orders) at once.
- **Data Processing & Validation:** Extracts relevant information from each file, merges rows with child data, and validates key fields.
- **Combined Output:** Generates a combined Excel file with an "OverallRequirement" and "PerOrderRequirement" sheet.
- **Error Reporting:** Provides a list of validation errors (e.g., missing or zero values) that help pinpoint data issues.
- **Preview & Download:** Offers a preview of the combined Excel output and download buttons for both the Excel file and a ZIP archive containing all outputs.
- **Beautiful UI:** Built with Streamlit and enhanced with custom CSS for a clean, modern look.

## Installation

1. **Clone the repository:**

   ```bash
   git clone https://github.com/yourusername/requirement-file-generator.git
   cd requirement-file-generator
   ```

2. **Create a virtual environment (optional but recommended):**

   ```bash
   python -m venv venv
   source venv/bin/activate   # On Windows use: venv\Scripts\activate
   ```

3. **Install dependencies:**

   ```bash
   pip install -r requirements.txt
   ```

   Your `requirements.txt` should include at least:

   ```
   streamlit
   pandas
   openpyxl
   xlsxwriter
   pyarrow
   ```

## Usage

To run the app locally on port 5500:

```bash
streamlit run app.py --server.port 5500
```

Then, open your browser and navigate to:  
[http://127.0.0.1:5500](http://127.0.0.1:5500)

### How It Works

1. **Upload Files:**  
   Select one or more Excel work order files using the file uploader.

2. **Processing:**  
   Click on the "Process Files" button. The app processes the files, validates the data, and combines the outputs.

3. **Preview & Errors:**  
   The app displays a preview of the combined Excel output (from the "OverallRequirement" sheet) along with any validation errors in a scrollable text area.

4. **Download:**  
   Download the combined Excel file or a ZIP archive that contains both the Excel output and a text file with validation errors.

## Deployment

You can deploy this app using one of several free hosting options:

### Streamlit Community Cloud

1. Push your repository to GitHub.
2. Go to [Streamlit Community Cloud](https://share.streamlit.io/) and link your GitHub repository.
3. Deploy the app.

### Heroku

1. Create a `Procfile` with:
   ```
   web: streamlit run app.py --server.port=$PORT --server.enableCORS false
   ```
2. Push your code to Heroku using Git or via GitHub integration.

## Customization

- **Styling:**  
  The app uses custom CSS (injected via `st.markdown`) to improve the look and feel. Feel free to modify the CSS to match your branding.
  
- **Processing Logic:**  
  The backend processing functions can be adjusted as needed for your specific work order file format.

## Contributing

Contributions are welcome! Please fork the repository and submit a pull request with your changes.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
