from flask import Flask, request, send_file, render_template
import pandas as pd
import zipfile
import io
import logging

# Initialize Flask app
app = Flask(__name__)

# Setup logging
logging.basicConfig(level=logging.INFO)

# Function to read and combine sheet data
def combine_sheet_across_files(sheet_name, file):
    try:
        df = pd.read_excel(file, sheet_name=sheet_name)
        df_cleaned = df[~df.iloc[:, 0].str.contains('Name', na=False)]
        return df_cleaned
    except ValueError:
        logging.error(f"Sheet {sheet_name} not found.")
        return pd.DataFrame()

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        try:
            # Read the uploaded ZIP file in memory
            uploaded_file = request.files['file']
            zip_file = io.BytesIO(uploaded_file.read())

            # Extract files from the ZIP in memory
            with zipfile.ZipFile(zip_file, 'r') as zip_ref:
                merged_data = {}  # Dictionary to hold DataFrames for each sheet
                for file_name in zip_ref.namelist():
                    if file_name.endswith('.xlsx'):
                        logging.info(f"Processing file: {file_name}")
                        with zip_ref.open(file_name) as file:
                            xls = pd.ExcelFile(file)
                            sheet_names = xls.sheet_names

                            for sheet_name in sheet_names:
                                if sheet_name not in merged_data:
                                    merged_data[sheet_name] = []
                                df_sheet = combine_sheet_across_files(sheet_name, file)
                                if not df_sheet.empty:
                                    merged_data[sheet_name].append(df_sheet)

            # Save the merged data to a new Excel file in memory
            output_file = io.BytesIO()
            with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
                for sheet_name, df_list in merged_data.items():
                    final_df = pd.concat(df_list, ignore_index=True)
                    final_df.to_excel(writer, sheet_name=sheet_name, index=False)

            output_file.seek(0)  # Reset the buffer to the beginning

            logging.info("File created and ready to download.")

            # Send the merged Excel file as a download
            return send_file(output_file, as_attachment=True, download_name='merged_output.xlsx',
                             mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

        except Exception as e:
            logging.error(f"Error processing file: {e}")
            return "An error occurred during processing", 500

    return render_template('upload_zip.html')

if __name__ == '__main__':
    app.run(debug=True)
