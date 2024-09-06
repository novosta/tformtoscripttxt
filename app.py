from flask import Flask, render_template, request, send_file
import pandas as pd
import os

app = Flask(__name__)

# Define the folder where uploaded files will be stored temporarily
UPLOAD_FOLDER = 'uploads/'
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# Route for the home page where users upload the file
@app.route('/')
def upload_file():
    return render_template('upload.html')

# Route for handling the file upload and conversion
@app.route('/convert', methods=['POST'])
def convert_file():
    if 'file' not in request.files:
        return 'No file uploaded!', 400
    
    file = request.files['file']
    
    if file.filename == '':
        return 'No selected file!', 400

    # Save the uploaded file
    file_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(file_path)

    # Convert the Excel file to script format
    output_file_path = os.path.join(UPLOAD_FOLDER, 'output_script_fixed.txt')
    convert_tform_to_script(file_path, output_file_path)

    # Send the output file back to the user
    return send_file(output_file_path, as_attachment=True)

# Function to convert T-Form Excel to script, with complete `(nan)` removal
def convert_tform_to_script(input_xlsx, output_txt):
    # Load the Excel file
    xlsx_data = pd.read_excel(input_xlsx)
    
    # Open the output file
    with open(output_txt, 'w') as script_file:
        for index, row in xlsx_data.iterrows():
            what_seen = str(row['What is seen on Camera']).strip()
            dialogue = str(row['Dialogue']).strip()
            
            # Skip rows where both columns are empty or contain 'nan'
            if not what_seen or what_seen.lower() == 'nan':
                what_seen = ''
            if not dialogue or dialogue.lower() == 'nan':
                dialogue = ''

            # Only write if there's content in either field
            if what_seen and dialogue:
                script_file.write(f"({what_seen}) {dialogue}\n\n")
            elif what_seen:
                script_file.write(f"({what_seen})\n\n")
            elif dialogue:
                script_file.write(f"{dialogue}\n\n")

    print(f"Script successfully written to {output_txt}")

if __name__ == '__main__':
    app.run(debug=True)
