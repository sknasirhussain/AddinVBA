import os
import tempfile
import requests
from flask import Flask, render_template, request
from flask_cors import CORS
from werkzeug.utils import secure_filename
from win32com.client import Dispatch
import pythoncom
from xlsx2html import xlsx2html
from flask import Blueprint

file_io_url = ''

app2 = Blueprint('view_form', __name__)
CORS(app2)

API_KEY = 'VUQ2N3O.P64G1AX-7QN4GP8-KVZZMBH-WK0S1XX'
UPLOAD_URL = 'https://file.io/'

@app2.route('/stocks', methods=['GET', 'POST'])
def view_form():
    print(request.files)
    if request.method == 'POST':
        # Example usage
        addin_path = "C:/Users/SkNasirHussain/pursuit software development pvt. ltd/Sujit Sarkar - MACROS/Add_Ins_Sheet.xlam"
        module_name = "StockPrice"
        macro_name = "GetStockDetails"
        title = "Stocks Result"
        
        pythoncom.CoInitialize()
        excel = Dispatch("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = True

        uploaded_file = request.files['file']
        if uploaded_file:
            filename = secure_filename(uploaded_file.filename)

            headers = {
                'Accept': 'application/json',
                'Authorization': 'Bearer ' + API_KEY,
            }

            files = {'file': (filename, uploaded_file)}
            response = requests.post(UPLOAD_URL, headers=headers, files=files)
            if response.status_code == 200:
                file_io_url = response.json()['link']

                download_response = requests.get(file_io_url)
                if download_response.status_code == 200:
                    temp = tempfile.gettempdir()
                    local_file_path = os.path.join(temp, filename)

                    with open(local_file_path, 'wb') as f:
                        f.write(download_response.content)

                    workbook = excel.Workbooks.Open(local_file_path)
                    excel.Application.Run("'" + addin_path + "'!" + module_name + "." + macro_name)

                    temp = tempfile.gettempdir()
                    file_path = os.path.join(temp, "alpha.xlsm")

                    workbook.SaveAs(file_path)
                    workbook.Close()
                    excel.Quit()

                    if file_path:
                        temp_path = "C:/Users/SkNasirHussain/Desktop/copy/Templates"
                        file_path_html = os.path.join(temp_path, "Stock.html")

                        xlsx2html(file_path, file_path_html)

                        with open(file_path_html, 'r') as html_file:
                            html_content = html_file.read()

                        html_content = html_content.replace('<head>', f'<head><title>{title}</title>')

                        with open(file_path_html, 'w') as html_file:
                            html_file.write(html_content)

                        return render_template('Stock.html', html_content=html_content)
                    else:
                        return "No file path"

                    if file_path:

                        # preparing to upload the updated file o file.io servers

                        with open(file_path, 'rb') as file:
                            response = requests.post(url, headers, files={'file': file })
                            print(f"Status code from server is {response.status_code}")

                            if response.status_code == 200:
                                print("File upload successful")

                                # File successfully uploaded to File.io

                                file_io_url = response.json()['link']
                                print(response.json())

                            else:
                                print("Error in uploading the file")
                                print(f"Status code: {response.status_code}")
                                print(f"Message: {response.text}")

                            # file downloading starts
                            print("Attempting to download file")
                            download_response = requests.get(file_io_url)

                            if download_response.status_code == 200:

                                with tempfile.NamedTemporaryFile(delete=False) as temp_file:
                                    
                                    temp_file.write(response.content)
                                    temp_file.flush()
                                    temp_file_path = temp_file.name
                                                                        
                                # Retrieve original filename from URL
                                file_name = os.path.basename(file_url) + ".xlsx"
                                
                                # Prompt user for download path
                                download_path = 'C:/Users/SkNasirHussain/Desktop/copy/Templates' #D:/Loan/LoanTemplate/Stock_portfolio_api/Templates
                                file_path = ""

                                if download_path:                                
                                    file_path = os.path.join(download_path, file_name)
                                else:
                                    file_path = os.path.join(os.getcwd(), file_name)
                                                                   
                                # Move temporary file to the desired download path
                                file_path1 = 'C:/Users/SKNASI~1/AppData/Local/Temp' + file_path
                                shutil.move(temp_file_path, file_path)

                                print(f"File downloaded successfully to: {file_path}")
                                
                            else:
                                raise Exception(f"Failed to retrieve file from {file_io_url}. Status Code: {download_response.status_code}")
                    
                    else:
                        print("No file created.")
                else:
                    return "File download failed. Status code:", download_response.status_code
            else:
                return response.json()

    return "Post method failed"

if __name__ == '__main__':
    app.run(debug=True)
