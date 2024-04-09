import os
import tempfile
from win32com.client import Dispatch
import qrcode
import PIL
from flask import Flask,render_template, redirect, send_file
from flask_cors import CORS
import pythoncom
from datetime import datetime
import requests
from xlsx2html import xlsx2html
from flask import url_for
import shutil
from flask import Blueprint

api = 'VUQ2N3O.P64G1AX-7QN4GP8-KVZZMBH-WK0S1XX'
 
url= 'https://file.io/'
 
headers = {
    'Accept' :  'application/json',
    'Authorization': 'Bearer ' + api,
    'Content-Type' : 'multipart/form-data'
}
 
title = "Loan Amortization Schedule"
 
def download_file(file_url):
    print("25-->",file_url)
    # return redirect(file_url)
    # Send a GET request to downlo:ad the file
    response = requests.get(file_url)
    print('28-->',response.status_code)
    # Check if request was successful
    if response.status_code == 200:
        # Create a temporary file
        with tempfile.NamedTemporaryFile(delete=False) as temp_file:
           
            temp_file.write(response.content)
            temp_file.flush()
            temp_file_path = temp_file.name
          
        # Retrieve original filename from URL
        file_name = os.path.basename(file_url) + ".xlsx"
       
        # Prompt user for download path
        download_path = "C:/Users/SkNasirHussain/pursuit software development pvt. ltd/Sujit Sarkar - EXCELGEN"
        file_path = ""
        if download_path:
           
            file_path = os.path.join(download_path, file_name)
        else:
            file_path = os.path.join(os.getcwd(), file_name)
        print('48-->',file_path, file_name)
        # Move temporary file to the desired download path
        file_path1 = 'C:/Users/SKNASI~1/AppData/Local/Temp' + file_path
        shutil.move(temp_file_path, file_path)
        print(f"File downloaded successfully to: {file_path}")
        return send_file(file_path, as_attachment=True)
    else:
        # Handle failed request
        print(f"Failed to download file. Status code: {response.status_code}")
 
app1 = Blueprint('calculate', __name__)
CORS(app1)
 
 
@app1.route('/calculate_loan/<float:amount>/<int:tenure>/<int:payments>')
def calculate(amount, tenure, payments):
 
    addin_path = r"C:/Users/SkNasirHussain/pursuit software development pvt. ltd/Sujit Sarkar - MACROS/Add_Ins.xlam"
    module_name = "NewLoanTemplate"
    macro_name = "SubCalculate"
    args = (amount, tenure, payments)
 
    print("Code is running...")
 
    try:
        pythoncom.CoInitialize()
        excel = Dispatch("Excel.Application")
        excel.DisplayAlerts = False
        excel.Visible = True  
 
        workbook = excel.Workbooks.Add()
        sheet = workbook.ActiveSheet
 
        excel.Application.Run("'" + addin_path + "'!" + module_name + "." + macro_name, *args)
 
        result = "Data fetched successfully"

        target_sheet = workbook.Worksheets("Sheet1 (2)")
        target_sheet.Activate()

        temp = tempfile.gettempdir()
 
        excelgen = "C:/Users/SkNasirHussain/pursuit software development pvt. ltd/Sujit Sarkar - EXCELGEN"
 
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_name = f"Loan_{timestamp}.xlsx"
       
        file_path = os.path.join(temp, file_name)

         #Generating QR code
        img = qrcode.make("https://pursuitsoftwarebiz-my.sharepoint.com/personal/sujit_s_pursuitsoftware_biz/_layouts/15/onedrive.aspx?id=%2Fpersonal%2Fsujit%5Fs%5Fpursuitsoftware%5Fbiz%2FDocuments%2FEXCELGEN&ct=1712576087622&or=Teams%2DHL&ga=1/" + file_name + "\n\n\n" + "Version 1")
        type(img)  # qrcode.image.pil.PilImage

        # Desired size in inches
        desired_size = (1.5, 1.5)

        # Convert inches to pixels at 72 DPI
        desired_width = int(desired_size[0] * 72)
        desired_height = int(desired_size[1] * 72)

        # Resize the image
        img = img.resize((desired_width, desired_height))


        qr_img_path = os.path.join(temp, "qr.png")
        img.save(qr_img_path)   

        # Insert QR code image into the worksheet
        img_width, img_height = img.size

        # Calculate position to insert the image at the bottom right corner
        
        range_B1 = target_sheet.Range("B1")

        # Get the left and top positions of the range "N2"
        qr_left = range_B1.Left
        qr_top = range_B1.Top

        target_sheet.Shapes.AddPicture(
        qr_img_path, 
        LinkToFile=False, 
        SaveWithDocument=True, 
        Left=qr_left, Top=qr_top, 
        Width=img_width, 
        Height=img_height
        )

        # Saving the Excel Work

        workbook.SaveAs(file_path)
        workbook.Close()
        excel.DisplayAlerts = True
 
        excel.DisplayAlerts = True
              
        excel.Quit()
        del excel
 
        pythoncom.CoInitialize()
 
        if file_path:
            print("Generating HTML file...")

            temp_path = 'C:/Users/SkNasirHussain/Desktop/AddinVBA/main/Templates'
            file_path_html = os.path.join(temp_path, "Loan.html")
           
            xlsx2html(file_path, file_path_html)

            print("Path of the html is: ",file_path_html)
           
            with open(file_path_html, 'r') as html_file:
                html_content = html_file.read()
 
            # Insert the title tag with the desired title into the HTML content
            html_content = html_content.replace('<head>', f'<head><title>{title}</title>')
           
            # Write the modified HTML content back to the HTML file
            with open(file_path_html, 'w') as html_file:
                html_file.write(html_content)
               
                download_path = "C:/Users/SkNasirHussain/pursuit software development pvt. ltd/Sujit Sarkar - EXCELGEN"
                
                if download_path:
                 
                    down_file_path = os.path.join(download_path, file_name)
                else:
                    down_file_path = os.path.join(os.getcwd(), file_name)
                
                # Move temporary file to the desired download path

                print("tring to store at ", down_file_path)
                shutil.move(file_path, down_file_path)
                print("File succcessfully stored to excelgen at: ", down_file_path, "!")
                
                # return render_template('Loan.html',html_content=html_content)
               
                # print("Attempting to upload file to File.io...")
                # # Upload the Excel file to File.io
                # with open(file_path, 'rb') as file:
                #     response = requests.post(url, headers, files={'file': file })
                #     if response.status_code == 200:
                #         print("File upload successful")
                #         # File successfully uploaded to File.io
                #         file_io_url = response.json()['link']
                #         print(response.json())
                #         download_file(file_io_url)

                #         #print(file_io_url)
                #         # Redirect the user to file_io_url

                #         return render_template('Loan.html', html_content=html_content)
                #         # return redirect(file_io_url)
                #         # encoded_file_url = urlsafe_b64encode(file_io_url.encode()).decode()
                #         # return redirect(url_for('download_file', file_url=encoded_file_url))    
 
                #     else:
                #         print('Upload failed', response.status_code)
                       
        else:
                print("no file path")
   
        return render_template('Loan.html', html_content=html_content)
 
    except IOError:
 
        print("Error",IOError )
        return "error occured"
 
if __name__ == "__main__":
    app.run(debug = True)