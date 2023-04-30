from flask import Flask, request, render_template, session, Response
from flask_cors import CORS, cross_origin
import pandas as pd
import os
import uuid
import io
import datetime

app = Flask(__name__)
cors = CORS(app)
app.config['CORS_HEADERS'] = 'Content-Type'
app.secret_key = str(uuid.uuid4())

# Define file paths for input and output files
MYDIR = os.path.dirname(__file__)
excelFilePath = os.path.join(MYDIR, 'data/input.xlsx')
csvFilePath = os.path.join(MYDIR, 'data/input.csv')
outputFilePath = os.path.join(MYDIR, 'data/output.xlsx')

today = datetime.date.today().strftime('%m-%d-%Y')

def processInventory(excel_data, csv_data):

    output = io.BytesIO()

    # logic begins
    excel_data = excel_data.iloc[:-1 , :]
    new_header = excel_data.iloc[1] #grab the first row for the header
    excel_data = excel_data[2:] #take the data less the header row
    excel_data.columns = new_header #set the header row as the df header

    excel_data["Warehouse ID"] = excel_data["Warehouse ID"].astype(str).astype(int)

    excel_data["Barcode"] = excel_data["Barcode"].astype('string')

    csv_data = csv_data[(csv_data["UPC Online"] == True) & (csv_data["Product Online"] == True)]

    csv_data["UPC"] = csv_data["UPC"].astype('string')

    excel_data_selected_styles = excel_data[excel_data['Design Season ID'] == 'PF23']

    excel_data = pd.merge(excel_data, csv_data, left_on = "Barcode", right_on = "UPC", how = "inner" )

    excel_data = excel_data.append(excel_data_selected_styles, ignore_index = True)

    df_10500 = excel_data[excel_data["Warehouse ID"] == 10500]

    df_10500 = df_10500[['Barcode','Style Nbr','Item Nbr','Design Season ID','Division Desc','Category Desc','MJ Division','Group Desc','Item Name','Color Name','Size','US Retail Price',
                        'Available Physical','Total Incoming Supply','Total Demand (On Order)','Total Available','ATS Today']]

    df_10501 = excel_data[excel_data["Warehouse ID"] == 10501]

    df_10501 = df_10501[['Barcode','Available Physical','Total Incoming Supply','Total Demand (On Order)','Total Available','ATS Today']]

    if len(df_10500) > len(df_10501):
        df_merge = pd.merge(df_10500, df_10501, on = "Barcode", how = "left" )
    else:
        df_merge = pd.merge(df_10501, df_10500, on = "Barcode", how = "left" )

    df_10501_low =df_merge[(df_merge["Available Physical_y"] <= 10) & (df_merge["Available Physical_y"] > 0)]
    df_10501_zero =df_merge[df_merge["Available Physical_y"] == 0]

    df_merge = df_merge.rename ( columns = {
        'Available Physical_x':'Available Physical',
        'Total Incoming Supply_x':'Total Incoming Supply',
        'Total Demand (On Order)_x':'Total Demand (On Order)',
        'Total Available_x': 'Total Available',
        'ATS Today_x':'ATS Today',

        'Available Physical_y':'Available Physical',
        'Total Incoming Supply_y':'Total Incoming Supply',
        'Total Demand (On Order)_y':'Total Demand (On Order)',
        'Total Available_y': 'Total Available',
        'ATS Today_y':'ATS Today'
       })

    df_merge['Barcode'] = df_merge['Barcode'].astype(float)
    
    # logic ends

    # Write the output data to an Excel file
    with pd.ExcelWriter(output) as writer:
        # Add your script logic here to write the output data to the Excel file
        
        bold_format = writer.book.add_format({'align': 'center', 'valign': 'vcenter','bold': 'True', 'border':1})
        cell_format = writer.book.add_format({'align': 'center', 'valign': 'vcenter'})

        df_merge.to_excel(writer, sheet_name="Inventory", startrow = 1, startcol = 0, index = False)
        df_10501_low.to_excel(writer, sheet_name="Low 10501 Inventory", startrow = 1, startcol = 0, index = False)
        df_10501_zero.to_excel(writer, sheet_name="Zero 10501 Inventory", startrow = 1, startcol = 0, index=False)

        worksheet = writer.sheets["Inventory"]
        worksheet.merge_range('M1:Q1', '10500', bold_format)
        worksheet.merge_range('R1:V1', '10501', bold_format)
        worksheet.set_column('A:E', 18, cell_format)
        worksheet.set_column('F:V', 22, cell_format)
        worksheet.freeze_panes(2, 1)

        worksheet2 = writer.sheets["Low 10501 Inventory"]
        worksheet2.merge_range('M1:Q1', '10500', bold_format)
        worksheet2.merge_range('R1:V1', '10501', bold_format)
        worksheet2.set_column('A:E', 18, cell_format)
        worksheet2.set_column('F:V', 22, cell_format)
        worksheet2.freeze_panes(2, 1)

        worksheet3 = writer.sheets["Zero 10501 Inventory"]
        worksheet3.merge_range('M1:Q1', '10500', bold_format)
        worksheet3.merge_range('R1:V1', '10501', bold_format)
        worksheet3.set_column('A:E', 18, cell_format)
        worksheet3.set_column('F:V', 22, cell_format)
        worksheet3.freeze_panes(2, 1)
        
        # excel_data.to_excel(writer, sheet_name="Inventory", startrow = 1, startcol = 0, index = False)
    
    return output

@app.route("/")
def home():
    return render_template('index.html')

@app.route("/process", methods=["POST"])
def process():
    if request.method != "POST":
        return ("Failed to download files")
    # Get the input files from the HTML form
    excel_file = request.files["uploaded-file-1"]
    csv_file = request.files["uploaded-file-2"]

    if excel_file and csv_file:

        # Read the input files using pandas
        excel_data = pd.read_excel(excel_file)
        csv_data = pd.read_csv(csv_file)

        # Add your script logic here to process the input data
        output = processInventory(excel_data, csv_data)

        # # Return the output file as a response to the HTTP request
        # with open(outputFilePath, 'rb') as f:
        #     output_data = io.BytesIO(f.read())
        buffer = output.getvalue()

        headers = {
            'Content-Disposition': f'attachment; filename= 10500 & 10501 Inventory {today}.xlsx',
            'Content-Type': 'application/vnd.ms-excel'
        }
        return Response(buffer, mimetype='application/vnd.ms-excel', headers=headers)

if __name__ == "__main__":
    app.run(debug=True)



