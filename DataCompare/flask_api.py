import pandas as pd
import os
import pathlib
import datetime as dt
from openpyxl import Workbook, load_workbook
from excel_compare import compare
import function_def as func
from flask import Flask , request , send_file
from flask_cors import CORS, cross_origin

app = Flask(__name__)
CORS(app, support_credentials=True)


@app.route('/upload' , methods=["POST"])
@cross_origin(supports_credentials=True)
def compareRestApi():
    uploaded_file = request.files['file']
    print("Init result file" , uploaded_file)
    uploaded_file.save(uploaded_file.filename)


    master_file_name = "Master.xlsx"
    start_time = dt.datetime.now()
    print("Excel Comparsion Start at: ",start_time)
    for master_file in pathlib.Path.cwd().glob(master_file_name):
        master_file_name = master_file_name

    master_wb = load_workbook(os.path.join(os.getcwd(),master_file))
    config_sheet = master_wb["config"]
    end_time = dt.datetime.now
    print("Excel Comparsion End at: ", end_time)
    print("Total Comparsion time: ", str(dt.datetime.now() - start_time))
    print("Done !")

    path = compare(config_sheet)
    return send_file(path, as_attachment=True)

if __name__ == "__main__":

    app.run(debug=True)