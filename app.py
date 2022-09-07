from flask_bootstrap import Bootstrap5
from flask import Flask, render_template, request
from FoxitPDFSDKPython3 import *

import pandas as pd

app = Flask(__name__)

sn = r"XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
key = r"XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
key =  key + r"XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" 
key =  key + r"XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" 
key =  key + r"XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" 
key =  key + r"XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" 
key =  key + r"XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" 
key =  key + r"XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" 
key =  key + r"XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX" 
key =  key + r"XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
key =  key + r"XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
key =  key + r"XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
key =  key + r"XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
key =  key + r"XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
key =  key + r"XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
key =  key + r"XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
key =  key + r"XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"

@app.route('/')
def initPdfSdk():
    sdkloaded = None
    code = Library.Initialize(sn, key)
    if code != e_ErrSuccess:
        sdkloaded = False
    else:
        sdkloaded = True
        
    return render_template("loadSDK.html", sdkloaded = sdkloaded)


@app.route('/loadData')
def selectFile():
    sdkloaded = None
    code = Library.Initialize(sn, key)
    if code != e_ErrSuccess:
        sdkloaded = False
    else:
        sdkloaded = True
    return render_template("selectFile.html", sdkloaded = sdkloaded)


@app.route('/loadData', methods = ['POST'])
def loadRowsToHtml():
    sdkloaded = None
    code = Library.Initialize(sn, key)
    if code != e_ErrSuccess:
        sdkloaded = False
    else:
        sdkloaded = True

    file_to_convert = request.files['fin_file']
    if file_to_convert.filename != '':
        file_to_convert.save(file_to_convert.filename)
        df = pd.read_csv(str(file_to_convert.filename), 
                        nrows=300,
                        usecols=["Invoice ID", "Product line", "Unit price","Quantity","Total","Date"],
                        parse_dates=["Date"])
        total_sales = df["Total"].sum()
        total_sales = "$ {:,.2f}".format(total_sales)
        with open('styles.txt', 'r') as myfile: styles = myfile.read()
        htmlfile = open("export.html","w")
        htmlfile.write("""<!DOCTYPE html>
                        <html>
                        <head>{1}</head>
                        <body>
                        <div class="resultTable">
                        <h1 style="color: #f36b16">SALES REPORT 2019</h1>
                        <h2 style="color: #f36b16">Total Sales = {3}</h1>
                        <h3 style="color: #f36b16">Extracted from: {2}</h3>
                        {0}
                        </div>
                        </body>
                        </html>""".format(df.to_html(classes="resultTable"), 
                                            styles, 
                                            str(file_to_convert.filename), 
                                            total_sales))
        htmlfile.close() 

    return render_template("loadToHtml.html", 
                            filename = str(file_to_convert.filename), 
                            preview_rows = df, 
                            sdkloaded = sdkloaded, 
                            total_sales = total_sales)



@app.route('/generatePDF')
def htmlToPdf():
    sdkloaded = None
    code = Library.Initialize(sn, key)
    if code != e_ErrSuccess:
        sdkloaded = False
    else:
        sdkloaded = True

    html = "X:/path/to/export.html"
    output_path =  "X:/path/to/Report2019.pdf"
    engine_path = "X:/path/to/fxhtml2pdf.exe"
    cookies_path = ""
    time_out = 50

    pdf_setting_data = HTML2PDFSettingData()
    pdf_setting_data.page_height = 640
    pdf_setting_data.page_width = 900
    pdf_setting_data.page_mode = 1
    pdf_setting_data.scaling_mode = 2

    Convert.FromHTML(html, engine_path, cookies_path, pdf_setting_data, output_path, time_out)

    doc = PDFDoc("Report2019.pdf")
    error_code = doc.Load("")
    if error_code!= e_ErrSuccess:
        return 0
    
    settings = WatermarkSettings()
    settings.flags = WatermarkSettings.e_FlagASPageContents | WatermarkSettings.e_FlagOnTop
    settings.offset_x = 0
    settings.offset_y = 0
    settings.opacity = 50
    settings.position = 1
    settings.rotation = -45.0
    settings.scale_x = 8.0
    settings.scale_y = 8.0

    text_properties = WatermarkTextProperties()
    text_properties.alignment = e_AlignmentCenter
    text_properties.color = 0xF68C21
    text_properties.font_style = WatermarkTextProperties.e_FontStyleNormal
    text_properties.line_space = 2
    text_properties.font_size = 14.0
    text_properties.font = Font(Font.e_StdIDTimesB)
    watermark = Watermark(doc, "CONFIDENTIAL", text_properties, settings)

    nPageCount = doc.GetPageCount()
    for i in range(0, nPageCount):
        page = doc.GetPage(i)
        page.StartParse(PDFPage.e_ParsePageNormal, None, False)
        watermark.InsertToPage(page)
    doc.SaveAs("Report2019.pdf", PDFDoc.e_SaveFlagNoOriginal)

    print("Converted HTML to PDF successfully.")

    success = True

    return render_template("generatePDF.html", sdkloaded = sdkloaded, success = success )








