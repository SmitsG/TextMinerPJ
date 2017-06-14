from flask import Flask, send_file, request, render_template, redirect, send_from_directory
import os
from datetime import date

import TextMiner

app = Flask(__name__)

# This function will return the main page of the website.
@app.route('/')
def helloWorld():
    return render_template("Home.html")

# This function will return the webpage for textmining.
@app.route('/textminen')
def startTextmining():
    return render_template("Textminen.html")

'''
This function is called by Textmining.html.
When both textfields are filled by the user and the user clicks on the submit button, textmining will start.
'''
@app.route("/get_search_param",methods=['POST'])
def getSearchParam():
    searchWords = str(request.form["searchWords"])
    maxNumberAbstracts = int(request.form["maxNumberAbstracts"])
    TextMiner.main(searchWords, maxNumberAbstracts)
    return redirect("/visualisatie")

#This function returns the visualisation website, for the visualisation of the wordcloud.
@app.route("/visualisatie")
def createImage():
    return render_template("Visualisatie.html", title="WordCloud")

#This funciton will send the WordCloud image to the visualisation html.
@app.route('/fig')
def fig():
    return send_file("WordCloud.png", mimetype='image/png')

#This function will return the download page, so the user can download the excel file.
@app.route('/download')
def download():
    return render_template("Download.html")

#This function will download the Excelfile to the user's computer
@app.route('/excel2', methods=['POST'])
def excel():
    uploads = app.config["UPLOAD_FOLDER"] = app.root_path
    fileName = "Excel.xls"
    return send_from_directory(directory=uploads, filename = fileName)

if __name__ == '__main__':
    app.run(debug=True)
