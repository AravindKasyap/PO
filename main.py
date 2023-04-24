import os
import time
# import magic
import urllib.request
from app import app
from pywinauto import Desktop, Application, timings
import unittest
import os
from appium import webdriver
from flask import Flask, flash, request, redirect, render_template, send_file
from werkzeug.utils import secure_filename
from pywinauto_recorder.player import *
from pywinauto import application, timings, Desktop, findwindows, findbestmatch, controls
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
from appium.webdriver.common.mobileby import MobileBy
# from selenium.webdriver.common.by import By
# from selenium.webdriver.support.ui import WebDriverWait
# from selenium.webdriver.support import expected_conditions as EC


ALLOWED_EXTENSIONS = set(
    ['txt', 'pdf', 'png', 'jpg', 'jpeg', 'gif', 'csv', 'xls'])


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def upload_form():
    # return render_template('upload.html')
    return render_template('index.html')


@app.route('/', methods=['POST'])
def upload_file():
    if request.method == 'POST':
        # check if the post request has the files part
        if 'files[]' not in request.files:
            flash('No file part')
            return redirect(request.url)
        files = request.files.getlist('files[]')
        for file in files:
            if file and allowed_file(file.filename):
                filename = secure_filename(file.filename)
                file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
        flash('File(s) successfully uploaded')
        # open_pbix_file_new()
        #automate_powerbi_desktop()
        new_automation()
        # open_pbix()

        return redirect('/')


@app.route('/download')
def download_file():
    # path = "html2pdf.pdf"
    # path = "info.xlsx"
    path = "input_template.csv"
    # path = "sample.txt"
    return send_file(path, as_attachment=True)


def automate_powerbi_desktop():
    app = Application().start(
        r"C:\Program Files\Microsoft Power BI Desktop\bin\PBIDesktop.exe", timeout=30)
    # app.window(title_re=".*Power BI Desktop")
    main_dlg = app.window(title='Untitled - Power BI Desktop')
    main_dlg.wait('visible')
    print(main_dlg.print_control_identifiers())

#         'C:\\Program Files\\Microsoft Power BI Desktop\\bin\\PBIDesktop.exe')
#     desktop = Desktop(backend='uia')
#     window = desktop.window(
#         title_re='Untitled - Power BI Desktop', visible_only=True)
#     # Wait for Power BI Desktop to fully launch
#     # window.wait('ready', timeout=60)
#     time.sleep(15)

#     with UIPath(u"Untitled - Power BI Desktop||Window"):
#         with UIPath(u"||Pane->||Pane->https://ms-pbi.pbi.microsoft.com/pbi/Web/Views/ReportView.htm||Pane->https://ms-pbi.pbi.microsoft.com/pbi/Web/Views/ReportView.htm - Web content||Pane->||Pane->||Pane->||Pane->||Pane->||Pane->||Pane->||Pane"):
#             click(u"||Document")
#             click(u"||Document")
#             click(u"||Document")
#             click(u"||Document")
#             click(u"||Document")
#         with UIPath(u"||Pane->||Pane->https://ms-pbi.pbi.microsoft.com/pbi/Web/Views/ReportView.htm||Pane->https://ms-pbi.pbi.microsoft.com/pbi/Web/Views/ReportView.htm - Web content||Pane->||Pane->||Pane->||Pane->||Pane->||Pane->||Pane->||Pane->||Document->||Pane"):
#             click(u"||Group")
#    # window.File.click_input()
#     # items=window.print_control_identifiers()
#     # print(items)
#     # window.GroupBox18.click_input()

#     # Interact with the Power BI Desktop window

#     # window.File.click_input()
#     # time.sleep(3)
#     # window.GroupBox18.click_input()
#     # time.sleep(20)
#     window.minimize()
#     window.restore()
#     window.maximize()


def new_automation():
    app = application.Application()
# Replace with the path to Power BI Desktop on your machine
    app.start("C:\\Program Files\\Microsoft Power BI Desktop\\bin\\PBIDesktop.exe")
    app.wait_cpu_usage_lower(threshold=5)
# Wait for the application to start up
    time.sleep(13)


    timings.wait_until_passes(20, 0.5, lambda: app.window(
    title="Untitled - Power BI Desktop"))

# Open the .pbix file
    app.window(title="Untitled - Power BI Desktop").type_keys("^o")
    app.window(title="Open").type_keys(
        "C:\\Users\\bigtapp\\Downloads\\Project_Backup.pbix")
    app.window(title="Open").type_keys("{ENTER}")
    app.window()

    # Wait for the file to open
    timings.wait_until_passes(20, 0.5, lambda: app.window(
        title="Project_Backup - Power BI Desktop"))
    desktop = Desktop(backend='uia')
    time.sleep(10)
    window = desktop.window(
        title_re='Project_Backup - Power BI Desktop', visible_only=True)

    home_tab = window.child_window(
        title="Home", auto_id="home", control_type="TabItem", found_index=1)
    home_tab.click_input()
    time.sleep(10)

    publish = window.child_window(
        title="Publish", auto_id="publish", control_type="Button", found_index=0)
    publish.click_input()
    time.sleep(3)


    child_window = window.window(title="Publish to Power BI", found_index=0)
    select_button_matches = findbestmatch.find_best_control_matches(
        "Select", child_window.descendants())
    select_button_matches_1 = select_button_matches[0]
    print(select_button_matches_1)

    if isinstance(select_button_matches_1, controls.uia_controls.ButtonWrapper):
        select_button_matches_1.click()
    time.sleep(10)
    app.kill()

    # chrome
    # driver = webdriver.Chrome()
    # url = "http://127.0.0.1:5000/"
    # driver.get(url)
    # # Navigate to the first tab and refresh the page
    # driver.refresh()
    # time.sleep(5)
    # driver.maximize_window()
    # #/html/body/header/div/nav/ul/li[4]/a
    # driver.find_element(By.XPATH,'/html/body/header/div/nav/ul/li[4]/a').click()
    # #driver.find_element_by_link_text("Dashboard").click()
    # #driver.find_element_by_name("Dashboard").click()

if __name__ == "__main__":
    app.run()
