#pip install pytailwindcss
#pip install -U selenium
from flask import Flask
from waitress import serve
from views import display
from selenium import webdriver
#reference:https://www.youtube.com/watch?v=kng-mJJby8g
#Activate website or localhost
app = Flask(__name__)
app.register_blueprint(display, url_prefix="/")

if __name__ == '__main__':
    app.run(debug=True)
    #this will be changed due to server change
    serve(app, host="0.0.0.0", port=8080)
    #default host is 8080
    #localhost name is>> http://localhost:8080/DSVSpotQuotationMain/


