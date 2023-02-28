from flask import Blueprint, render_template, request, jsonify
#Activate setting of parameters and argruments, plus route
#Website Front end and HTML CSS Display
display = Blueprint(__name__, "DSVSpotQuotationMain")
@display.route("/")
def home():
    return render_template("Home.html", titlename="DSV Spot Quotation", clientname="kim", cityoforigin="manila")

@display.route("/")
def includesidebar():
    try:
        return render_template("sidebar.html")
    except Exception as e:
        return str(e)


