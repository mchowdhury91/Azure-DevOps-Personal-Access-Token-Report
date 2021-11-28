import uuid
import requests
import json
import threading
import random
import xlsxwriter
import time
from flask import Flask, render_template, session, request, redirect, url_for, Response
from flask_session import Session  # https://pythonhosted.org/Flask-Session
from datetime import datetime
import msal
import app_config
import jinja2

app = Flask(__name__)
app.config.from_object(app_config)
Session(app)

# This section is needed for url_for("foo", _external=True) to automatically
# generate http scheme when this sample is running on localhost,
# and to generate https scheme when it is deployed behind reversed proxy.
# See also https://flask.palletsprojects.com/en/1.0.x/deploying/wsgi-standalone/#proxy-setups
from werkzeug.middleware.proxy_fix import ProxyFix
app.wsgi_app = ProxyFix(app.wsgi_app, x_proto=1, x_host=1)

class PersonalAccessToken:
    def __init__(self, displayName, owner, isValid, validFrom, validTo):
            self.displayName = displayName
            self.owner = owner
            self.isValid = isValid
            self.validFrom = validFrom
            self.validTo = validTo

@app.route("/")
def index():
    if not session.get("user"):
        return redirect(url_for("login"))
    return render_template('index.html', user=session["user"], version=msal.__version__)

@app.route("/display_progress", methods=['GET','POST'])
def progress():
    def generate():
        x = 0
        while x <= 100:
            print(x)
            x = x + 10
            time.sleep(0.2)
            yield "data:" + str(x) +"\n\n"
    return Response(generate(), mimetype='text/event-stream')

@app.route("/login")
def login():
    # Technically we could use empty list [] as scopes to do just sign in,
    # here we choose to also collect end user consent upfront
    session["flow"] = _build_auth_code_flow(scopes=app_config.SCOPE)
    return render_template("login.html", auth_url=session["flow"]["auth_uri"], version=msal.__version__)

@app.route(app_config.REDIRECT_PATH)  # Its absolute URL must match your app's redirect_uri set in AAD
def authorized():
    try:
        cache = _load_cache()
        result = _build_msal_app(cache=cache).acquire_token_by_auth_code_flow(
            session.get("flow", {}), request.args)
        if "error" in result:
            return render_template("auth_error.html", result=result)
        session["user"] = result.get("id_token_claims")
        _save_cache(cache)
    except ValueError:  # Usually caused by CSRF
        pass  # Simply ignore them
    return redirect(url_for("index"))

@app.route("/logout")
def logout():
    session.clear()  # Wipe out user and its token cache from session
    return redirect(  # Also logout from your tenant's web session
        app_config.AUTHORITY + "/oauth2/v2.0/logout" +
        "?post_logout_redirect_uri=" + url_for("index", _external=True))

@app.route("/graphcall", methods=['GET','POST'])
def graphcall():
    token = _get_token_from_cache(app_config.SCOPE)
    if not token:
        return redirect(url_for("login"))
    response = requests.get(  # Use token to call downstream service
        app_config.USERLISTENDPOINT,
        headers={'Authorization': 'Bearer ' + token['access_token']},
        ).json()
    users = response['value']
    userTokenArr = []
    today = datetime.now()
    todayString = today.strftime("%b-%d-%Y-%#H-%M")
    workbookName = "PersonalAccessTokens_" + todayString + ".xlsx"
    def generate(users,userTokenArr,workbookName):
        #setup Excel spreadsheet to export PAT data to
        workbook = xlsxwriter.Workbook(workbookName)
        bold = workbook.add_format({'bold': True})

        worksheet = workbook.add_worksheet()
        
        worksheet.write("A1","DisplayName", bold)
        worksheet.write("B1","Owner", bold)
        worksheet.write("C1","isValid", bold)
        worksheet.write("D1","validFrom", bold)
        worksheet.write("E1","validTo", bold)
        colA_width = 15
        colB_width = 15
        colC_width = 10
        colD_width = 15
        colE_width = 15


        userTokenArr = []
        userCount = response['count']
        row = 2
        x = 0
        for user in users:
            thisUserPATEndPoint = app_config.USERPATENDPOINT.replace("{subjectDescriptor}",user['descriptor'])
            userTokens = requests.get(  # Use token to call downstream service
                thisUserPATEndPoint,
                headers={'Authorization': 'Bearer ' + token['access_token']},
                ).json()
                
            if(len(userTokens["value"]) > 0):
                for userToken in userTokens["value"]:
                    newUserToken = PersonalAccessToken(userToken["displayName"], user["displayName"], userToken["isValid"], userToken["validFrom"], userToken["validTo"])
                    userTokenArr.append(newUserToken)
                    print("Token DisplayName: " + newUserToken.displayName)
                    print("Current userTokenArr length: " + str(len(userTokenArr)))
                    cellA = "A" + str(row)
                    cellB = "B" + str(row)
                    cellC = "C" + str(row)
                    cellD = "D" + str(row)
                    cellE = "E" + str(row)

                    worksheet.write(cellA, userToken["displayName"])
                    worksheet.write(cellB, user["displayName"])
                    worksheet.write(cellC, userToken["isValid"])
                    worksheet.write(cellD, userToken["validFrom"])
                    worksheet.write(cellE, userToken["validTo"])
                    
                    colA_width = max(colA_width, len(userToken["displayName"])+1)
                    colB_width = max(colB_width, len(user["displayName"])+1)
                    #colC_width = max(colC_width, len(userToken["isValid"])+1)
                    colD_width = max(colD_width, len(userToken["validFrom"])+1)
                    colE_width = max(colE_width, len(userToken["validTo"])+1)

                    row += 1
            
            x += 1
            progress = str(x) + "/" + str(userCount) + " users processed...processed user = " + user['displayName']
            
            if(x == userCount):
                message = progress + "...Complete"
            else:
                message = progress
            message += ";"
            for userToken in userTokens["value"]:
                message += "<tr>"
                message += "<td style='width: 20%'>" + userToken["displayName"] + "</td>"
                message += "<td style='width: 20%'>" + user["displayName"] + "</td>"
                message += "<td style='width: 20%'>" + str(userToken["isValid"]) + "</td>"
                message += "<td style='width: 20%'>" + userToken["validFrom"] + "</td>"
                message += "<td style='width: 20%'>" + userToken["validTo"] + "</td>"
                message += "</tr>"
                print(message)
            yield 'data: %s\n\n' % message
            

        worksheet.set_column(0,0,colA_width)
        worksheet.set_column(1,1,colB_width)
        worksheet.set_column(2,2,colC_width)
        worksheet.set_column(3,3,colD_width)
        worksheet.set_column(4,4,colE_width)
        workbook.close()

    return Response(generate(users,userTokenArr,workbookName), mimetype="text/event-stream")

def _load_cache():
    cache = msal.SerializableTokenCache()
    if session.get("token_cache"):
        cache.deserialize(session["token_cache"])
    return cache

def _save_cache(cache):
    if cache.has_state_changed:
        session["token_cache"] = cache.serialize()

def _build_msal_app(cache=None, authority=None):
    return msal.ConfidentialClientApplication(
        app_config.CLIENT_ID, authority=authority or app_config.AUTHORITY,
        client_credential=app_config.CLIENT_SECRET, token_cache=cache)

def _build_auth_code_flow(authority=None, scopes=None):
    return _build_msal_app(authority=authority).initiate_auth_code_flow(
        scopes or [],
        redirect_uri=url_for("authorized", _external=True))

def _get_token_from_cache(scope=None):
    cache = _load_cache()  # This web app maintains one cache per session
    cca = _build_msal_app(cache=cache)
    accounts = cca.get_accounts()
    if accounts:  # So all account(s) belong to the current signed-in user
        result = cca.acquire_token_silent(scope, account=accounts[0])
        _save_cache(cache)
        return result

app.jinja_env.globals.update(_build_auth_code_flow=_build_auth_code_flow)  # Used in template

if __name__ == "__main__":
    app.run(threaded=True)

