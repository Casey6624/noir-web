import React, { useState } from "react"

import * as Msal from 'msal';

import "./LoginScreen.css"

export default function LoginScreen(props){

    let [welcome, setWelcome] = useState("")

    let [jsonData, setJsonData] = useState("")

    var applicationConfig = {
        clientID: "edaa9aec-1db8-4991-b21d-524c1e5cffba",
        authority: "https://login.microsoftonline.com/common",
        graphScopes: ["user.read"],
        graphEndpoint: "https://graph.microsoft.com/v1.0/me"
    };

    var myMSALObj = new Msal.UserAgentApplication(applicationConfig.clientID, applicationConfig.authority, acquireTokenRedirectCallBack,
        {storeAuthStateInCookie: true, cacheLocation: "localStorage"});

        console.log(myMSALObj)

    function signIn() {
        myMSALObj.loginPopup(applicationConfig.graphScopes).then(function (idToken) {
            //Login Success
            showWelcomeMessage();
            acquireTokenPopupAndCallMSGraph();
        }, function (error) {
            console.log(error);
        });
    }

    function signOut() {
        myMSALObj.logout();
    }

   function acquireTokenPopupAndCallMSGraph() {
        //Call acquireTokenSilent (iframe) to obtain a token for Microsoft Graph
        myMSALObj.acquireTokenSilent(applicationConfig.graphScopes).then(function (accessToken) {
            callMSGraph(applicationConfig.graphEndpoint, accessToken, graphAPICallback);
        }, function (error) {
            console.log(error);
            // Call acquireTokenPopup (popup window) in case of acquireTokenSilent failure due to consent or interaction required ONLY
            if (error.indexOf("consent_required") !== -1 || error.indexOf("interaction_required") !== -1 || error.indexOf("login_required") !== -1) {
                myMSALObj.acquireTokenPopup(applicationConfig.graphScopes).then(function (accessToken) {
                    callMSGraph(applicationConfig.graphEndpoint, accessToken, graphAPICallback);
                }, function (error) {
                    console.log(error);
                });
            }
        });
    }

    function callMSGraph(theUrl, accessToken, callback) {
        var xmlHttp = new XMLHttpRequest();
        xmlHttp.onreadystatechange = function () {
            if (this.readyState == 4 && this.status == 200)
                callback(JSON.parse(this.responseText));
        }
        xmlHttp.open("GET", theUrl, true); // true for asynchronous
        xmlHttp.setRequestHeader('Authorization', 'Bearer ' + accessToken);
        xmlHttp.send();
    }

    function graphAPICallback(data) {
        //Display user data on DOM
        setJsonData = JSON.stringify(data, null, 2); 
    }

    function showWelcomeMessage() {
        setWelcome = `Welcome   ${myMSALObj.getUser().name}`
        var loginbutton = document.getElementById('SignIn');
    }

    // This function can be removed if you do not need to support IE
    function acquireTokenRedirectAndCallMSGraph() {
        //Call acquireTokenSilent (iframe) to obtain a token for Microsoft Graph
        myMSALObj.acquireTokenSilent(applicationConfig.graphScopes).then(function (accessToken) {
          callMSGraph(applicationConfig.graphEndpoint, accessToken, graphAPICallback);
        }, function (error) {
            console.log(error);
            //Call acquireTokenRedirect in case of acquireToken Failure
            if (error.indexOf("consent_required") !== -1 || error.indexOf("interaction_required") !== -1 || error.indexOf("login_required") !== -1) {
                myMSALObj.acquireTokenRedirect(applicationConfig.graphScopes);
            }
        }); 
    }

    function acquireTokenRedirectCallBack(errorDesc, token, error, tokenType)
    {
     if(tokenType === "access_token")
     {
         callMSGraph(applicationConfig.graphEndpoint, token, graphAPICallback);
     } else {
            console.log("token type is:"+tokenType);
     } 

    }

    // Browser check variables
    var ua = window.navigator.userAgent;
    var msie = ua.indexOf('MSIE ');
    var msie11 = ua.indexOf('Trident/');
    var msedge = ua.indexOf('Edge/');
    var isIE = msie > 0 || msie11 > 0;
    var isEdge = msedge > 0;

    //If you support IE, our recommendation is that you sign-in using Redirect APIs
    //If you as a developer are testing using Edge InPrivate mode, please add "isEdge" to the if check 
    if (!isIE) {
        if (myMSALObj.getUser()) {// avoid duplicate code execution on page load in case of iframe and popup window.
            showWelcomeMessage();
            acquireTokenPopupAndCallMSGraph();
        }
    }
    else {
        document.getElementById("SignIn").onclick = function () {
            myMSALObj.loginRedirect(applicationConfig.graphScopes);
        };

        if (myMSALObj.getUser() && !myMSALObj.isCallback(window.location.hash)) {// avoid duplicate code execution on page load in case of iframe and popup window.
            showWelcomeMessage();
            acquireTokenRedirectAndCallMSGraph();
        }
    }

    return(
        <div id="LoginPageContainer">
        <p>{welcome}</p>
            <h1 id="LoginPageTitle">Noir Mail Login</h1>
            <p>{jsonData}</p>
            <button id="LoginPageButton" onClick={signIn}>Sign In With A Microsoft Account</button>
            <br />
            <button id="LoginPageButton" onClick={signOut}>Sign Out</button>
        </div>
    )


}