import React, { useState, useEffect } from "react"

import * as Msal from 'msal';

import "./LoginScreen.css"

import applicationConfig from "./appConfig"

export default function LoginScreen(props){

    let [welcome, setWelcome] = useState("")

    let [jsonData, setJsonData] = useState("")

    function getMessages(){
        let getMessageAPI = "https://outlook.office.com/api/v2.0/me/messages"

        fetch(getMessageAPI)
        .then(res => res.json())
        .then(data => {
            console.log(data)
        })
    }

    var myMSALObj = new Msal.UserAgentApplication(applicationConfig.clientID, applicationConfig.authority, acquireTokenRedirectCallBack, {storeAuthStateInCookie: true, cacheLocation: "localStorage"});


    useEffect(function(){
    
            console.log(myMSALObj)
    }, [])

    function signIn() {
        myMSALObj.loginPopup(applicationConfig.graphScopes).then(function (idToken) {
            //Login Success
            showWelcomeMessage();
            acquireTokenPopupAndCallMSGraph();
            getMessages()
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
        let welcomeName = myMSALObj.getUser().name
        setWelcome(`Welcome ${welcomeName}`)
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