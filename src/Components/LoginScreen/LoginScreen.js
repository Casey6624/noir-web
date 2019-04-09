import React, { useState } from "react"

import * as Msal from 'msal';

import "./LoginScreen.css"

import applicationConfig from "./appConfig"

export default function LoginScreen(props){

    let [userObj, setUserObj] = useState({})

    let [mailFolders, setMailFolders] = useState({})

    function getMessages(){
        let getMessageAPI = "https://graph.microsoft.com/v1.0/me/messages"

        fetch(getMessageAPI)
        .then(res => res.json())
        .then(data => {
            console.log(data)
        })
    }

    function getMailFolders(){
        console.log()
        let getMessagesGraphEndpoint = "https://graph.microsoft.com/v1.0/me/mailFolders"
        myMSALObj.acquireTokenSilent(applicationConfig.graphScopes).then(function (accessToken) {
            console.log(`Access Token is ${accessToken}`)
            callMSGraph(getMessagesGraphEndpoint, accessToken, getMailFoldersCallback);
        }, function (error) {
            console.log(error);
            // Call acquireTokenPopup (popup window) in case of acquireTokenSilent failure due to consent or interaction required ONLY
            if (error.indexOf("consent_required") !== -1 || error.indexOf("interaction_required") !== -1 || error.indexOf("login_required") !== -1) {
                myMSALObj.acquireTokenPopup(applicationConfig.graphScopes).then(function (accessToken) {
                    console.log(`Access Token is ${accessToken}`)
                    callMSGraph(getMessagesGraphEndpoint, accessToken, getMailFoldersCallback);
                }, function (error) {
                    console.log(error);
                });
            }
        });
    }

    function getMailFoldersCallback(data){
        console.log(data)
    }

    function getMostRecentMessage(){
        let API_URL = "https://graph.microsoft.com/v1.0/me/mailfolders/inbox/messages"
        myMSALObj.acquireTokenSilent(applicationConfig.graphScopes).then(function (accessToken) {
            fetch(API_URL, {
                headers: {
                    "Accept": "application/json",
                    "Authorization": `Bearer ${accessToken}`
                }
            })
            .then(res => res.json())
            .then(data => console.log(data))
            .catch(err => console.log(err))
        })
    }


    

    var myMSALObj = new Msal.UserAgentApplication(applicationConfig.clientID, applicationConfig.authority, acquireTokenRedirectCallBack, {storeAuthStateInCookie: true, cacheLocation: "localStorage"});

    function signIn() {
        myMSALObj.loginPopup(applicationConfig.graphScopes).then(function (idToken) {
            //Login Success
            acquireTokenPopupAndCallMSGraph();
        }, function (error) {
            console.log(`Couldn't contact API! ${error}`);
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
        console.table(data)

        setUserObj(data)
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
            <h1 id="LoginPageTitle">Noir Mail Login</h1>
            <p>{userObj.displayName}</p>
            <p>{userObj.mobilePhone}</p>
            <button id="LoginPageButton" onClick={signIn}>Sign In With A Microsoft Account</button>
            <br />
            <button id="LoginPageButton" onClick={signOut}>Sign Out</button>
            <br />
            <button id="LoginPageButton" onClick={getMailFolders}>Get Mail Folders</button>
            <br />
            <button id="LoginPageButton" onClick={getMostRecentMessage}>Get Most Recent Email</button>
        </div>
    )


}

