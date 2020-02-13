// Browser check variables
// If you support IE, our recommendation is that you sign-in using Redirect APIs
// If you as a developer are testing using Edge InPrivate mode, please add "isEdge" to the if check
const ua = window.navigator.userAgent;
const msie = ua.indexOf("MSIE ");
const msie11 = ua.indexOf("Trident/");
const msedge = ua.indexOf("Edge/");
const isIE = msie > 0 || msie11 > 0;
const isEdge = msedge > 0;

// Some globals 
let signInType;
let access_token;

// Select DOM elements to work with
const welcomeDiv = document.getElementById("WelcomeMessage");
const signInButton = document.getElementById("SignIn");
const cardDiv = document.getElementById("card-div");
const mailButton = document.getElementById("readMail");
const profileDiv = document.getElementById("profile-div");

// Create the main myMSALObj instance
// configuration parameters are located at config.js
const myMSALObj = new Msal.UserAgentApplication(msalConfig); 

// Register Callbacks for Redirect flow
myMSALObj.handleRedirectCallback(authRedirectCallBack);

function authRedirectCallBack(error, response) {
  if (error) {
    console.log(error);
  } else {
    if (response.tokenType === "id_token" && myMSALObj.getAccount() && !myMSALObj.isCallback(window.location.hash)) {

      console.log('id_token acquired at: ' + new Date().toString());

      showWelcomeMessage();
      acquireTokenRedirectAndCallMSGraph(graphConfig.graphMeEndpoint, loginRequest, graphAPICallback);
    } else if (response.tokenType === "access_token") {
      access_token = response.accessToken
      console.log('access_token acquired at: ' + new Date().toString());

      callMSGraph(graphConfig.graphEndpoint, response.accessToken, graphAPICallback);
    } else {
      console.log("token type is:" + response.tokenType);
    }
  }
}

// Redirect: once login is successful and redirects with tokens, call Graph API
if (myMSALObj.getAccount() && !myMSALObj.isCallback(window.location.hash)) {
  // avoid duplicate code execution on page load in case of iframe and Popup window.
  showWelcomeMessage();
  acquireTokenRedirectAndCallMSGraph(graphConfig.graphMeEndpoint, loginRequest, graphAPICallback);
}

function signIn(method) {
  
  signInType = isIE ? "Redirect" : method;

  if (signInType === "Popup") {
    myMSALObj.loginPopup(loginRequest)
      .then(loginResponse => {  

        console.log('id_token acquired at: ' + new Date().toString());

        if (myMSALObj.getAccount()) {
          // avoid duplicate code execution on page load in case of iframe and Popup window.

          showWelcomeMessage();
          acquireTokenPopupAndCallMSGraph(graphConfig.graphMeEndpoint, loginRequest, graphAPICallback);
        }
    }).catch(function (error) {
      console.log(error);
    });

  } else if (signInType === "Redirect") {
    myMSALObj.loginRedirect(loginRequest)
  }
}

function readMail() {
  if (myMSALObj.getAccount()) {
    if(signInType === "Popup") {
      acquireTokenPopupAndCallMSGraph(graphConfig.graphMailEndpoint, tokenRequest, graphAPICallback);
      mailButton.style.display = 'none';
    } else if (signInType === "Redirect") {
      acquireTokenRedirectAndCallMSGraph(graphConfig.graphMeEndpoint, tokenRequest, graphAPICallback);
      mailButton.style.display = 'none';
    }
  }
}

function signOut() {
  myMSALObj.logout();
}

// Call to the resource acquiring a token to a specific scope set
function acquireTokenPopupAndCallMSGraph(endpoint, request) {
  //Call acquireTokenSilent (iframe) to obtain a token for Microsoft Graph
  myMSALObj.acquireTokenSilent(request)
    .then(tokenResponse => {

        access_token = tokenResponse.accessToken
        console.log('access_token acquired at: ' + new Date().toString());

        callMSGraph(endpoint, tokenResponse.accessToken, graphAPICallback);
    }).catch(error => {
        console.log(error);
        // Call acquireTokenPopup (Popup window) in case of acquireTokenSilent failure 
        // due to consent or interaction required ONLY
        if (requiresInteraction(error.errorCode)) {
          myMSALObj.acquireTokenPopup(request)
            .then(tokenResponse => {

              access_token = tokenResponse.accessToken
              console.log('access_token acquired at: ' + new Date().toString());

              callMSGraph(endpoint, tokenResponse.accessToken, graphAPICallback);
            }).catch(error => {
              console.log(error);
            });
        }
    });
}

// This function can be removed if you do not need to support IE
function acquireTokenRedirectAndCallMSGraph(endpoint, request) {
    //Call acquireTokenSilent (iframe) to obtain a token for Microsoft Graph
    myMSALObj.acquireTokenSilent(request)
      .then(tokenResponse => {

        access_token = tokenResponse.accessToken
        console.log('access_token acquired at: ' + new Date().toString());

        callMSGraph(endpoint, tokenResponse.accessToken, graphAPICallback);
      }).catch(error => {
          console.log("error is: " + error);
          console.log("stack:" + error.stack);

          //Call acquireTokenRedirect in case of acquireToken Failure
          if (requiresInteraction(error.errorCode)) {
            myMSALObj.acquireTokenRedirect(request);
          }
      });
}

function showWelcomeMessage() {

  // Reconfiguring DOM elements
  cardDiv.style.display = 'initial';
  welcomeDiv.innerHTML = `Welcome ${myMSALObj.getAccount().name}`;
  signInButton.nextElementSibling.style.display = 'none';
  signInButton.setAttribute("onclick", "signOut();");
  signInButton.setAttribute('class', "btn btn-success dropdown-toggle")
  signInButton.innerHTML = "Sign Out";
}

function graphAPICallback(data, endpoint) {
  console.log('Graph API responded at: ' + new Date().toString());

  if (endpoint === graphConfig.graphMeEndpoint) {
    const title = document.createElement('p');
    title.innerHTML = "<strong>Title: </strong>" + data.jobTitle;
    const email = document.createElement('p');
    email.innerHTML = "<strong>Mail: </strong>" + data.mail;
    const phone = document.createElement('p');
    phone.innerHTML = "<strong>Phone: </strong>" + data.businessPhones[0];
    const address = document.createElement('p');
    address.innerHTML = "<strong>Location: </strong>" + data.officeLocation;
    profileDiv.appendChild(title);
    profileDiv.appendChild(email);
    profileDiv.appendChild(phone);
    profileDiv.appendChild(address);
  
    
  } else if (endpoint === graphConfig.graphMailEndpoint) {

    if (data.value.length < 1) {
      alert("Your mailbox is empty!")
    } else {
      const tabList = document.getElementById("list-tab");
      const tabContent = document.getElementById("nav-tabContent");
      data.value.map((d, i) => {
        // Keeping it simple
        if (i < 10) {
          const listItem = document.createElement("a");
          listItem.setAttribute("class", "list-group-item list-group-item-action")
          listItem.setAttribute("id", "list" + i + "list")
          listItem.setAttribute("data-toggle", "list")
          listItem.setAttribute("href", "#list" + i)
          listItem.setAttribute("role", "tab")
          listItem.setAttribute("aria-controls", i)
          listItem.innerHTML = d.subject;
          tabList.appendChild(listItem)
  
          const contentItem = document.createElement("div");
          contentItem.setAttribute("class", "tab-pane fade")
          contentItem.setAttribute("id", "list" + i)
          contentItem.setAttribute("role", "tabpanel")
          contentItem.setAttribute("aria-labelledby", "list" + i + "list")
          contentItem.innerHTML = "<strong> from: " + d.from.emailAddress.address + "</strong><br><br>" + d.bodyPreview + "...";
          tabContent.appendChild(contentItem);
        }
      });
    }
  }
}

function requiresInteraction(errorCode) {
  if (!errorCode || !errorCode.length) {
    return false;
  }
  
  return (
    errorCode === "consent_required" ||
    errorCode === "interaction_required" ||
    errorCode === "login_required"
  );
}