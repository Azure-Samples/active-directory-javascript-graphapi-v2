// Create the main myMSALObj instance
// configuration parameters are located at authConfig.js
const myMSALObj = new Msal.UserAgentApplication(msalConfig); 

let accessToken;

// Register Callbacks for Redirect flow
myMSALObj.handleRedirectCallback(authRedirectCallBack);

function authRedirectCallBack(error, response) {
  if (error) {
      console.log(error);
  } else {
      if (response.tokenType === "id_token") {
          console.log("id_token acquired at: " + new Date().toString()); 
          
          if (myMSALObj.getAccount()) {
            showWelcomeMessage(myMSALObj.getAccount());
          }

      } else if (response.tokenType === "access_token") {
        console.log("access_token acquired at: " + new Date().toString());
        accessToken = response.accessToken;

        try {
          callMSGraph(graphConfig.graphMailEndpoint, accessToken, updateUI);
        } catch(err) {
          console.log(err)
        } finally {
          profileButton.classList.add('d-none');
          mailButton.classList.remove('d-none');
        }
      } else {
          console.log("token type is:" + response.tokenType);
      }
  }
}

if (myMSALObj.getAccount()) {
  showWelcomeMessage(myMSALObj.getAccount());
}

function signIn() {
  myMSALObj.loginRedirect(loginRequest);
}

function signOut() {
  myMSALObj.logout();
}

// This function can be removed if you do not need to support IE
function getTokenRedirect(request, endpoint) {
  return myMSALObj.acquireTokenSilent(request)
      .then((response) => {
        console.log(response);
        if (response.accessToken) {
            console.log("access_token acquired at: " + new Date().toString());
            accessToken = response.accessToken;

            if (accessToken) {
              try {
                callMSGraph(endpoint, accessToken, updateUI);
              } catch(err) {
                console.log(err)
              } finally {
                profileButton.classList.add('d-none');
                mailButton.classList.remove('d-none');
              }
            }
        }
      })
      .catch(error => {
          console.log("silent token acquisition fails. acquiring token using redirect");
          // fallback to interaction when silent call fails
          return myMSALObj.acquireTokenRedirect(request);
      });
}

function seeProfile() {
  getTokenRedirect(loginRequest, graphConfig.graphMeEndpoint);
}
  
function readMail() {
  if (accessToken) {
    callMSGraph(graphConfig.graphMailEndpoint, accessToken, updateUI);
  } else {
    getTokenRedirect(tokenRequest, graphConfig.graphMailEndpoint);
  }
}