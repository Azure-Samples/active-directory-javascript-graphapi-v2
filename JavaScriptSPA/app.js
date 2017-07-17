
var graphAPIMeEndpoint = "https://graph.microsoft.com/v1.0/me";
var graphAPIScopes = ["https://graph.microsoft.com/user.read"];

// Initialize application
var userAgentApplication = new Msal.UserAgentApplication(msalconfig.clientID, null, displayUserInfo, {
    redirectUri: msalconfig.redirectUri
});

//Previous version of msal uses redirect url via a property
if (userAgentApplication.redirectUri) userAgentApplication.redirectUri = msalconfig.redirectUri;

var user = userAgentApplication.getUser();
window.onload = function () {
    //Add support to display user info in case of reload of the page
    if (!userAgentApplication.isCallback(window.location.hash) && window.parent === window && !window.opener) {
        if (user) {
            displayUserInfo();
        }
    }

}

/**
 * Display the results from Web API call in json format
 * 
 * @param {object} data - Results from API call
 * @param {object} token - The access token
 * @param {object} responseElement - HTML element to show the results
 * @param {object} showTokenElement - HTML element to show the RAW token
 */
function showAPIResponse(data, token, responseElement, showTokenElement) {
    console.log(data);
    responseElement.innerHTML = JSON.stringify(data, null, 4);
    if (showTokenElement) {
        showTokenElement.parentElement.classList.remove("hidden");
        showTokenElement.innerHTML = token;
    }
}

/**
 * Show an error message in the page
 * @param {any} endpoint - the endpoint used for the error message
 * @param {any} error - the error string
 * @param {any} errorElement - the HTML element in the page to display the error
 */
function showError(endpoint, error, errorElement) {
    console.error(error);
    var formattedError = JSON.stringify(error, null, 4);
    if (formattedError.length < 3) {
        formattedError = error;
    }
    errorElement.innerHTML = "Error calling " + endpoint + ": " + formattedError;
}

/**
 * Displays user information based on the information contained in the id token
 * And also calls the method to display the user profile via Microsoft Graph API
 */
function displayUserInfo() {
    var user = userAgentApplication.getUser();
    if (!user) {
        //If user is not signed in, then prompt user to sing-in via loginRedirect
        userAgentApplication.loginRedirect(graphAPIScopes);
    } else {
        // If user is already signed in, display the user info
        var userInfoElement = document.getElementById("userInfo");
        userInfoElement.parentElement.classList.remove("hidden");
        userInfoElement.innerHTML = JSON.stringify(user, null, 4);

        // Show Sign-Out button
        document.getElementById("signOutButton").classList.remove("hidden");

        //Now Call Graph API to show the user profile information
        callGraphAPI();
    }
}

/**
 * Call the Microsoft Graph API and display the results on the page
 */
function callGraphAPI() {
    var user = userAgentApplication.getUser();
    if (user) {
        var responseElement = document.getElementById("graphResponse");
        responseElement.parentElement.classList.remove("hidden");
        responseElement.innerText = "Calling Graph ...";
        callWebApiWithScope(graphAPIMeEndpoint,
            graphAPIScopes,
            responseElement,
            document.getElementById("errorMessage"),
            document.getElementById("accessToken"));
    } else {
        showError(graphAPIMeEndpoint, "User has not signed-in", document.getElementById("errorMessage"));
    }
}

/**
 * Call a Web API that requires scope, then display the response
 * 
 * @param {string} endpoint - The Web API endpoint
 * @param {object} scope - An array containing the API scopes
 * @param {object} responseElement - HTML element used to display the results
 * @param {object} errorElement = HTML element used to display an error message
 * @param {object} showTokenElement = HTML element used to display the RAW access token
 */
function callWebApiWithScope(endpoint, scope, responseElement, errorElement, showTokenElement) {
    //Try to acquire the token silently first
    userAgentApplication.acquireTokenSilent(scope)
        .then(function (token) {
            //After the access token is acquired, call the Web API, sending the acquired token
            callWebApiWithToken(endpoint, token, responseElement, errorElement, showTokenElement);
        }, function (error) {
            //If the acquireTokenSilent fails, then acquire the token interactively via acquireTokenPopup
            if (error) {
                userAgentApplication.acquireTokenPopup(scope).then(function (token) {
                        //After the access token is acquired, call the Web API, sending the acquired token
                        callWebApiWithToken(endpoint, token, responseElement, errorElement, showTokenElement);
                    },
                    function (error) {
                        showError(endpoint, error, errorElement);
                    });
            } else {
                showError(endpoint, error, errorElement);
            }
        });
}

/**
 * Call a Web API using an access token.
 * 
 * @param {any} endpoint - Web API endpoint
 * @param {any} token - Access token
 * @param {object} responseElement - HTML element used to display the results
 * @param {object} errorElement = HTML element used to display an error message
 * @param {object} showTokenElement = HTML element used to display the RAW access token
 */
function callWebApiWithToken(endpoint, token, responseElement, errorElement, showTokenElement) {
    var headers = new Headers();
    var bearer = "Bearer " + token;
    headers.append("Authorization", bearer);
    var options = {
        method: "GET",
        headers: headers
    };

    // Note that fetch API is not available in all browsers
    fetch(endpoint, options)
        .then(function (response) {
            var contentType = response.headers.get("content-type");
            if (response.status === 200 && contentType && contentType.indexOf("application/json") !== -1) {
                response.json()
                    .then(function (data) {
                        // Display response in the page
                        showAPIResponse(data, token, responseElement, showTokenElement);
                    })
                    .catch(function (error) {
                        showError(endpoint, error, errorElement);
                    });
            } else {
                response.json()
                    .then(function (data) {
                        // Display response in the page
                        showError(endpoint, data, errorElement);
                    })
                    .catch(function (error) {
                        showError(endpoint, error, errorElement);
                    });
            }
        })
        .catch(function (error) {
            showError(endpoint, error, errorElement);
        });
}

/**
 * Sign-out the user
 */
function signOut() {
    userAgentApplication.logout();
}