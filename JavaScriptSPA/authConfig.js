 
  // Config object to be passed to Msal on creation
  const msalConfig = {
    auth: {
      clientId: "361c9ef8-eb0e-4c81-a8b8-bd4b6e2d5b15",
      authority: "https://login.microsoftonline.com/common/",
      redirectUri: "http://localhost:3000/",
    },
    cache: {
      cacheLocation: "sessionStorage", // This configures where your cache will be stored
      storeAuthStateInCookie: false, // Set this to "true" if you are having issues on IE11 or Edge
      forceRefresh: false // Set this to "true" to skip a cached token and go to the server to get a new
    }
  };  
  
  // Add here scopes for id token to be used at MS Identity Platform endpoints.
  const loginRequest = {
    scopes: ["openid", "profile", "User.Read"],
  };
