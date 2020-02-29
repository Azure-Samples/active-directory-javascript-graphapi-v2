  // Add here the endpoints for MS Graph API services you would like to use.
  const graphConfig = {
    graphMeEndpoint: "https://graph.microsoft.com/v1.0/me",
    graphMailEndpoint: "https://graph.microsoft.com/v1.0/me/messages"
  };

  // Add here scopes for access token to be used at MS Graph API endpoints.
  const tokenRequest = {
      scopes: ["Mail.Read"]
  };