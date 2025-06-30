require("dotenv").config();
const express = require("express");
const axios = require("axios");
const { ConfidentialClientApplication } = require("@azure/msal-node");

const app = express();
const port = 3000;

// MSAL config
const msalConfig = {
  auth: {
    clientId: process.env.CLIENT_ID,
    authority: `https://login.microsoftonline.com/${process.env.TENANT_ID}`,
    clientSecret: process.env.CLIENT_SECRET,
  },
};

const msalClient = new ConfidentialClientApplication(msalConfig);

// Scopes for Microsoft Graph
const SCOPES = ["user.read", "openid", "profile", "offline_access"];

// Redirect to Microsoft login
app.get("/login", (req, res) => {
  const authUrlParams = {
    scopes: SCOPES,
    redirectUri: process.env.REDIRECT_URI,
  };

  msalClient
    .getAuthCodeUrl(authUrlParams)
    .then((response) => res.redirect(response))
    .catch((error) => res.status(500).send(error.message));
});

// Handle the redirect with auth code
app.get("/redirect", async (req, res) => {
  const tokenRequest = {
    code: req.query.code,
    scopes: SCOPES,
    redirectUri: process.env.REDIRECT_URI,
  };

  try {
    const tokenResponse = await msalClient.acquireTokenByCode(tokenRequest);

    const accessToken = tokenResponse.accessToken;
    const refreshToken = tokenResponse.refreshToken;

    // Example call to Microsoft Graph (get profile)
    const graphRes = await axios.get("https://graph.microsoft.com/v1.0/me", {
      headers: { Authorization: `Bearer ${accessToken}` },
    });

    res.json({
      accessTokenExpiresIn: tokenResponse.expiresOn,
      refreshToken: refreshToken ? "[Hidden for security]" : "N/A",
      profile: graphRes.data,
    });
  } catch (err) {
    console.error(err);
    res.status(500).send("Auth failed");
  }
});

// Simple homepage
app.get("/", (req, res) => {
  res.send(`<a href="/login">Login with Microsoft</a>`);
});

app.listen(port, () => {
  console.log(`Server listening at http://localhost:${port}`);
});
