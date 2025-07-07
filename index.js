require("dotenv").config();
const express = require("express");
const axios = require("axios");
const { ConfidentialClientApplication } = require("@azure/msal-node");
const sqlite3 = require("sqlite3").verbose();

const app = express();
const port = 3001;
const cors = require("cors");
const db = new sqlite3.Database(":memory:");

app.use(
  cors({
    origin: "http://localhost:3000", // Adjust this to your frontend URL
    credentials: true, // Allow credentials if needed
  })
);

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

db.serialize(() => {
  // Create a table to store user data
  db.run(`CREATE TABLE IF NOT EXISTS user_email_integration (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    userId TEXT NOT NULL,
    displayName TEXT NOT NULL,
    email TEXT NOT NULL,
    idToken TEXT NOT NULL,
    accessToken TEXT NOT NULL,
    refreshToken TEXT NOT NULL
  )`);

  // Redirect to Microsoft login
  app.get("/login", (req, res) => {
    console.log("Login request received:");
    const getAuthCodeUrlParams = {
      scopes: SCOPES,
      redirectUri: process.env.REDIRECT_URI,
    };

    msalClient
      .getAuthCodeUrl(getAuthCodeUrlParams)
      .then((response) => {
        console.log("response", response);
        res.redirect(response);
      })
      .catch((error) => {
        console.log("Error getting auth code URL:", error);
        res.status(500).send(error.message);
      });
  });

  function getUserData(email) {
    return new Promise((resolve, reject) => {
      db.all(
        "SELECT * FROM user_email_integration WHERE email = ?",
        [email],
        (err, rows) => {
          console.log("Fetching user data rows:", rows);
          if (err) {
            console.error("Error fetching user data:", err);
            reject(err);
          }
          if (!rows?.length) {
            console.log("No user found with email:", email);
            reject(new Error("User not found"));
          }
          const user = rows[0];
          console.log("User found:", user);
          resolve(user);
        }
      );
    });
  }

  // function getAllUsers() {
  //   return new Promise((resolve, reject) => {
  //     db.all("SELECT * FROM user_email_integration", (err, rows) => {
  //       console.log("Fetching all users rows:", rows);
  //       if (err) reject(err);
  //       else resolve(rows);
  //     });
  //   });
  // }

  function addUser(userData) {
    return new Promise((resolve, reject) => {
      db.run(
        "INSERT INTO user_email_integration (userId, displayName, email, idToken, accessToken, refreshToken) VALUES (?, ?, ?, ?, ?, ?)",
        [
          userData.id,
          userData.displayName,
          userData.mail,
          userData.idToken,
          userData.accessToken,
          userData.refreshToken,
        ],
        function (err) {
          if (err) {
            console.error("Error adding user:", err);
            reject(err);
          } else {
            console.log("User added with ID:", this.lastID);
            resolve(this.lastID);
          }
        }
      );
    });
  }

  app.get("/calendar/source/:email", async (req, res) => {
    console.log("Calendar source request received", req.params.email);
    try {
      const user = await getUserData(req.params.email);
      if (!user) {
        console.log("No user found with email:", req.params.email);
        return res.status(404).send("User not found");
      }
      res.json({
        userId: user.userId,
        displayName: user.displayName,
        email: user.email,
        idToken: user.idToken,
        accessToken: user.accessToken,
        refreshToken: user.refreshToken || "",
      });
    } catch (error) {
      console.error("Error fetching user data:", error);
      return res.status(404).send("User not found");
    }
  });

  // Handle the redirect with auth code
  app.get("/redirect", async (req, res) => {
    console.log("Redirect received with request:", req);
    const tokenRequest = {
      code: req.query.code,
      scopes: SCOPES,
      redirectUri: process.env.REDIRECT_URI,
    };

    try {
      const tokenResponse = await msalClient.acquireTokenByCode(tokenRequest);

      // console.log("Response:", tokenResponse);

      const accessToken = tokenResponse.accessToken;
      const refreshToken = tokenResponse.refreshToken;
      const idToken = tokenResponse.idToken;

      // Example call to Microsoft Graph (get profile)
      const graphRes = await axios.get("https://graph.microsoft.com/v1.0/me", {
        headers: { Authorization: `Bearer ${accessToken}` },
      });

      // console.log({
      //   accessTokenExpiresIn: tokenResponse.expiresOn,
      //   refreshToken: refreshToken ? "[Hidden for security]" : "N/A",
      //   profile: graphRes.data,
      // });
      await addUser({
        id: graphRes.data.id,
        displayName: graphRes.data.displayName,
        mail: graphRes.data.mail || graphRes.data.userPrincipalName,
        idToken: idToken,
        accessToken: accessToken,
        refreshToken: refreshToken || "",
      });
      // await getAllUsers();
      // res.send(`Hello, ${graphRes.data.displayName}!`);
      res.redirect(`http://localhost:3000/calendar`);
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
});
