import * as Bluebird from "bluebird";
import * as Crypto from "crypto";
import * as Knex from "knex";
import connections from "../db/connections";
import Token from "../db/token";
import * as express from "express";
import * as https from "https";
import * as querystring from "querystring";
import * as jwt from "jsonwebtoken";

const DevMode: boolean = process.env.NODE_ENV === "development";

let REFRESH_MINIMUM = 60;

let router = express.Router({ mergeParams: true });
export default router;

const connection: Knex = Knex(connections[process.env.NODE_ENV]);
connection.migrate.latest(connections);

let salt = null;
function getSalt() {
  if (salt === null) {
    if (DevMode) {
      salt = [];
    } else {
      salt = JSON.parse(process.env.salt);
    }
  }
  return salt;
}

let clientInfo = null;
let getClientInfo = function () {
  if (clientInfo == null) {
    if (DevMode) {
      clientInfo = require("../../secrets/clientSecret");
    } else {
      clientInfo = JSON.parse(process.env.ClientSecretJson);
    }
  }
  return clientInfo;
};

let oauth = null;
let getOAuth = function () {
  if (oauth === null) {
    let info = getClientInfo();
    oauth = {
      authEndpoint: "/oauth2/authorize",
      baseUrl: "app.vssps.visualstudio.com",
      clientId: info.client_id.toString(),
      clientSecret: info.client_secret.toString(),
      redirectUri: info.redirect_uris[0],
      scopes: info.scopes,
      tokenEndpoint: "/oauth2/token",
    };
  }
  return oauth;
};

function beautify(body) {
  let begin = "-----BEGIN CERTIFICATE-----";
  let end = "-----END CERTIFICATE-----";

  body = body.replace(/-/g, "+");
  body = body.replace(/_/g, "/");

  let arr = [];
  arr.push(begin);
  while (body.length > 0) {
    let line = body.slice(0, 64);
    arr.push(line);
    body = body.slice(64);
  }
  arr.push(end);
  return arr.join("\n");
}

function bytesToString(bytes) {
  let str = "";
  for (let i = 0; i < bytes.length; i++) {
    str += String.fromCharCode(bytes[i]);
  }
  return str;
}

export function getUID(token: string, callback: (token: string) => any): void {
  let decoded = jwt.decode(token, { complete: true });
  let appctx = JSON.parse(decoded.payload.appctx);
  https.get(appctx.amurl, (response) => {
    let output = "";
    response.on("data", (d) => {
      output += d;
    });

    response.on("end", () => {
      let responseBlob = JSON.parse(output);
      responseBlob.keys.forEach((key) => {
        if (key.keyinfo.x5t === decoded.header.x5t) {
          let publicKey = beautify(key.keyvalue.value);
          jwt.verify(token, publicKey, { algorithms: ["RS256"] }, (err, verified) => {
            if (!verified) {
              callback("");
            }
            let id = appctx.msexchuid;
            let url = appctx.amurl;
            let input = bytesToString(getSalt()) + id + url;
            let hash = Crypto.createHash("sha256");
            hash.update(input);
            let body = hash.digest("base64");
            body = body.replace(/\+/g, "-");
            body = body.replace(/\//g, "_");
            callback(body);
          });
        }
      });
    });
  });
}

export function getToken(uid: string, callback: (token: Token) => any): void {
  connection.select(Token.TokenKey, Token.ExpiryKey, Token.RefreshKey).from(Token.TableName).where(Token.IdKey, uid).then((output: Token[]) => {
    if (output == null || output.length === 0) {
      callback(null);
    } else {
      let token: Token = output[0]; // There should only be 1 row

      callback(token);
    }

  });
};

function newToken(assertion, refresh, callback) {
  let auth = getOAuth();
  let data = {
    assertion: assertion,
    client_assertion: auth.clientSecret,
    client_assertion_type: "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
    grant_type: refresh ? "refresh_token" : "urn:ietf:params:oauth:grant-type:jwt-bearer",
    redirect_uri: auth.redirectUri,
  };
  let options = {
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    host: oauth.baseUrl,
    method: "POST",
    path: oauth.tokenEndpoint,
  };
  let request = https.request(options, function (response) {
    let str = "";
    let errored = false;

    response.on("data", function (chunk) {
      str += chunk;
    });

    response.on("end", function () {
      if (!errored) {
        let result = JSON.parse(str);
        callback(null, result.access_token, result.refresh_token, result);
      }
    });

    response.on("error", function (err) {
      errored = true;
      callback(err);
    });
  });
  request.write(querystring.stringify(data));
  request.end();
}

function db(req, res) {
  getUID(req.query.user, (uid) => {
    getToken(uid, (token: Token) => {
      if (token != null) { // recieved row
        let expiryLimit = new Date();
        expiryLimit.setMinutes(expiryLimit.getMinutes() + REFRESH_MINIMUM);
        if (token.expiry > expiryLimit) { // if the token doesn't expire before our limit
          res.send("success");
        } else {
          refreshToken(uid, token.refresh, res);
        }
      } else {
        res.send("failure");
      }
    });
  });
};
router.use("/db", db);

function callback(req, res) {
  let user = req.query.state;
  newToken(req.query.code, false, (err, accessToken, refreshToken, results) => {
    if (err) {
      console.log(err);
    } else {
      res.redirect("../done");
      saveToken(Token.getInstance(user, accessToken, results["expires_in"], refreshToken));
    }
  });

};
router.use("/callback", callback);

function authorize(req, res) {

  getUID(req.query.user, (uid) => {
    let auth = getOAuth();
    let authParams = {
      client_id: auth.clientId,
      redirect_uri: auth.redirectUri,
      response_type: "Assertion",
      scope: auth.scopes,
      state: uid,
    };
    res.redirect("https://" + oauth.baseUrl + oauth.authEndpoint + "?" + querystring.stringify(authParams));
  });
};
router.use("/", authorize);

export function disconnect(user, callback) {
  deleteToken(user).then(() => callback()).catch((error) => callback(error));
}

function refreshToken(user, refresh, res) {
  newToken(refresh, true, (err, accessToken, refreshToken, results) => {
    if (err) {
      console.log(err);
    } else {
      deleteToken(user)
        .then(() => saveToken(Token.getInstance(user, accessToken, results["expires_in"], refreshToken)))
        .then(() => res.send("success"));
    }
  });
};

async function deleteToken(id: string): Bluebird<void> {
  return connection.delete().from(Token.TableName).where(Token.IdKey, id).then(() => {
    console.log(`Removed: ${id}`);
  });
}

async function saveToken(token: Token): Bluebird<void> {
  return connection.insert(token).into(Token.TableName).then(() => {
    console.log(`Added: ${JSON.stringify(token)}`);
  });
}
