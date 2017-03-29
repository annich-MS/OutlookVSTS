import * as Bluebird from "bluebird";
import * as Crypto from "crypto";
import * as Express from "express";
import * as https from "https";
import * as jwt from "jsonwebtoken";
import * as Knex from "knex";
import * as querystring from "querystring";

import connections from "../auth/connections";
import Token, { ServerTokenData } from "../auth/token";
import AuthInfo from "../auth/authInfo";

const Salt: number[] = JSON.parse(process.env.SALT);
const Auth: AuthInfo = AuthInfo.getInstance();

const RefreshMinimumInMinutes: number = 15;
const RefreshMinimum: number = RefreshMinimumInMinutes * 60 * 1000; // sec/min * ms/sec => ms/min

const connection: Knex = Knex(connections[process.env.NODE_ENV]);
connection.migrate.latest();

const router = Express.Router({ mergeParams: true });
export default router;

/**
 * Converts an Auth Certificate recieved by the AppContext to the form required by jwt
 * @param body the Auth Certificate to be converted
 */
function beautify(body: string): string {

  body = body.replace(/-/g, "+");
  body = body.replace(/_/g, "/");

  let arr: string[] = [];
  arr.push("-----BEGIN CERTIFICATE-----");
  while (body.length > 0) {
    let line = body.slice(0, 64);
    arr.push(line);
    body = body.slice(64);
  }
  arr.push("-----END CERTIFICATE-----");
  return arr.join("\n");
}

/**
 * helper function to turn a byteArray into a string
 * @param bytes 
 */
function bytesToString(bytes: number[]): string {
  let str = "";
  for (let i = 0; i < bytes.length; i++) {
    str += String.fromCharCode(bytes[i]);
  }
  return str;
}

/**
 * processes and validates an office.js UserIdentityToken and converts to a internal UID for authentication
 * @param token a UserIdentityToken to be processed
 */
export async function getUID(token: string): Promise<string> {
  return new Promise<string>((resolve, reject) => {

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
                reject();
              }
              let id = appctx.msexchuid;
              let url = appctx.amurl;
              let input = bytesToString(Salt) + id + url;
              let hash = Crypto.createHash("sha256");
              hash.update(input);
              let body = hash.digest("base64");
              body = body.replace(/\+/g, "-");
              body = body.replace(/\//g, "_");
              resolve(body);
            });
          }
        });
      });
    });
  });
}

/**
 * querys the database for a valid token given a UID
 * @param uid the UID to recieve a token for
 */
export async function getToken(uid: string): Promise<Token> {
    let output: Token[] = await connection.select(Token.TokenKey, Token.ExpiryKey, Token.RefreshKey).from(Token.TableName).where(Token.IdKey, uid);
    if (output == null || output.length === 0) {
      return null;
    } else {
      let token: Token = output[0]; // There should only be 1 row
      return Token.getSanitized(token);
    }
};

/**
 * Querys VSTS for a new token
 * @param assertion the assertion token
 * @param refresh a flag to determine if the assertion is a refresh or new token request
 * @param callback 
 */
function newToken(uid: string, assertion: string, refresh: boolean, callback: (error: any, token?: Token) => void) {
  let data = {
    assertion: assertion,
    client_assertion: Auth.secret,
    client_assertion_type: "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
    grant_type: refresh ? "refresh_token" : "urn:ietf:params:oauth:grant-type:jwt-bearer",
    redirect_uri: Auth.redirect,
  };
  let options = {
    headers: {
      "Content-Type": "application/x-www-form-urlencoded",
    },
    host: Auth.baseUrl,
    method: "POST",
    path: `/${Auth.tokenEndpoint}`,
  };
  let request = https.request(options, function (response) {
    let str = "";
    let errored = false;

    response.on("data", function (chunk) {
      str += chunk;
    });

    response.on("end", function () {
      if (!errored) {
        let result: ServerTokenData = JSON.parse(str);
        callback(null, Token.getInstance(uid, result));
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

/**
 * Entry point to check if user exists
 * @param req the request recieved
 * @param res the response to be sent
 */
async function db(req: Express.Request, res: Express.Response) {
  let uid: string = await getUID(req.query.user);
  let token: Token = await getToken(uid);
  if (token != null) {
    let expiryLimit: number = Date.now() + RefreshMinimum;
    if (token.expiry > expiryLimit) { // if the token doesn't expire before our limit
      res.send("success");
    } else {
      refreshToken(uid, token.refresh, res);
    }
  } else {
    res.send("failure");
  }
};
router.use("/db", db);

/**
 * Entry point for returning authentication
 * @param req the request recieved
 * @param res the response to be sent
 */
function callback(req, res) {
  newToken(req.query.state, req.query.code, false, (err: any, token: Token) => {
    if (err) {
      console.log(err);
    } else {
      res.redirect("../done");
      saveToken(token);
    }
  });

};
router.use("/callback", callback);

/**
 * Entry point for requesting authentication
 * @param req the request recieved
 * @param res the response to be sent
 */
function authorize(req, res) {

  getUID(req.query.user).then((uid) => {
    let authParams = {
      client_id: Auth.id,
      redirect_uri: Auth.redirect,
      response_type: "Assertion",
      scope: Auth.scopes,
      state: uid,
    };
    res.redirect(`https://${Auth.baseUrl}/${Auth.authEndpoint}?${querystring.stringify(authParams)}`);
  });
};
router.use("/", authorize);

/**
 * Entry point function for removing a user from the database
 * @param user 
 * @param callback 
 */
export function disconnect(user, callback) {
  deleteToken(user).then(() => callback()).catch((error) => callback(error));
}

/**
 * Gets a new token, then replaces the expired token with the new one
 * @param user the user to replace a token for
 * @param refresh the refresh token
 * @param res the response to be sent
 */
function refreshToken(user: string, refresh: string, res: Express.Response) {
  newToken(user, refresh, true, (err: any, token: Token) => {
    if (err) {
      console.log(err);
    } else {
      deleteToken(user)
        .then(() => saveToken(token))
        .then(() => res.send("success"));
    }
  });
};

/**
 * Removes the token for a given user from the db
 * @param id the UID to remove from the db
 */
async function deleteToken(id: string): Bluebird<void> {
  return connection.delete().from(Token.TableName).where(Token.IdKey, id);
}

/**
 * Insert a token into the database
 * @param token the token to be added to the db
 */
async function saveToken(token: Token): Bluebird<void> {
  return connection.insert(token).into(Token.TableName);
}
