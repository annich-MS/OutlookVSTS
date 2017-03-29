import * as Authenticate from "./authenticate";
import Token from "../auth/token";
import * as Express from "express";
import * as querystring from "querystring";
import * as Buffer from "buffer";
import * as request from "request-promise";
import * as stream from "string-to-stream";
import * as flow from "xml-flow";

let router = Express.Router({ mergeParams: true });
export default router;
let API1_0 = "1.0";
let API2_0_PREVIEW = "2.0-preview.1";
let API2_0 = "2.0";
let API3_0_PREVIEW = "3.0-preview";

let BATCH_SIZE: number = 100;

let FIELDS = {
  AREA_PATH: "/fields/System.AreaPath",
  DESCRIPTION: "/fields/System.Description",
  ITERATION_PATH: "/fields/System.IterationPath",
  RELATIONS: "/relations/-",
  REPRO_STEPS: "/fields/Microsoft.VSTS.TCM.ReproSteps",
  TITLE: "/fields/System.Title",
};

function createError(type, more) {
  return JSON.stringify({ error: { more: more, type: type } });
}

/**
 * Makes an authenticated https request
 * 
 * @param {string} user - the user to authenticate as
 * @param {Object} options - options to use in the https request
 * @param {requestCallback} callback - the callback to make upon completion
 */
async function makeAuthenticatedRequest(user, options): Promise<string> {
  let uid: string = await Authenticate.getUID(user);
  let token: Token = await Authenticate.getToken(uid);
  return await handleAuthentication(token, options);
}

/**
 * Parses the authentication response and passes the https request forward if successful
 * 
 * @param {Object} output
 * @param {Object} context
 * @param {Object} context.callback
 */
async function handleAuthentication(token: Token, options): Promise<string> {
  if (token != null) {
    options.headers.Authorization = `Bearer ${token.token}`;
    return await makeHttpsRequest(options);
  } else {
    console.log(`could not find token for user ${token.id}`);
    return null;
  }
}

function toUTF8Array(str) {
  let utf8 = [];
  for (let i = 0; i < str.length; i++) {
    let charcode = str.charCodeAt(i);
    if (charcode < 0x80) {
      utf8.push(charcode);
    } else if (charcode < 0x800) {
      utf8.push(0xc0 | (charcode >> 6),
        0x80 | (charcode & 0x3f));
    } else if (charcode < 0xd800 || charcode >= 0xe000) {
      utf8.push(0xe0 | (charcode >> 12),
        0x80 | ((charcode >> 6) & 0x3f),
        0x80 | (charcode & 0x3f));
    } else { // surrogate pair
      i++;
      // UTF-16 encodes 0x10000-0x10FFFF by
      // subtracting 0x10000 and splitting the
      // 20 bits of 0x0-0xFFFFF into two halves
      charcode = 0x10000 + (((charcode & 0x3ff) << 10)
        | (str.charCodeAt(i) & 0x3ff));
      utf8.push(0xf0 | (charcode >> 18),
        0x80 | ((charcode >> 12) & 0x3f),
        0x80 | ((charcode >> 6) & 0x3f),
        0x80 | (charcode & 0x3f));
    }
  }
  return utf8;
}

/**
 * Makes an https request based on the contents of the options letiable.
 * 
 * @param {any} options
 * @param {any} callback
 */
async function makeHttpsRequest(options): Promise<string> {
  // add derived headers
  if (options.body) {
    options.headers["Content-Length"] = toUTF8Array(options.body).length;
  }
  console.log(options.method + ": " + options.uri);
  try {
    let output: string = await request(options);
    if (!options.isXML) {
      try {
        JSON.parse(output);
      } catch (e) {
        output = createError("Unparseable output", output);
      }
    }
    return output;
  } catch (error) {
    console.log(JSON.stringify(error));
    return createError("Request Error", error);
  }

}

/**
 * Helper function for formatting the options object for https requests
 * 
 * @param {any} query - query arguments as json object
 * @param {string} path - path part of url to call
 * @param {any} headers - json object representing a list of headers
 * @param {string} host -  hostname part of url to call
 * @param {any} method - expected method for the https request
 * @returns options object for use with https request
 */
function createOptions(input, method) {
  return {
    body: input.body,
    headers: input.headers || {},
    method: method,
    uri: "https://" + input.host + encodeURI(input.path) + "?" + querystring.stringify(input.query),
  };
}

/**
 * Helper function to produce formatted PATCH request items
 * 
 * @param {any} path - the letiable to be set
 * @param {any} value - the value to set
 * @returns a vaild item for PATCH requests
 */
function jsonPatchItem(path, value) {
  return { "op": "add", "path": path, "value": value };
}

/**
 * Retrieves an item from visual studio
 * 
 * @param {any} req
 * @param {any} res
 */
router.getItem = function (req, res) {
  let input = req.query;
  input.query["api-version"] = API1_0;
  input.host = input.host + ".visualstudio.com";
  input.path = "/DefaultCollection/_apis/wit/workitems";
  makeAuthenticatedRequest(input.user, createOptions(input, "GET")).then((output) => { res.send(output); });
};
router.use("/getItem", router.getItem);

/**
 * 
 * 
 * @param {any} req
 * @param {any} res
 */
router.me = function (req, res) {
  let input = req.query;
  if (!input.query) { input.query = {}; }
  input.query["api-version"] = API1_0;
  input.host = "app.vssps.visualstudio.com";
  input.path = "/_apis/profile/profiles/me";
  makeAuthenticatedRequest(input.user, createOptions(input, "GET")).then((output) => { res.send(output); });
};
router.use("/me", router.me);

/**
 * 
 * 
 * @param {any} req
 * @param {any} res
 */
router.accounts = function (req, res) {
  let input = req.query;
  if (!input.query) { input.query = {}; }
  input.query.memberId = input.memberId;
  input.query["api-version"] = API1_0;
  input.host = "app.vssps.visualstudio.com";
  input.path = "/_apis/Accounts";
  makeAuthenticatedRequest(input.user, createOptions(input, "GET")).then((output) => { res.send(output); });
};
router.use("/accounts", router.accounts);

/**
 * 
 * 
 * @param {any} req
 * @param {any} res
 */
router.projects = function (req, res) {
  let input = req.query;
  if (!input.query) { input.query = {}; }

  input.query["api-version"] = API1_0;
  input.host = input.account + ".visualstudio.com";
  input.path = "/DefaultCollection/_apis/projects";
  makeAuthenticatedRequest(input.user, createOptions(input, "GET")).then((output) => { res.send(output); });

};
router.use("/projects", router.projects);

/**
 * 
 * 
 * @param {any} req
 * @param {any} res
 */
router.getTeams = async function (req, res) {
  let input = req.query;
  if (!input.query) { input.query = {}; }

  input.query["api-version"] = API1_0;
  input.host = input.account + ".visualstudio.com";
  input.path = "/DefaultCollection/_apis/projects/" + input.project + "/teams";

  let uid: string = await Authenticate.getUID(input.user);
  let token: Token = await Authenticate.getToken(uid);
  let expected: number = 0;

  let teams = [];
  try {

    while (teams.length === expected) {
      input.query.$top = BATCH_SIZE;
      input.query.$skip = expected;
      let options = createOptions(input, "GET");
      options.headers.Authorization = `Bearer ${token.token}`;
      let output = JSON.parse(await makeHttpsRequest(options));
      teams = teams.concat(output.value);
      expected += BATCH_SIZE;
    }
    res.send({ value: teams });
  } catch (error) {
    res.send(error);
  }
};
router.use("/getTeams", router.getTeams);

/**
 * 
 *
 * @param {any} req
 * @param {any} res
 */
router.getTeamField = function (req, res) {
  let input = req.query;
  if (!input.query) { input.query = {}; }

  input.query["api-version"] = API2_0_PREVIEW;
  input.host = input.account + ".visualstudio.com";
  input.path = "/DefaultCollection/" + input.project + "/" + input.team + "/_apis/Work/TeamSettings/TeamFieldValues";
  makeAuthenticatedRequest(input.user, createOptions(input, "GET")).then((output) => { res.send(output); });
};
router.use("/getTeamField", router.getTeamField);

/**
 * 
 * 
 * @param {any} req
 * @param {any} res
 */
router.getCurrentIteration = function (req, res) {
  let input = req.query;
  if (!input.query) { input.query = {}; }
  input.query["$timeframe"] = "current";
  input.query["api-version"] = API2_0_PREVIEW;
  input.host = input.account + ".visualstudio.com";
  input.path = "/defaultcollection/" + input.project + "/" + input.team + "/_apis/work/teamsettings/iterations";
  makeAuthenticatedRequest(input.user, createOptions(input, "GET")).then((output) => { res.send(output); });
};
router.use("/getCurrentIteration", router.getCurrentIteration);

router.getMessage = async function (req, res) {
  let input = req.query;
  let message: string = await downloadMessageFromEWS(input.ewsId, input.url, input.token);
  res.send(message);
};
router.use("/getMessage", router.getMessage);

async function downloadMessageFromEWS(messageId, ewsUrl, token): Promise<string> {
  let body = `<?xml version="1.0" encoding="utf-8"?>` +
    `<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"` +
    `               xmlns:xsd="http://www.w3.org/2001/XMLSchema"` +
    `               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/" ` +
    `               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">` +
    `    <soap:Header>` +
    `        <RequestServerVersion Version="Exchange2013" ` +
    `                              xmlns="http://schemas.microsoft.com/exchange/services/2006/types"` +
    `                              soap:mustUnderstand="0" />` +
    `    </soap:Header>` +
    `    <soap:Body>` +
    `        <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages"` +
    `                 xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">` +
    `            <ItemShape>` +
    `                <t:BaseShape>IdOnly</t:BaseShape>` +
    `                <t:IncludeMimeContent>true</t:IncludeMimeContent>` +
    `            </ItemShape>` +
    `            <ItemIds>` +
    `                <t:ItemId Id="` + messageId + `"/>` +
    `            </ItemIds>` +
    `       </GetItem>` +
    `   </soap:Body>` +
    `</soap:Envelope>`;
  let options = {
    body: body,
    headers: {
      Authorization: "Bearer " + token,
      "Content-Type": "text/xml; charset=utf-8",
      "Content-Length": body.length,
    },
    isXML: true,
    method: "POST",
    uri: ewsUrl,
  };
  let output: string = await makeHttpsRequest(options);
  let id: string = await extractMessageId(output);
  return id;
}

async function extractMessageId(response): Promise<string> {
  let parser = new flow(stream(response));
  let done = false;
  let output: string = "";
  parser.on("tag:t:mimecontent", (element) => {
    done = true;
    output = element["$text"];
  });
  parser.on("end", () => {
    if (!done) {
      output = createError("Invalid EWS response", response);
    }
  });
  return output;
}

router.uploadAttachment = function (req, res) {
  let input = req.query;
  input.body = decodeBase64Data(req.body);
  input.host = input.account + ".visualstudio.com";
  input.path = "/DefaultCollection/_apis/wit/attachments";
  input.query = {
    "api-version": API1_0,
    "filename": input.filename,
  };
  makeAuthenticatedRequest(input.user, createOptions(input, "POST")).then((output) => { res.send(output); });
};
router.use("/uploadAttachment", router.uploadAttachment);

function decodeBase64Data(data) {
  return (new Buffer.Buffer(data, "base64")).toString("utf8");
}

router.createTask = function (req, res) {
  let input = req.query;
  input.host = input.account + ".visualstudio.com";
  input.path = "/DefaultCollection/" + input.project + "/_apis/wit/workitems/$" + input.type;
  input.query = {
    "api-version": API1_0,
  };
  input.headers = {
    "Content-Type": "application/json-patch+json",
  };
  input.body = [
    jsonPatchItem(FIELDS.TITLE, input.title),
    jsonPatchItem(FIELDS.AREA_PATH, input.areapath),
    jsonPatchItem(FIELDS.ITERATION_PATH, input.project + input.iteration),
    jsonPatchItem(FIELDS.DESCRIPTION, req.body),
  ];
  if (input.attachment !== "") {
    input.body.push(jsonPatchItem(FIELDS.RELATIONS, { "rel": "AttachedFile", "url": input.attachment }));
  }

  input.body = JSON.stringify(input.body);
  makeAuthenticatedRequest(input.user, createOptions(input, "PATCH")).then((output) => { res.send(output); });
};
router.use("/createTask", router.createTask);

router.reply = async function (req, res) {
  let input = req.query;
  input.host = "outlook.office365.com";
  input.path = "/api/v2.0/me/messages/" + input.item + "/replyAll";
  input.headers = {
    "Content-Type": "application/json",
    "Authorization": "Bearer " + input.token,
  };
  input.body = req.body;
  input.isXML = true;
  let output: string = await makeHttpsRequest(createOptions(input, "POST"));
  res.send(output);
};
router.use("/reply", router.reply);

router.backlog = function (req, res) {
  let input = req.query;
  input.host = input.account + ".visualstudio.com";
  input.path = "/defaultcollection/" + input.project + "/" + input.team + "/_apis/work/teamsettings";
  input.query = {
    "api-version": API3_0_PREVIEW,
  };
  makeAuthenticatedRequest(input.user, createOptions(input, "GET")).then((output) => {
    res.send(output);
  });
};
router.use("/backlog", router.backlog);

router.disconnect = function (req, res) {
  Authenticate.getUID(req.query.user).then((uid) => {
    Authenticate.disconnect(uid, (err) => {
      let output = "{}";
      if (err) {
        output = createError("Database Error", err);
      }
      res.send(output);
    });
  });
};
router.use("/disconnect", router.disconnect);
