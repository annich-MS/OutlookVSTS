var express = require('express');
var fs = require('fs');
var tedious = require('tedious');
var https = require('https');
var querystring = require('querystring');
var PROD = process.env.prod;
var TYPES = tedious.TYPES;

var REFRESH_MINIMUM = 60;

var router = express.Router({ mergeParams: true });
module.exports = router;


var _clientInfo = null;
getClientInfo = function () {
  if (_clientInfo == null) {
    if (PROD) {
      _clientInfo = JSON.parse(process.env.ClientSecretJson);
    }
    else {
      _clientInfo = require('../secrets/clientSecret');
    }
  }
  return _clientInfo;
}

var _oauth = null;
getOAuth = function () {
  if (_oauth === null) {
    var clientInfo = getClientInfo();
    _oauth = {
      clientId: clientInfo.client_id.toString(),
      clientSecret: clientInfo.client_secret.toString(),
      baseUrl: 'app.vssps.visualstudio.com',
      authEndpoint: '/oauth2/authorize',
      tokenEndpoint: '/oauth2/token',
      redirectUri: clientInfo.redirect_uris[0]
    };
  }
  return _oauth;
};

var dbConfig = "";
getDbConfig = function () {
  if (dbConfig === "") {
   if (PROD) {
      dbConfig = process.env.dbConfigJson;
    }
    else {
      var dbFile = require('../secrets/dbConfig.js')
      dbConfig = JSON.stringify(dbFile);
    }
  }
  return dbConfig;
}

var table = PROD ? 'dbo.Users' : 'dbo.TestUsers'

var GET_TOKEN_QUERY = "SELECT TOP 1 x.Token, x.Expiry, x.Refresh FROM " + table + " AS x WHERE Id=@Id"
var SAVE_TOKEN_QUERY = 
`IF EXISTS(` + GET_TOKEN_QUERY + `)
  UPDATE ` + table + ` SET Token=@Token, Expiry=DATEADD(ss, @Expiry, GETDATE()), Refresh=@Refresh WHERE Id=@Id;
ELSE
  INSERT INTO ` + table + `(Id, Token, Expiry, Refresh) VALUES (@Id, @Token, DATEADD(ss, @Expiry, GETDATE()), @Refresh);`;

createConnection = function (reason, callback) {
  var config = JSON.parse(getDbConfig());
  var connection = new tedious.Connection(config);
  connection.on('end', () => {
    console.log("Close Connection for " + reason);
  });
  connection.on('error', (err) => {
    console.log(err);
  });
  connection.on('connect', function (err) {
    console.log("Open Connection for " + reason);
    if (err) {
      console.log(err);
    }
    callback(connection);
  });
}


router.getToken = function (user, callback) {
  createConnection("getToken", (connection) => {
    var request = new tedious.Request(GET_TOKEN_QUERY, function (err, rowcount, rows) {
      if (err) {
        callback({ success: false, error: err });
      }
      if (rows == null || rowcount == 0) {
        callback({ success: false });
      }
      else {
        var row = rows[0]; //There should only be 1 row

        var data = {
          token: row[0].value,
          expiry: Date.parse(row[1].value),
          refresh: row[2].value
        };
        callback({ success: true, data: data })
      }
      connection.close();
    });
    request.addParameter('Id', TYPES.VarChar, user);
    connection.execSql(request);
  });
}

router.newToken = function (user, assertion, refresh, callback) {

  var oauth = getOAuth();
  var data = {
    assertion: assertion, 
    client_assertion_type: "urn:ietf:params:oauth:client-assertion-type:jwt-bearer",
    grant_type: refresh ? "refresh_token" : "urn:ietf:params:oauth:grant-type:jwt-bearer",
    client_assertion: oauth.clientSecret,
    redirect_uri: oauth.redirectUri
  };
  var options = {
    host: oauth.baseUrl,
    path: oauth.tokenEndpoint,
    method: 'POST',
    headers: {
      'Content-Type': 'application/x-www-form-urlencoded',
    }
  };
  var request = https.request(options, function(response) {
    var str = '';
    var errored = false;
    
    response.on('data', function(chunk) {
      str += chunk
    });
    
    response.on('end', function() {
      if(!errored) {
        var result = JSON.parse(str);
        callback(null, result.access_token, result.refresh_token, result);
      }
    });
    
    response.on('error', function(err) {
      errored = true;
      callback(err);
    });
  });
  request.write(querystring.stringify(data));
  request.end();
}

router.db = function (req, res) {
  router.getToken(req.query.user, (response) => {
    if (response.success) { // recieved row
      var data = response.data;
      var expiryLimit = new Date();
      expiryLimit.setMinutes(expiryLimit.getMinutes() + REFRESH_MINIMUM);
      if (data.expiry > expiryLimit) { // if the token doesn't expire before our limit
        res.send("success");
      }
      else {
        console.log("middle state")
        router.refreshToken(req.query.user, data.refresh, res);
      }
    }
    else {
      res.send("failure");
    }
  });
};
router.use('/db', router.db);

router.callback = function (req, res) {
  user = req.query.state;
  router.newToken(user, req.query.code, false, (err, access_token, refresh_token, results) => {
      if(err) {
        console.log(err);
      } else {
        res.redirect("../done");
        console.log(user + ":" + access_token.substring(0,10) + ":" + results['expires_in'] + ":" + refresh_token.substring(0,10));
        saveToken(user, access_token, results['expires_in'], refresh_token);
      }
    });

};
router.use('/callback', router.callback);

router.authorize = function (req, res) {
  
  var oauth = getOAuth();
  var authParams = {
    client_id: oauth.clientId,
    response_type: 'Assertion',
    state: req.query.user,
    scope: 'vso.chat_manage vso.dashboards vso.dashboards_manage vso.project_manage vso.work_write',
    redirect_uri: oauth.redirectUri,
  };
  res.redirect("https://" + oauth.baseUrl + oauth.authEndpoint + '?' + querystring.stringify(authParams));

};
router.use('/', router.authorize);

router.refreshToken = function (user, refresh, res) {
  
  router.newToken(user, refresh, true, (err, access_token, refresh_token, results) => {
    if(err) {
      console.log(err);
    } else {
      res.send("success");
      saveToken(user, access_token, results['expires_in'], refresh_token);
    }
  });
};


function saveToken(user, access_token, expires_in, refresh_token) {
  createConnection("save Token", (connection) => {
    var request = new tedious.Request(SAVE_TOKEN_QUERY, function (err) {
      if (err) {
        console.log(err);
      }
      connection.close();
    });
    request.addParameter('Id', TYPES.VarChar, user);
    request.addParameter('Token', TYPES.VarChar, access_token);
    request.addParameter('Expiry', TYPES.Int, expires_in);
    request.addParameter('Refresh', TYPES.VarChar, refresh_token);
    connection.execSql(request);
  });
}
