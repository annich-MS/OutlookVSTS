var express = require('express');
var fs = require('fs');
var https = require('https');
var querystring = require('querystring');
var DEBUG = require('../debug');
var Authenticate = require('./Authenticate');
var Buffer = require('buffer').Buffer;

var router = express.Router({ mergeParams: true });
module.exports = router;
var API1_0 = "1.0";
var API2_0_Preview = '2.0-preview.1';
var API2_0 = "1.0";

function parseResponse(response, callback) {

  var str = '';

  //another chunk of data has been recieved, so append it to `str`
  response.on('data', function (chunk) {
    str += chunk;
  });
  //the whole response has been recieved, so we just print it out here
  response.on('end', function () {
    callback(str);
  });
};

function getRequest(query, path, headers, host, callback) {
  var options = {
    host: host,
    path: encodeURI(path) + '?' + querystring.stringify(query),
    method: 'GET',
    headers: headers
  };
  console.log('get: ' + options.path)
  https.get(options, function (response) { parseResponse(response, callback); console.log("response:" + response)});
}

function patchRequest(query, body, path, headers, host, callback) {
  headers['Content-Length'] = body.length;
  var options = {
    host: host,
    path: encodeURI(path) + '?' + querystring.stringify(query),
    method: 'PATCH',
    headers: headers
  };
  console.log('PATCH: ' + options.path);
  var request = https.request(options, function (response) { parseResponse(response, callback) });
  request.write(body);
  request.end();
}

function postRequest(query, body, path, headers, host, callback) {
  headers['Content-Length'] = body.length;
  var options = {
    host: host,
    path: encodeURI(path) + '?' + querystring.stringify(query),
    method: 'POST',
    headers: headers
  };
  console.log('POST: ' + host + ' ' + options.path);
  var request = https.request(options, function (response) { parseResponse(response, callback) });
  request.write(body);
  console.log('Body length:' + body.length);
  request.end();
}

function wrapToken(token) {
  return 'Bearer ' + token;
}

function getToken(user, callback) {
  Authenticate.getToken(user, (output) => {
    if (output.success) {
      callback(output.data.token);
    } else {
      callback("");
    }
  });
}

function jsonPatchItem(path, value) {
  return { 'op': 'add', 'path': path, 'value': value };
}

router.getItem = function (req, res) {
  var input = req.query;
  var query = {
    "ids": input.ids,
    "fields": input.fields,
    'api-version': API1_0
  };
  var host = input.instance + '.visualstudio.com';
  var path = '/DefaultCollection/_apis/wit/workitems';
  var headers = {};
  getToken(input.user, (token) => {
    if (token) {
      headers.Authorization = wrapToken(token);
      getRequest(query, path, headers, host, (output) => { res.send(output) });
    } else {
      console.log("could not find token for user " + input.user);
    }
  });
}
router.use('/getItem', router.getItem);

router.me = function (req, res) {
  var input = req.query;
  var query = {
    'api-version': API1_0
  };
  var host = 'app.vssps.visualstudio.com';
  var path = '/_apis/profile/profiles/me';
  var headers = {};
  getToken(input.user, (token) => {
    if (token) {
      headers.Authorization = wrapToken(token);
      getRequest(query, path, headers, host, (output) => { res.send(output) });
    } else {
      console.log('could not find token for user ' + input.user);
    }
  });
}
router.use('/me', router.me);

router.accounts = function (req, res) {
  var input = req.query;
  var query = {
    'memberId': input.memberId, //is this the correct way to get ID?
    'api-version': API1_0
  };
  var host = 'app.vssps.visualstudio.com';
  var path = '/_apis/Accounts';
  var headers = {};
  getToken(input.user, (token) => {
    if (token) {
      headers.Authorization = wrapToken(token);
      getRequest(query, path, headers, host, (output) => { res.send(output); console.log("json:"+output.JSON)});
    } else {
      console.log("could not find token for user " + input.user);
    }
  });
}
router.use('/accounts', router.accounts);

router.projects = function (req, res) {
  var input = req.query;
  var query = {
    'api-version': "1.0"
  };
  var host = input.account + ".visualstudio.com";
  var path = '/DefaultCollection/_apis/projects';
  var headers = {};
  getToken(input.user, (token) => {
    if (token) {
      headers.Authorization = wrapToken(token);
      getRequest(query, path, headers, host, (output) => { res.send(output) });
    } else {
      console.log("could not find token for user " + input.user);
    }
  })
}
router.use('/projects', router.projects);

router.getTeams = function (req, res) {
  var input = req.query;
  var query = {
    'api-version': API1_0
  }
  var host = input.account + '.visualstudio.com';
  var path = '/DefaultCollection/_apis/projects/' + input.project + '/teams';
  var headers = {};
  getToken(input.user, (token) => {
    if (token) {
      headers.Authorization = wrapToken(token);
      getRequest(query, path, headers, host, (output) => { res.send(output) });
    } else {
      console.log("could not find token for user " + input.user);
    }
  })
}
router.use('/getTeams', router.getTeams);

router.getTeamField = function (req, res) {
  var input = req.query;
  var query = {
    'api-version': API2_0_Preview
  }
  var host = input.account + '.visualstudio.com';
  var path = '/DefaultCollection/' + input.project + '/' + input.team + '/_apis/Work/TeamSettings/TeamFieldValues';
  var headers = {};
  getToken(input.user, (token) => {
    if (token) {
      headers.Authorization = wrapToken(token);
      getRequest(query, path, headers, host, (output) => { res.send(output) });
    } else {
      console.log("could not find token for user " + input.user);
    }
  })
}
router.use('/getTeamField', router.getTeamField);

router.createBug = function (req, res) {
  console.log ('in vstsrest now bug');
  var input = req.query;
  var query = {
    'api-version': API1_0
  }
  var host = input.account + '.visualstudio.com';
  var path = '/DefaultCollection/' + input.project + '/_apis/wit/workitems/$Bug'
  var headers = {
    'Content-Type': 'application/json-patch+json'
  };
  var body = JSON.stringify([
    jsonPatchItem('/fields/System.Title', input.title),
    jsonPatchItem('/fields/Microsoft.VSTS.TCM.ReproSteps', input.body),
    jsonPatchItem('/fields/System.AreaPath', input.areaPath)
  ]);
  getToken(input.user, (token) => {
    if (token) {
      headers.Authorization = wrapToken(token);
      patchRequest(query, body, path, headers, host, (output) => { res.send(output) });
    } else {
      console.log("could not find token for user " + input.user);
    }
  })
}
router.use('/newBug', router.createBug);

router.newWorkItem = function (req, res) {
  console.log ('in vstsrest now');
  var input = req.query;
  var query = {
    'api-version': API1_0 
  }
  var host = input.account + '.visualstudio.com';
  var path = '/DefaultCollection/' + input.project + '/_apis/wit/workitems/$' + input.type
  var headers = {
    'Content-Type': 'application/json-patch+json'
  };
  console.log('created the path');
  if (input.type === "Bug"){
   var body = JSON.stringify([
    jsonPatchItem('/fields/System.Title', input.title),
    jsonPatchItem('/fields/Microsoft.VSTS.TCM.ReproSteps', input.body),
    jsonPatchItem('/fields/System.AreaPath', input.areaPath),
    jsonPatchItem('/fields/System.IterationPath', input.currentIteration),
  ]);
  }
  else { 
    var body = JSON.stringify([
    jsonPatchItem('/fields/System.Title', input.title),
    jsonPatchItem('/fields/System.Description', input.body),
    jsonPatchItem('/fields/System.AreaPath', input.areaPath),
    jsonPatchItem('/fields/System.IterationPath', input.currentIteration),
  ]);

  }

  getToken(input.user, (token) => {
    if (token) {
      headers.Authorization = wrapToken(token);
      patchRequest(query, body, path, headers, host, (output) => { res.send(output) });
    } else {
      console.log("could not find token for user " + input.user);
    }
  })
}
router.use('/newWorkItem', router.newWorkItem);

router.getCurrentIteration = function (req, res) {
  console.log ('in vstsrest iteration');
  var input = req.query;
  var query = {
    '$timeframe': 'current',
    'api-version': '2.0-preview.1'
  }
  var host = input.account + '.visualstudio.com';
  var path = '/defaultcollection/' + input.project + '/' + input.team + '/_apis/work/teamsettings/iterations'
  var headers = {}; 
    getToken(input.user, (token) => {
    if (token) {
      console.log('good token');
      headers.Authorization = wrapToken(token);
      getRequest(query, path, headers, host, (output) => { res.send(output) });
    } else {
      console.log("could not find token for user " + input.user);
    }
  })
}
router.use('/getCurrentIteration', router.getCurrentIteration);

router.uploadAttachment = function (req, res) {
  var input = req.query;
  console.log(req);
  var body = req.body;
  var query = {
    'fileName': input.title + '.eml',
    'api-version': API1_0,
  }
  var host = input.account + '.visualstudio.com';
  var path = '/DefaultCollection/_apis/wit/attachments'
  var headers = {
    'Content-Type': 'application/json'
  };
  function decodeBase64Data(base64Data) {
    var myBuffer = new Buffer(base64Data, 'base64');
    return myBuffer.toString('utf8');
  }
  
  req.on('data', function(chunk) {
    body += chunk;
  }).on('end', function() {
    var decodedBody = decodeBase64Data(body);
    console.log('ATTACHMENT body length: ' + decodedBody.length);
    console.log('calling getToken for user: ' + input.user)
    getToken(input.user, (token) => {
      if (token) { 
        console.log('got token');
        headers.Authorization = wrapToken(token);
        postRequest(query, decodedBody, path, headers, host, (output) => {
          console.log('Attachment request accepted');
          console.log(JSON.stringify(output));
          res.send(output);
        });
      } else {
        console.log("could not find token for user " + input.user);
      }
    });
  });
}
router.use('/uploadAttachment', router.uploadAttachment);

router.attachAttachment = function (req, res) {
  var input = req.query;
  var query = {
    'api-version': API1_0,
  }
  var host = input.account + '.visualstudio.com';
  var path = '/DefaultCollection/_apis/wit/workitems/' + input.id
  var headers = {
    'Content-Type': 'application/json-patch+json'
  };
   var body = 
   JSON.stringify([
 {
    "op": "add",
    "path": "/relations/-",
    "value": {
      "rel": "AttachedFile",
      "url": input.attachmenturl,
      "attributes": {
      }
    }
  }
]);
  getToken(input.user, (token) => {
    if (token) {
      headers.Authorization = wrapToken(token);
      patchRequest(query, body, path, headers, host, (output) => { res.send(output) });
    } else {
      console.log("could not find token for user " + input.user);
    }
  })
}
router.use('/attachAttachment', router.attachAttachment);

