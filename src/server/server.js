var path = require('path');
var express = require('express');
var bodyParser = require('body-parser');
var authenticate = require('./routes/authenticate');
var rest = require('./routes/VstsRest');
var PROD = process.env.prod;

var app = express();
var port = process.env.PORT || 3001;

app.use(bodyParser.json());         // to support JSON-encoded bodies
app.use(bodyParser.urlencoded({     // to support URL-encoded bodies
  extended: true,
  limit: '50mb'
}));
app.use(bodyParser.text({limit: '50mb'})); // support text bodies

// Routers

app.use('/authenticate', authenticate);
app.use('/rest', rest);
app.use('/log', function(req,res) {
  console.log(req.query.msg);
});
app.use('/public', express.static(path.join(__dirname, '../../public')));

app.get('*', (req, res) => {
  res.sendFile(path.join(__dirname, '../..', 'public/html', 'index.html'));
});

if(PROD != true){
  var https = require('https');
  var fs = require('fs');

  const options = {
    key: fs.readFileSync('secrets/key.pem'),
    cert: fs.readFileSync('secrets/cert.pem')
  };

  https.createServer(options, app).listen(port, function() {
    console.log('Listening at https://localhost:' + port);
  });
}
else{
  console.log("WARNING: you are not running in debug mode. nothing will work!");
  app.listen(port, "localhost", function(err) {
    if (err) {
      console.log(err);
      return;
    }
    console.log('Listening at http://localhost:' + port);
  });
}
