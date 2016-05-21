var path = require('path');
var express = require('express');
var webpack = require('webpack');
var bodyParser = require('body-parser');
var config = require('./config/webpack.prod');
var authenticate = require('./routes/authenticate');
var vsts = require('./routes/vsts');

var app = express();
var compiler = webpack(config);
var port = process.env.PORT || 3000;

app.use( bodyParser.json() );       // to support JSON-encoded bodies
app.use(bodyParser.urlencoded({     // to support URL-encoded bodies
  extended: true
}));

app.use(require('webpack-dev-middleware')(compiler, {
  noInfo: true,
  publicPath: config.output.publicPath
}));

app.use(require('webpack-hot-middleware')(compiler));

// Routes

var authRouter = express.Router({mergeParams: true});
var vstsRouter = express.Router({mergeParams: true});
app.get('/', function (req, res) {
    res.sendFile(__dirname + '/index.html');
});
app.use('/authenticate', authRouter);
authRouter.use('/callback', authenticate.callback);
authRouter.use('/', authenticate.authorize);
app.use('/VSTS', vstsRouter);
app.use('/js', express.static(__dirname + '/js'));
app.use('/css', express.static(__dirname + '/css'));

app.listen(port, 'localhost', err => {
  if (err) {
    console.log(err);
    return;
  }

  console.log(`Listening at http://localhost:${port}`);
});
