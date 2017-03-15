import authenticate from "./routes/authenticate";
import rest from "./routes/vstsRest";
let path = require("path");
let express = require("express");
let bodyParser = require("body-parser");

const DevMode: boolean = process.env.NODE_ENV === "development";

let app = express();
let port = process.env.PORT || 3001;

app.use(bodyParser.json());         // to support JSON-encoded bodies
app.use(bodyParser.urlencoded({     // to support URL-encoded bodies
  extended: true,
  limit: "50mb"
}));
app.use(bodyParser.text({ limit: "50mb" })); // support text bodies

// Routers

app.use("/authenticate", authenticate);
app.use("/rest", rest);
app.use("/log", function (req, res) {
  console.log(req.query.msg);
});
app.use("/public", express.static(path.join(__dirname, "../public")));

app.get("*", (req, res) => {
  res.sendFile(path.join(__dirname, "..", "public/html", "index.html"));
});

if (DevMode) {
  let https = require("https");
  let fs = require("fs");

  const options = {
    cert: fs.readFileSync("cert.pem"),
    key: fs.readFileSync("key.pem"),
  };

  https.createServer(options, app).listen(port, function () {
    console.log("Listening at https://localhost:" + port);
  });
}
else {
  console.log("WARNING: you are not running in debug mode. nothing will work!");
  app.listen(port, "localhost", function (err) {
    if (err) {
      console.log(err);
      return;
    }
    console.log("Listening at http://localhost:" + port);
  });
}
