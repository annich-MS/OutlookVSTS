var webpack = require('webpack');
var path = require('path');

var APP_DIR = path.join(__dirname, '..', 'app');

module.exports = {
  entry: "../app/Index.tsx",
  output: {
    filename: "../app/build/app.js"
  },
  devtool: "source-map",
  resolve: {
    // Add '.ts' and '.tsx' as resolveable extensions.
    extensions: ["", ".webpack.js", ".web.js", ".ts", ".tsx", ".js"]
  },
  module: {
    loaders: [
      // All files with a '.ts' or '.tsx' extension will be handled by 'ts-loader'
      { test: /\.tsx?#/, loader: "ts-loader" },
      { test: /\.json#/, loader: "json" }
    ],
    preLoaders: [
      // All output '.js' files will have any sourcemaps re-processed by 'source-map-loader'.
      { test: /\.js$/, loader: "source-map-loader" }
    ]
  }
}