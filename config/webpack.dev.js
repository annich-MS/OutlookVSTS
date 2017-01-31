
module.exports = {
  entry: "./src/client/Index.tsx",
  output: {
    filename: "./bin/app.js"
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