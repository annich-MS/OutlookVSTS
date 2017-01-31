var path = require('path');
var webpack = require('webpack');

var APP_DIR = path.join(__dirname, '..', 'src/client');

module.exports = {
  devtool: 'source-map',
  entry: './src/client/index.tsx',
  module: {
    preLoaders: [{
      test: /\.tsx?$/,
      loader: 'tslint',
      include: APP_DIR
    }],
    loaders: [{
      test: /\.tsx?$/,
      loaders: ['babel', 'ts'],
      include: APP_DIR
    },
      { test: /\.css$/, loader: "style-loader!css-loader" }]
  },
  output: {
    path: path.join(__dirname, '..', 'static'),
    filename: 'app.js',
    publicPath: '/static/'
  },
  plugins: [
    new webpack.optimize.OccurrenceOrderPlugin(),
    new webpack.DefinePlugin({
      'process.env': {
        'NODE_ENV': JSON.stringify('production')
      }
    })/*,
    new webpack.optimize.UglifyJsPlugin({
      compressor: {
        warnings: false
      }
    })*/
  ],
  resolve: {
    root: [path.resolve('../src/client')],
    extensions: ['', '.jsx', '.js', '.tsx', '.ts', '.css']
  },
  tslint: {
    emitErrors: false,
    failOnHint: false
  }
}
