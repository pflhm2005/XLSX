const path = require('path');

module.exports = {
  entry: './lib/index.js',
  output: {
    filename: 'dist.min.js',
    path: path.resolve(__dirname),
  },
  mode: 'production'
}