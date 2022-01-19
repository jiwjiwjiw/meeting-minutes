const path = require('path')
const GasPlugin = require('gas-webpack-plugin')
const CleanPlugin = require('clean-webpack-plugin')
const CopyPlugin = require('copy-webpack-plugin')

const config = {
  mode: 'production',
  entry: './src/index.ts',
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: ['ts-loader'],
        exclude: /node_modules/
      }
    ]
  },
  resolve: {
    extensions: ['.tsx', '.ts', '.js']
  },
  output: {
    path: path.resolve(__dirname, 'dist'),
    filename: 'bundle.js'
  },
  plugins: [
    new GasPlugin(),
    new CleanPlugin.CleanWebpackPlugin(),
    new CopyPlugin({
      patterns: [
        {
          from: 'src/**/*.html',
          to: '[name][ext]'
        }
      ]
    })
  ]
}

module.exports = config
