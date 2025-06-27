const HtmlWebpackPlugin = require('html-webpack-plugin');
const CopyWebpackPlugin = require('copy-webpack-plugin');
const path = require('path');

module.exports = {
  entry: {
    taskpane: './src/taskpane/taskpane.tsx',
    commands: './src/commands/commands.ts'
  },
  resolve: {
    extensions: ['.ts', '.tsx', '.js', '.jsx']
  },
  module: {
    rules: [
      {
        test: /\.tsx?$/,
        use: 'ts-loader',
        exclude: /node_modules/
      },
      {
        test: /\.css$/,
        use: ['style-loader', 'css-loader']
      }
    ]
  },
  plugins: [
    new HtmlWebpackPlugin({
      template: './src/taskpane/taskpane.html',
      filename: 'taskpane.html',
      chunks: ['taskpane']
    }),
    new HtmlWebpackPlugin({
      template: './src/commands/commands.html',
      filename: 'commands.html',
      chunks: ['commands']
    }),
    new CopyWebpackPlugin({
      patterns: [
        { from: './assets', to: 'assets', noErrorOnMissing: true },
        { from: './manifest.xml', to: 'manifest.xml' },
        { from: './sideload.html', to: 'sideload.html' },
        { from: './test-cert.html', to: 'test-cert.html' }
      ]
    })
  ],
  output: {
    filename: '[name].js',
    path: path.resolve(__dirname, 'dist'),
    clean: true
  },
  devServer: {
    port: 3000,
    https: true,
    hot: true,
    static: {
      directory: path.resolve(__dirname, 'dist'),
      publicPath: '/'
    }
  }
};