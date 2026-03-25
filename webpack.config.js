const path = require("path");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");

module.exports = async (env, argv) => {
  const isDev = argv.mode === "development";

  const devServerOptions = isDev
    ? {
        devServer: {
          port: 3000,
          server: {
            type: "https",
            options: await require("office-addin-dev-certs").getHttpsServerOptions(),
          },
          static: [
            {
              directory: path.join(__dirname, "assets"),
              publicPath: "/assets",
            },
            {
              directory: path.join(__dirname),
              publicPath: "/",
              serveIndex: false,
            },
          ],
          headers: {
            "Access-Control-Allow-Origin": "*",
          },
        },
      }
    : {};

  return {
    entry: {
      taskpane: "./src/index.tsx",
    },
    output: {
      path: path.resolve(__dirname, "dist"),
      filename: "[name].bundle.js",
      clean: true,
    },
    resolve: {
      extensions: [".ts", ".tsx", ".js", ".jsx"],
    },
    module: {
      rules: [
        {
          test: /\.tsx?$/,
          use: "ts-loader",
          exclude: /node_modules/,
        },
        {
          test: /\.css$/,
          use: ["style-loader", "css-loader"],
        },
      ],
    },
    plugins: [
      new HtmlWebpackPlugin({
        template: "./src/taskpane.html",
        filename: "taskpane.html",
        chunks: ["taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          { from: "assets", to: "assets" },
          { from: "manifest.xml", to: "manifest.xml" },
          { from: "src/index.html", to: "index.html" },
          { from: "src/auth-callback.html", to: "auth-callback.html" },
        ],
      }),
    ],
    ...devServerOptions,
    devtool: isDev ? "source-map" : false,
  };
};
