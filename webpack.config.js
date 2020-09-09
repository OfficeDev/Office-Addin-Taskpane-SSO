const { CleanWebpackPlugin } = require("clean-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require("path");
const webpack = require("webpack");

const urlDev="https://localhost:3000/";
const urlProd="https://www.contoso.com/"; // CHANGE THIS TO YOUR PRODUCTION DEPLOYMENT LOCATION

module.exports = async (env, options) => {
    const dev = options.mode === "development";
    const buildType = dev ? "dev" : "prod";
    const config = {
        devtool: "source-map",
        entry: {
            commands: "./src/commands/commands.ts",
            fallbackauthdialog: "./src/helpers/fallbackauthdialog.ts",
            polyfill: "@babel/polyfill",
            taskpane: "./src/taskpane/taskpane.ts",
        },
        output: {
            path: path.resolve(process.cwd(), 'dist'),
        },
        resolve: {
            extensions: [".ts", ".tsx", ".html", ".js"]
        },
        module: {
            rules: [
                {
                    test: /\.ts$/,
                    exclude: /node_modules/,
                    use: "babel-loader"
                },
                {
                    test: /\.tsx?$/,
                    exclude: /node_modules/,
                    use: "ts-loader"
                },
                {
                    test: /\.html$/,
                    exclude: /node_modules/,
                    use: "html-loader"
                },
                {
                    test: /\.(png|jpg|jpeg|gif)$/,
                    loader: "file-loader",
                    options: {
                      name: '[path][name].[ext]'
                    }
                }
            ]
        },
        plugins: [
            new CleanWebpackPlugin(),
            new HtmlWebpackPlugin({
                filename: "taskpane.html",
                template: "./src/taskpane/taskpane.html",
                chunks: ["polyfill", "taskpane"]
            }),
            new HtmlWebpackPlugin({
                filename: "commands.html",
                template: "./src/commands/commands.html",
                chunks: ["polyfill", "commands"]
            }),
            new HtmlWebpackPlugin({
                filename: "fallbackauthdialog.html",
                template: "./src/helpers/fallbackauthdialog.html",
                chunks: ["polyfill", "fallbackauthdialog"]
            }),
            new CopyWebpackPlugin({
                patterns: [
                {
                  to: "taskpane.css",
                  from: "./src/taskpane/taskpane.css"
                },
                {
                  to: "[name]." + buildType + ".[ext]",
                  from: "manifest*.xml",
                  transform(content) {
                    if (dev) {
                      return content;
                    } else {
                      return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
                    }
                  }
                }
              ]})
        ]
    };

    return config;
};