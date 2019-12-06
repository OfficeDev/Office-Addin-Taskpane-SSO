const { CleanWebpackPlugin } = require("clean-webpack-plugin");
const CopyWebpackPlugin = require("copy-webpack-plugin");
const HtmlWebpackPlugin = require("html-webpack-plugin");
const path = require("path");
const webpack = require("webpack");

module.exports = async (env, options) => {
    const dev = options.mode === "development";
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
                    use: "file-loader"
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
            new CopyWebpackPlugin([
                {
                    to: "taskpane.css",
                    from: "./src/taskpane/taskpane.css"
                }
            ]),
            new CopyWebpackPlugin([
                {
                    to: "assets",
                    from: "./assets"
                }
            ])
        ]
    };

    return config;
};