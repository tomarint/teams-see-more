// Generated using webpack-cli https://github.com/webpack/webpack-cli

const path = require("path");
const CopyPlugin = require("copy-webpack-plugin");

const isProduction = process.env.NODE_ENV == "production";

const config = {
  entry: {
    background: path.resolve(__dirname, "src/background.ts"),
    foreground: path.resolve(__dirname, "src/foreground.ts")
  },
  output: {
    path: path.resolve(__dirname, "dist"),
    filename: "[name].js"
  },
  devtool: false,  // not to use eval in debug build
  plugins: [
    new CopyPlugin({
      patterns: [
        {
          from: "src",
          globOptions: {
            ignore: ["**/*.ts"]
          },
        }
      ],
    }),
  ],
  module: {
    rules: [
      {
        test: /\.(ts|tsx)$/i,
        loader: "ts-loader",
        exclude: ["/node_modules/"],
      },
    ],
  },
  resolve: {
    extensions: [".tsx", ".ts", ".js"],
  },
};

module.exports = () => {
  if (isProduction) {
    config.mode = "production";
  } else {
    config.mode = "development";
  }
  return config;
};
