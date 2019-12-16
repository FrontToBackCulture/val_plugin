// const path = require("path");
// const webpack = require("webpack");
// const webpack_rules = [];
// const webpackOption = {
//     // entry: "src/taskpane/taskpane.js",
//     entry: ["@babel/polyfill", "./src/taskpane/taskpane.js" ],
//     output: {
//         path: path.resolve(__dirname + '/build', ""),
//         filename: "taskpane.bundle.js",
//     },
//     module: {
//         rules: webpack_rules
//     }
// };
// let babelLoader = {
//     test: /\.js$/,
//     exclude: /(node_modules|bower_components)/,
//     use: {
//         loader: "babel-loader",
//         options: {
//             presets: ["@babel/preset-env"]
//         }
//     }
// };
// webpack_rules.push(babelLoader);
// module.exports = webpackOption;

const path = require('path');

module.exports = {
	devtool: "inline-source-map",
	entry: {
		taskpane: [
			'@babel/polyfill',
			'./src/taskpane/taskpane.js',
			// './src/taskpane/popup.js'
		],
		popup: ['@babel/polyfill','./src/taskpane/popup.js']
	},
	mode: 'development',
	output: {
		filename: '[name].bundle.js',
		path: path.resolve(__dirname + '/build', "")
	},
	module: {
		rules: [{
			test: /\.js$/,
			exclude: /node_modules/,
			use: {
				loader: 'babel-loader',
				options: {
					presets: [
						["@babel/preset-env",
						 {
							 targets: {
								"ie": "11"
							 },
							 useBuiltIns: "usage"
						}]
					]
				}
			}
			
		}]
	},
};