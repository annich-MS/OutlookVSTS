module.exports = {
    entry: ["babel-polyfill", "./src/client/Index"],
    output: {
        filename: "app.js",
        path: __dirname + "/public/js"
    },
    module: {
        rules: [
            {
                enforce: "pre",
                test: /\.tsx?$/,
                use: "source-map-loader"
            },
            {
                test: /\.tsx?$/,
                loader: 'awesome-typescript-loader',
                exclude: /node_modules/
            }

        ]

    },
    resolve: {
        // Add '.ts' and '.tsx' as resolvable extensions.
        extensions: [".ts", ".tsx", ".js"]
    },
    devtool: "inline-source-map",
};