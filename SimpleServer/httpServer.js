'use strict'
const fs = require('fs');
const http = require('http');
const url = require('url');

http.createServer((req, res) => {
    res.writeHead(200, { 'Content-Type': 'text/plain' });
    res.write('Request received ' + req.url);
    var postData = '';
    var x = "Test";

    // Get all post data when receive data event.
    req.on('data', function (chunk) {
        postData += chunk;
    });

    req.on('end', function () {
        res.write("Sent data: " + postData.toString());
        fs.writeFileSync('most-recent-message.json', postData)
        res.end();
    });
}).listen(8080);

console.log('server running on port 8080');