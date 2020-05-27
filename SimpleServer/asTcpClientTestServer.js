var net = require('net');
const fetch = require("node-fetch");

var textChunk = '';
var server = net.createServer(function(socket) {
	socket.on('data', function(data){
    console.log(data);
    var textChunk = data.toString('utf8');
    console.log("Textchunk type: ", typeof textChunk);
    console.log("Received: ", textChunk);
		socket.write("You sent: ", JSON.stringify(textChunk));
    });
});
server.listen(59090, '127.0.0.1');
console.log('server running on port 59090');
