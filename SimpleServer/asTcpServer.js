var net = require('net');
const fetch = require("node-fetch");

var textChunk = '';
var server = net.createServer(function(socket) {
	socket.on('data', function(data){
      console.log(data);
      var mostRecentMessage = require('./most-recent-message.json');
      console.log("Received: ", data.toString("utf8"));
		  socket.write(JSON.stringify(mostRecentMessage));
    });
});
server.listen(59090, '127.0.0.1');
console.log('server running on port 59090');