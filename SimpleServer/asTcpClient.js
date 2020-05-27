var net = require('net');
var client = new net.Socket();
client.connect(59090, '127.0.0.1', function() {
	console.log('Connected');
	var mostRecentMessage = require('./most-recent-message.json');
	console.log("Sending...")
	client.write(JSON.stringify(mostRecentMessage));
});
client.on('data', function(data) {
	console.log('Answer: ' + JSON.parse(data));
	client.destroy();
});
client.on('close', function() {
	console.log('Connection closed');
});