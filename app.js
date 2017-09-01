var express = require('express'),
		path    = require('path');

var app = express();

app.use(express.static(path.join(__dirname, './')));

//首页
app.get('/', function(req, res) {
	res.sendfile('index.html');
});


app.get('/example', function(req, res) {
	res.sendfile('example.html');
});

//获取excel数据
app.get('/xlsx', function(req, res) {
	res.sendfile('agenda.xlsx');
});

//启动app
app.listen(3000, function() {
	console.log('server is started.');
})

