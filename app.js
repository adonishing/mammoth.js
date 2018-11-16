const koa = require('koa');
const static = require('koa-static');

const app = new koa();

app.use(static('/home/ubuntu/mammoth.js/browser-demo'));
               
const port = 8000
app.listen(port,function(){
    console.log('start on prot', port);
});