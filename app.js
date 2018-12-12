const koa = require('koa');
const static = require('koa-static');
const path = require('path')
const Router = require('koa-router');
const request = require('request');
const bodyParser = require('koa-bodyparser');

const app = new koa();
var router = new Router();
app.use(bodyParser({
    enableTypes:['json', 'text']
}));


function ole2mml(id, b64){
    return new Promise((resolve, reject)=>{
        request.post('http://127.0.0.1:4567/ole2mml',{
            qs: { id},
            body: b64
        },function (error, response, body) {
            if (error) reject(new Error(error));
            resolve(body);
          })
    })
}

router.post('/ole2mml',async (ctx, next)=>{
    ctx.body = await ole2mml(ctx.request.query.id, ctx.request.body);
    ctx.type = 'text/xml';
})

app.use(static(path.join(__dirname, "browser-demo")));
app.use(router.routes());

const port = 8000
app.listen(port,function(){
    console.log('start on port', port);
});