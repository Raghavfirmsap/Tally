const express = require('express')
const path = require('path')
const spawn = require("child_process").spawn;
const app = express()
const port = 3000
app.use(express.urlencoded());
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
})

app.get('/result', (req, res) => {
    res.send('Files have been downloaded.')
  })

app.post("/",function(req,res){
    var file=req.body.filePath;
    var fileName=req.body.fileName;
    var startdate = '20220401'
    var enddate = '20220431'
    var process = spawn('python3',['main.py', file, startdate, enddate, fileName]);
    process.stdout.on('data', function(data) {
    console.log(data.toString());
    })
    res.send(JSON.stringify(req.body));
});

app.listen(port, () => {
    console.log(`Example app listening on port ${port}`)
})
