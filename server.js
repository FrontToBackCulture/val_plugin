const Express = require("express");
const HTTPS = require("https");
const HTTP = require("http");
const FS = require("fs");
const request = require("request");
var app = Express();

var path = require('path');
// const valURL = `http://ec2-18-139-42-208.ap-southeast-1.compute.amazonaws.com`;
// const valURL = `http://localhost:8080`;
const valURL = `https://val.fronttobackculture.com`
app.use('/', Express.static(path.join(__dirname, '/src/taskpane')));
app.use('/src', Express.static(path.join(__dirname, 'src')));
app.use('/node_modules', Express.static(path.join(__dirname, 'node_modules')));
app.use('/build', Express.static(path.join(__dirname, 'build')));

// app.get("/excel/", (request, response, next) => {
//     response.sendFile(__dirname + '/src/taskpane/taskpane.html');
// });

app.get("/excel/pullFullData", (req, res) => {
    console.log("PULLING FULL DATA ===========>")
    console.log(req.query)
    let options = {
        // url: "http://val.fronttobackculture.com/api/1/repo/retrieve_records",
        url: `${valURL}/api/1/repo/retrieve_records`,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify(req.query),
        method: "POST"
    }
    request(options, (error, response, body) => {
        console.log("HUEHUEHUEHEUHEUH===============>")
        res.send(body)
    })

});

app.get("/excel/pullPartialData", (req, res) => {
    let options = {
        // url: "val.fronttobackculture.com/api/1/repo/retrieve_records",
        url: `${valURL}/api/1/repo/retrieve_records`,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify(req.query),
        method: "POST"
    }
    request(options, (error, response, body) => {
        console.log("HUEHUEHUEHEUHEUH===============>")
        res.send(body)
    })
});
app.get("/excel/saveMapping", (req, res) => {
    let options = {
        // url: "https://val.fronttobackculture.com/api/1/saveSettings",
        url: `${valURL}/api/1/saveSettings`,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify(req.query),
        method: "POST"
    }
    request(options, (error, response, body) => {
        console.log("SETTINGS SAVED!===============>")

        res.send(body)
    })
});
app.get("/excel/retrieveMapping", (req, res) => {
    let options = {
        // url: "https://val.fronttobackculture.com/api/1/getsettings",
        url: `${valURL}/api/1/getsettings`,
        headers: { 'content-type': 'application/json' },
        qs: req.query,
        method: "GET"
    }

    request(options, (error, response, body) => {
        console.log("SETTINGS RETRIEVED!!===============>")
        res.send(JSON.parse(body))
    })
});

app.get("/excel/getRepoDetails", (req, res) => {
    // request(`https://val.fronttobackculture.com/api/1/query/repo/getrepotable?api_token=${req.query.api_token}&repo_id=${req.query.repo_id}`, (err, resp, body) => {
    request(`${valURL}/api/1/query/repo/getrepotable?api_token=${req.query.api_token}&repo_id=${req.query.repo_id}`, (err, resp, body) => {
        res.send(JSON.parse(body))
    })
})

app.get("/excel/getRepoTypes", (req, res) => {
    // request(`https://val.fronttobackculture.com/api/1/query/repo/getrepotype?api_token=${req.query.api_token}`, (err, resp, body) => {
    request(`${valURL}/api/1/query/repo/getrepotype?api_token=${req.query.api_token}`, (err, resp, body) => {
        console.log("==================>getRepoTypes")
        console.log(err)
        res.send(JSON.parse(body))
    })
})
app.get("/excel/getUserProjects", (req, res) => {
    console.log("GETTING PROJECTS ===========>")
    // request(`https://val.fronttobackculture.com/api/1/getAllProjects?api_token=${req.query.api_token}`, (err, resp, body) => {
    request(`${valURL}/api/1/getAllProjects?api_token=${req.query.api_token}`, (err, resp, body) => {
        res.send(JSON.parse(body))
    })
})

app.get("/excel/getUserPhases", (req, res) => {
    console.log("GETTING PHASE ===========>")
    let propertiesObject = {
        api_token: req.query.api_token
    }
    let id = 'all';
    let url = `${valURL}/api/1/query/phase/getphase/${id}`
    request({ url: url, qs: propertiesObject }, (err, resp, body) => {
        res.send(JSON.parse(body))
    })
})

app.get("/excel/updateRecord", (req, res) => {
    console.log("=====> sending update rest api", req.query)
    var options = {
        "method": "PATCH",
        "url": `${valURL}/api/1/audit/updateViaExternal`,
        "headers": {
            "Cache-Control": "no-cache",
            "Content-Type": "application/x-www-form-urlencoded"
        },
        "form": req.query

    }

    request(options, function (error, response, body) {
        if (error) {
            console.log(error)
            throw new Error(error);
        }
        res.send();
    });
})

app.get("/excel/login", (req, res) => {
    let options = {
        // url: "https://val.fronttobackculture.com/api/1/login",
        url: `${valURL}/api/1/login`,
        headers: { 'content-type': 'application/json' },
        body: JSON.stringify(req.query),
        method: "POST"
    }
    request(options, (error, response, body) => {
        console.log("HUEHUEHUEHEUHEUH===============>")
        console.log(body)
        console.log(error)
        if (typeof body == "string") {
            body = JSON.parse(body)
        }
        res.send(body)

    })

});

HTTPS.createServer({
    key: FS.readFileSync("server.key"),
    cert: FS.readFileSync("server.crt")
}, app).listen(9000, () => {
    console.log("Listening at :9000...");
});

// HTTP.createServer(app).listen(3000, () => {
//     console.log("Listening at :3000...");
// });