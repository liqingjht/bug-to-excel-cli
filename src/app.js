var request = require("request");
var cheerio = require("cheerio");
var Excel = require('exceljs');
var colors = require("colors");
var program = require("commander");
var readlineSync = require('readline-sync');
var Agent = require('agentkeepalive');
var ProgressBar = require('progress');
var fs = require('fs');

var putError = console.log;
global.console.error = function(error) {
     putError(colors.red("[Error]: " + error));
}

program.on('--help', function() {
    console.log('  \n  Examples:\n');
    console.log('    node app.js -u "http://..." -s Name');
})

program.option('-u, --url <url>', 'Url of bug list to generate.')
    .option('-s, --specifyName ', 'Specify string of file name.')
    .parse(process.argv);

var fileName = "BugList-" + (new Date()).toLocaleDateString() +  (program.specifyName? "-"+program.specifyName : "");

var url = "";
if (!program.url) {
    var count = 0;
    while (url == "") {
        if (++count > 3) {
            program.outputHelp();
            process.exit(1);
        }
        url = readlineSync.question('Please input the url of bug list: ').trim().replace(/^"(.*)"$/g, "$1");
    }
}
else {
    url = program.url;
}

url = decodeURIComponent(url);
url = encodeURI(url);

var urlIndex = url.indexOf("/bugzilla3/");
if (urlIndex != -1) {
    var root = url.slice(0, urlIndex + 11);
}
else {
    var root = url.replace(/^((https?:\/\/)?[^\/]*\/).*/ig, "$1");
    root = (/^https?:\/\//ig.test(root) ? root : "http://" + root);
}
var bugUrl = root + "show_bug.cgi?ctype=xml&id=";

Agent = (root.toLowerCase().indexOf("https://") != -1)? Agent.HttpsAgent: Agent;
var keepaliveAgent = new Agent({
    maxSockets: 100,
    maxFreeSockets: 10,
    timeout: 60000,
    freeSocketKeepAliveTimeout: 30000
});

var option = {
    agent: keepaliveAgent,
    headers: {"User-Agent": "NodeJS", Host: url.replace(/^((https?:\/\/)?([^\/]*)\/).*/g, "$3")},
    url: url
};

getFunc(option, function (url, $) {
    var bugs = new Array();
    var td = $("table.bz_buglist tr td.bz_id_column a");
    td.each(function (key) {
        bugs.push(td.eq(key).text());
    })
    if (bugs.length > 0) {
        console.log("");
        global.bar = new ProgressBar('Getting Bugs [:bar] :percent | ETA: :etas | :current/:total', {
            complete: "-",
            incomplete: " ",
            width: 25,
            clear: false,
            total: bugs.length,
        });
    }
    else {
        console.error("No bugs can be found.");
        process.exit(1);
    }
    return bugs; 
}).then(function (bugs) {
    var done = 0;
    return Promise.all(bugs.map(function (eachBug, index) {
        option.url = bugUrl + eachBug;
        var promiseGetOne = getFunc(option, function (url, $) {
            var oneInfo = new Object();
            oneInfo.url = url.replace(/ctype=xml&/ig, "");
            oneInfo.id = $("bug_id").text();
            oneInfo.summary = $("short_desc").text();
            oneInfo.reporter = $("reporter").text();
            oneInfo.product = $("product").text();
            oneInfo.component = $("component").text();
            oneInfo.version = $("version").text();
            oneInfo.status = $("bug_status").text();
            oneInfo.priority = $("priority").text();
            oneInfo.security = $("bug_security").text();
            oneInfo.assign = $("assigned_to").text();
            oneInfo.comment = new Array();
            var comments = $("long_desc");
            comments.each(function (key) {
                var who = comments.eq(key).find("who").text();
                var when = comments.eq(key).find("bug_when").text();
                when = when.replace(/([^\s]+)\s.*$/g, "$1");
                var desc = comments.eq(key).find("thetext").text();
                if (key == 0 && who == oneInfo.reporter) {
                    oneInfo.detail = desc;
                    return true;
                }
                oneInfo.comment.push({ 'who': who, 'when': when, 'desc': desc });
            })

            return oneInfo;
        })

        promiseGetOne.then(function () {
            done++;
            bar.tick();
            if (done == bugs.length) {
                console.log("\n");
            }
        })

        return promiseGetOne;
    }))
}).then(function (bugLists) {
    var workbook = new Excel.Workbook();
    var productNum = 0;

    for (var i in bugLists) {
        bugInfo = bugLists[i];

        var sheet = workbook.getWorksheet(bugInfo.product);
        if (sheet === undefined) {
            sheet = workbook.addWorksheet(bugInfo.product);
            productNum++;
        }

        try {
            sheet.getColumn("id");
        }
        catch (error) {
            sheet.columns = [
                { header: 'Bug ID', key: 'id' },
                { header: 'Summary', key: 'summary', width: 35 },
                { header: 'Bug Detail', key: 'detail', width: 75 },
                { header: 'Priority', key: 'priority', width: 8 },
                { header: 'Version', key: 'version', width: 15 },
                { header: 'Status', key: 'status', width: 15 },
                { header: 'Component', key: 'component', width: 15 },
                { header: 'Comments', key: 'comment', width: 60 },
                { header: 'Assign To', key: 'assign', width: 20 },
                { header: 'Reporter', key: 'reporter', width: 20 },
            ];
        }

        var comment = "";
        for (var j in bugInfo.comment) {
            comment += bugInfo.comment[j].who + " (" + bugInfo.comment[j].when + " ):\r\n";
            comment += bugInfo.comment[j].desc.replace(/\n/gm, "\r\n") + "\r\n";
            comment += "-------------------------------------------------------\r\n"
        }
        sheet.addRow({
            id: { text: bugInfo.id, hyperlink: bugInfo.url },
            summary: bugInfo.summary,
            detail: bugInfo.detail ? bugInfo.detail.replace(/\n/gm, "\r\n") : "",
            priority: bugInfo.priority,
            version: bugInfo.version,
            status: bugInfo.status,
            component: bugInfo.component,
            comment: comment,
            assign: bugInfo.assign,
            reporter: bugInfo.reporter,
        });

        sheet.eachRow(function (Row, rowNum) {
            Row.eachCell(function (Cell, cellNum) {
                if (rowNum == 1)
                    Cell.alignment = { vertical: 'middle', horizontal: 'center', size: 25, wrapText: true }
                else
                    Cell.alignment = { vertical: 'top', horizontal: 'left', wrapText: true }
            })
        })
    }

    fileName = ((productNum > 1) ? "" : bugInfo.product + "-") + fileName + ".xlsx";
    var files = fs.readdirSync("./");
    var postfix = 1;
    while (files.indexOf(fileName) != -1) {
        fileName = fileName.replace(/(\(\d+\))?\.xlsx$/g, "(" + (postfix++) + ").xlsx");
        if (postfix > 99) {
            console.warn("It may occur somethins wrong.");
            break;
        }
    }

    return workbook.xlsx.writeFile(fileName);
}).then(function() {
    console.log("Generate xlsx file successfully. Filename is " + colors.cyan(fileName));
}).catch(function(err) {
    console.error(err);
    process.exit(1);
})

function getFunc(getOption, parseFunc) {
    return new Promise(function(resolve, reject) {
        request.get(getOption, function(error, response, body) {
            if(!error && response.statusCode == 200) {
                var $ = cheerio.load(body);
                var result = parseFunc(getOption.url, $);
                resolve(result);
            }
            else {
                reject(error);
            }
        })
    })
}

