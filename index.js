// Include async module by absolute module install path.
var async = require('async');
var fs = require('fs');
var program = require('commander');
var handlebars = require('handlebars');
var XLSX = require('xlsx');
var nodemailer = require('nodemailer');

program
  .version('0.1.0')
  .option('-h, --host [host]', 'Host')
  .option('-u, --user [user]', 'User')
  .option('-p, --pass [pass]', 'Password')
  .option('--port <port>', 'Port', 465)
  .option('--secure', 'Secure', true)
  .option('-s, --subject', 'Subject', 'My Subject')
  .option('-i, --input-file [inputFile]', 'Recipients file')
  .option('-t, --template [template]', 'Template', 'email')
  .option('-w, --workers <workers>', 'Workers number', 2)
  .option('-c, --init [init]', 'Init spreadsheet')
  .option('-m, --import [import]', 'Import spreadsheet')
  .parse(process.argv);

if (program.init) {
    var wb2 = {SheetNames:["Recipients"], Sheets:{Recipients:XLSX.utils.aoa_to_sheet([['name', 'email']])}};
    /* write it */
    XLSX.writeFile(wb2, program.init);
    console.log("Created", program.init);
    process.exit(0);
} else if (program.import) {
    importXlsx(program.import);
    console.log("Imported", program.import, "to", program.import.replace(/xlsx$/, 'json'));
    process.exit(0);
} else if (!program.inputFile) {
    program.outputHelp();
    process.exit(0);
}

var transporter = nodemailer.createTransport({
    host: program.host,
    port: program.port,
    secure: program.secure, // true for 465, false for other ports
    auth: {
        user: program.user,
        pass: program.pass
    }
});

function importXlsx(recipientsFilename) {
    if (!recipientsFilename.endsWith('.xlsx')) {
        console.log(recipientsFilename, "is not a valid xlsx file");
        process.exit(0);
    }
    var wb = XLSX.readFile(recipientsFilename, {cellDates:true});
    allRecipients = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {raw:true, header:["name", "email"]});
    allRecipients.shift();        
    saveRecipients(recipientsFilename, allRecipients)
}

function loadRecipients(recipientsFilename) {
    var allRecipients;
    if (recipientsFilename.endsWith('.xlsx')) {
        var wb = XLSX.readFile(recipientsFilename, {cellDates:true});
        allRecipients = XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]], {raw:true, header:["name", "email"], skipHeader: true});
        allRecipients.shift();
    } else {
        allRecipients = JSON.parse(fs.readFileSync(recipientsFilename));
    }
    var uncompleted = 0;
    for (var obj of allRecipients) {
        if (!!!obj.completed) {
            uncompleted++;
        }
    }
    console.log("Loaded", allRecipients.length, "contacts from", recipientsFilename);
    console.log("About to send", uncompleted, "emails...");
    return allRecipients;
} 

function saveRecipients(recipientsFilename, data) {
    if (recipientsFilename.endsWith('.xlsx')) {
        recipientsFilename = recipientsFilename.replace(/xlsx$/, 'json');
    } 
    fs.writeFileSync(recipientsFilename, JSON.stringify(data));
} 

function isCompleted(obj) {
    return !!obj.completed;
}

function loadTemplate(name, ext) {
    console.log("Loading template", name+ext);
    var source = fs.readFileSync(name+ext);
    return handlebars.compile(source.toString(), { strict: true });
}

var textTemplate = loadTemplate(program.template,'.txt');
var htmlTemplate = loadTemplate(program.template,'.html');

// Create the queue object. The first parameter is a function object.
// The second parameter is the worker number that execute task at same time.
function sendEmail(object,callback) {

    // Get queue start run time.
    var date = new Date();
    var time = date.toTimeString();

    // Print task start info.
    console.log("Start task " + JSON.stringify(object) + " at " + time);

    var mail = {
        from: program.user,
        to: object.email,  //Change to email address that you want to receive messages on
        subject: program.subject,
        text: textTemplate(object),
        html: htmlTemplate(object)
    };

    transporter.sendMail(mail, (err, data) => {
        if (err) {
            console.log("Task " + JSON.stringify(object) + " failed ");
            callback(err);
        } else {
            // Get timeout time.
            date = new Date();
            time = date.toTimeString();

            // Print task timeout data.
            console.log("End task " + JSON.stringify(object) + " at " + time);
            object.completed = time;
            callback();
        }
    });
};

function main() {
    var queue = async.queue(sendEmail, program.workers);
    var recipients = loadRecipients(program.inputFile);
    
    // Loop to add object in the queue with prefix 2.
    for (var obj of recipients) {
        if (!isCompleted(obj)) {
            queue.push(obj,function (err) {
                if (err) { console.log(err); }
                else {
                    saveRecipients(program.inputFile, recipients);
                }
                //console.log(recipients);
            });    
        }
    }    
}

main();