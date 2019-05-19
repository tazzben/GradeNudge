function doGet(e) {
    var template = HtmlService.createTemplateFromFile('Page');
    return template.evaluate()
        .setTitle('Grade Nudge')
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}

function getOAuthToken() {
    var returnData = {};
    var scriptProperties = PropertiesService.getScriptProperties();
    returnData.dev_key = scriptProperties.getProperty('DEVELOPER_KEY');
    DriveApp.getRootFolder();
    returnData.token = ScriptApp.getOAuthToken();
    return returnData;
}
function capitalizeFirstLetter(s) {
    return s.charAt(0).toUpperCase() + s.slice(1);
}

function CheckData(data) {
    return interfaceClass.submitWindow(data);
}

function SendGradebook(dataarray) {
    var returnData = {};
    returnData.emailCol = -1;
    returnData.pointsCol = [];
    returnData.nameCol = -1;
    returnData.needsDomain = false;
    returnData.domainDefault = "";
    returnData.fileID = "";
    returnData.fileName = "";

    var savedfile = false;
    var listToDel = [];
    for (var i = 0; i < dataarray.files.length; i++) {
        var sFile = dataarray.files[i];
        var data_index = sFile.blob.indexOf('base64') + 7;
        var filedata = sFile.blob.slice(data_index, sFile.blob.length);
        var decoded = Utilities.base64Decode(filedata);
        var resource = {
            title: sFile.name,
            mimeType: sFile.type
        };
        var b = Utilities.newBlob(decoded, sFile.type, sFile.name);
        var file = Drive.Files.insert(resource, b, {
            convert: true
        });
        if (walkSheet.checkFileType(file.getId(), MimeType.GOOGLE_SHEETS)) {
            var resultofSearch = processUploadedFile.findEmailDictionary(file.getId());
            if (resultofSearch) {
                var savedfile = file.getId();
                f = DriveApp.getFileById(file.getId());
                returnData.fileName = f.getName();
                f.setTrashed(true);
            } else {
                listToDel.push(file.getId());
            }
        } else {
            Drive.Files.remove(file.getId());
        }
    }

    if (savedfile === false && processUploadedFile.checkSheetList.length > 0) {

        for (var i = 0; i < processUploadedFile.checkSheetList.length; i++) {
            var fileproc = processUploadedFile.addDerivedColumn(processUploadedFile.checkSheetList[i]);
            if (fileproc) {
                savedfile = processUploadedFile.checkSheetList[i];
                f = DriveApp.getFileById(processUploadedFile.checkSheetList[i]);
                returnData.fileName = f.getName();
                break;
            }
        }
    }


    if (savedfile !== false) {
        for (var i = 0; i < listToDel.length; i++) {
            if (listToDel[i] !== savedfile) {
                Drive.Files.remove(listToDel[i]);
            } else {
                f = DriveApp.getFileById(listToDel[i]);
                f.setTrashed(true);
            }
        }
        var resultofprocessing = processUploadedFile.walkSheetHeader(savedfile);
        if (resultofprocessing) {
            returnData.emailCol = processUploadedFile.emailCol;
            returnData.pointsCol = processUploadedFile.pointsCol;
            returnData.nameCol = processUploadedFile.nameCol;
            returnData.needsDomain = processUploadedFile.needsDomain;
            returnData.fileID = savedfile;
            if (returnData.needsDomain) {
                returnData.domainDefault = processUploadedFile.grabDomain();
            }
        }
        if (returnData.pointsCol === -1 || returnData.pointsCol.length === 0) {
            Drive.Files.remove(savedfile);
        }
    }

    return returnData;
}

var commonSettings = {};
commonSettings.fullName = ['name', 'student', 'student name', 'full name', 'full names', 'names', 'students', 'student names'];
commonSettings.firstName = ['first', 'first name', 'first names'];
commonSettings.lastName = ['last', 'last name', 'last names'];
commonSettings.email = ['email', 'e-mail', 'e-mail address', 'email address', 'e-mail addresses', 'email addresses'];
commonSettings.points = ['points', 'point', 'current points', 'final points'];
commonSettings.replace = ['replace', 'replacement', 'substitute'];
commonSettings.fullPoints = ['full points', 'possible', 'possible points'];
commonSettings.message = ['message', 'messages'];
commonSettings.username = ['username', 'usernames', 'login', 'logins'];
commonSettings.teststudents = ['test student', 'demo student', 'demo', 'test'];

commonSettings.determineDecSeperator = function () {
    var decSep = ".";
    try {
        var sep = parseFloat(3 / 2).toLocaleString().substring(1, 2);
        if (sep === ',') {
            decSep = sep;
        }
    } catch (e) {}
    return decSep;
};

commonSettings.determineFormulaSeperator = function () {
    if (commonSettings.determineDecSeperator() === ',') {
        return ";";
    } else {
        return ",";
    }
};

var processUploadedFile = {};
processUploadedFile.nameCol = -1;
processUploadedFile.emailCol = -1;
processUploadedFile.pointsCol = [];
processUploadedFile.needsDomain = false;
processUploadedFile.checkSheetList = [];
processUploadedFile.checkMaxColumns = 0;
processUploadedFile.emailDictionary = {
    names: [],
    emails: []
};


processUploadedFile.addDerivedColumn = function (file) {
    var spreadsheet = SpreadsheetApp.openById(file);
    var sheets = spreadsheet.getSheets();
    if (sheets.length > 0) {
        var firstSheet = sheets[0];
        var data = firstSheet.getDataRange().getValues();
        if (processUploadedFile.checkForNumbers(data)) {
            var lengthOfVector = data.length - 1;
            if (lengthOfVector > 0) {
                var lastCol = firstSheet.getLastColumn();
                firstSheet.getRange(1, (lastCol + 1)).setValue("Total - Calculated by Grade Nudge (Caution)");
                var newRange = firstSheet.getRange(2, (lastCol + 1), lengthOfVector);
                var formulas = [];
                var numberOfCols = data[1].length;
                var buildFEq = "SUM(ARRAYFORMULA(IFERROR(VALUE(R[0]C[-" + numberOfCols.toString() + "]:R[0]C[-1])" + commonSettings.determineFormulaSeperator() + "0)))";

                var formulaEq = [buildFEq, ];
                for (var i = 0; i < lengthOfVector; i++) {
                    formulas.push(formulaEq);
                }
                newRange.setFormulasR1C1(formulas);
                return true;
            }
        }
    }
    return false;
};

processUploadedFile.checkForNumbers = function (data) {
    for (var i = 1; i < data.length; i++) {
        for (var j = 0; j < data[i].length; j++) {
            if (isFloat(data[i][j])) {
                return true;
            }
        }
    }
    return false;
};

processUploadedFile.findEmailDictionary = function (file) {
    var spreadsheet = SpreadsheetApp.openById(file);
    var sheets = spreadsheet.getSheets();
    for (var j = 0; j < sheets.length; j++) {
        var data = sheets[j].getDataRange().getValues();
        var searchNameCol = -1;
        var searchEmailCol = -1;
        var searchUserCol = -1;
        var countNonCols = 0;
        var searchFirstCol = -1;
        var searchLastCol = -1;
        if (data.length > 0) {
            var firstRow = data[0];
            for (var i = 0; i < firstRow.length; i++) {
                if ((commonSettings.fullName.indexOf(firstRow[i].toString().toLowerCase().trim()) > -1) && searchNameCol === -1) {
                    searchNameCol = i;
                } else if ((commonSettings.firstName.indexOf(firstRow[i].toString().toLowerCase().trim()) > -1) && searchFirstCol === -1) {
                    searchFirstCol = i;
                } else if ((commonSettings.lastName.indexOf(firstRow[i].toString().toLowerCase().trim()) > -1) && searchLastCol === -1) {
                    searchLastCol = i;
                } else if ((commonSettings.email.indexOf(firstRow[i].toString().toLowerCase().trim()) > -1) && searchEmailCol === -1) {
                    searchEmailCol = i;
                } else if ((commonSettings.username.indexOf(firstRow[i].toString().toLowerCase().trim()) > -1 || firstRow[i].toString().toLowerCase().trim().substr(firstRow[i].toString().trim().length - 8) == 'login id') && searchUserCol === -1) {
                    searchUserCol = i;
                } else if (commonSettings.points.indexOf(firstRow[i].toString().toLowerCase().trim()) > -1 || firstRow[i].toString().toLowerCase().trim().substr(0, 5) == 'total' || firstRow[i].toString().toLowerCase().trim().substr(0, 12) == 'course total' || firstRow[i].toString().toLowerCase().trim().substr(0, 24) == 'blackboard points earned') {
                    return true;
                } else if (firstRow[i].toString().toLowerCase().trim().length > 0) {
                    countNonCols = countNonCols + 1;
                }
            }
            if (firstRow.length > 0 && firstRow[0].toString().trim() == '' && searchNameCol === -1) {
                searchNameCol = 0;
            }
            if (searchEmailCol === -1) {
                searchEmailCol = searchUserCol;
            }

            if ((searchNameCol > -1 || searchFirstCol > -1) && j === 0 && countNonCols > 0) {
                if (processUploadedFile.checkMaxColumns < countNonCols) {
                    processUploadedFile.checkMaxColumns = countNonCols;
                    processUploadedFile.checkSheetList.unshift(file);
                } else {
                    processUploadedFile.checkSheetList.push(file);
                }
            }
            if (searchNameCol > -1 && searchEmailCol > -1) {
                for (var i = 1; i < data.length; i++) {
                    if (data[i][searchNameCol].toString().trim().length > 0 && data[i][searchEmailCol].toString().trim().length > 0) {
                        processUploadedFile.emailDictionary.emails.push(data[i][searchEmailCol].toString().trim());
                        processUploadedFile.emailDictionary.names.push(data[i][searchNameCol].toString().trim());
                    }
                }
            } else if (searchEmailCol > -1 && searchFirstCol > -1 && searchLastCol > -1) {
                for (var i = 1; i < data.length; i++) {
                    var builtName = data[i][searchFirstCol].toString().trim() + " " + data[i][searchLastCol].toString().trim();
                    if (builtName.length > 0 && data[i][searchEmailCol].toString().trim().length > 0) {
                        processUploadedFile.emailDictionary.emails.push(data[i][searchEmailCol].toString().trim());
                        processUploadedFile.emailDictionary.names.push(builtName);
                    }
                }
            }
        }
    }
    return false;
};

processUploadedFile.walkSheetHeader = function (file) {
    var spreadsheet = SpreadsheetApp.openById(file);
    var tempusernameCol = -1;
    var sheets = spreadsheet.getSheets();
    var tempfirstnameCol = -1;
    var templastnameCol = -1;
    if (sheets.length > 0) {
        var firstSheet = sheets[0];
        var data = firstSheet.getDataRange()
            .getValues();
        if (data.length > 0) {
            var firstRow = data[0];
            for (var i = 0; i < firstRow.length; i++) {
                if ((commonSettings.fullName.indexOf(firstRow[i].toString().toLowerCase().trim()) > -1) && processUploadedFile.nameCol === -1) {
                    processUploadedFile.nameCol = i;
                } else if ((commonSettings.firstName.indexOf(firstRow[i].toString().toLowerCase().trim()) > -1) && tempfirstnameCol === -1) {
                    tempfirstnameCol = i;
                } else if ((commonSettings.lastName.indexOf(firstRow[i].toString().toLowerCase().trim()) > -1) && templastnameCol === -1) {
                    templastnameCol = i;
                } else if ((commonSettings.email.indexOf(firstRow[i].toString().toLowerCase().trim()) > -1) && processUploadedFile.emailCol === -1) {
                    processUploadedFile.emailCol = i;
                } else if (commonSettings.points.indexOf(firstRow[i].toString().toLowerCase().trim()) > -1 || firstRow[i].toString().toLowerCase().trim().substr(0, 5) == 'total' || firstRow[i].toString().toLowerCase().trim().substr(0, 12) == 'course total' || firstRow[i].toString().toLowerCase().trim().substr(0, 24) == 'blackboard points earned') {
                    processUploadedFile.pointsCol.push({
                        name: firstRow[i].toString().trim(),
                        id: i
                    });
                } else if ((commonSettings.username.indexOf(firstRow[i].toString().toLowerCase().trim()) > -1 || firstRow[i].toString().toLowerCase().trim().substr(firstRow[i].toString().trim().length - 8) == 'login id') && tempusernameCol === -1) {
                    tempusernameCol = i;
                }
            }
            if (firstRow.length > 0 && firstRow[0].toString().trim() == '' && processUploadedFile.nameCol === -1) {
                processUploadedFile.nameCol = 0;
            }
        }

        if (processUploadedFile.emailCol === -1) {
            processUploadedFile.emailCol = tempusernameCol;
        }
        if (processUploadedFile.emailCol === -1 && processUploadedFile.emailDictionary.names.length > 0 && (processUploadedFile.nameCol > -1 || (tempfirstnameCol > -1 && templastnameCol > -1))) {
            var lastCol = firstSheet.getLastColumn();
            processUploadedFile.emailCol = lastCol;
            firstSheet.getRange(1, (lastCol + 1)).setValue("E-Mail");
            for (var i = 1; i < data.length; i++) {
                if (processUploadedFile.nameCol > -1) {
                    var namedata = data[i][processUploadedFile.nameCol].toString().trim();
                } else {
                    var namedata = data[i][tempfirstnameCol].toString().trim() + " " + data[i][templastnameCol].toString().trim();
                }
                var foundid = processUploadedFile.emailDictionary.names.indexOf(namedata);
                if (foundid > -1 && namedata.length > 0) {
                    var foundemail = processUploadedFile.emailDictionary.emails[foundid].toString().trim();
                    if (foundemail.length > 0) {
                        firstSheet.getRange((i + 1), (lastCol + 1)).setValue(foundemail);
                    }
                }
            }
            var data = firstSheet.getDataRange().getValues();
        }
        if (processUploadedFile.nameCol === -1) {
            processUploadedFile.nameCol = tempfirstnameCol;
        }
        if (processUploadedFile.emailCol > -1 && processUploadedFile.pointsCol.length > 0 && processUploadedFile.nameCol > -1) {
            for (var i = 1; i < data.length; i++) {
                if (data[i][processUploadedFile.nameCol].trim().length > 0 && data[i][processUploadedFile.emailCol].toString().trim().length > 0) {
                    var emaildata = data[i][processUploadedFile.emailCol].toString().trim();
                    if (emaildata.indexOf("@") > -1) {
                        processUploadedFile.needsDomain = false;
                    } else {
                        processUploadedFile.needsDomain = true;
                    }
                    break;
                }
            }
            return true;
        } else if (processUploadedFile.pointsCol.length > 0 && processUploadedFile.nameCol > -1) {
            return true;
        } else {
            return false;
        }
    }
    return false;
};
processUploadedFile.grabDomain = function () {
    var domain = "";
    var emailofuser = Session.getActiveUser().getEmail().toString().trim();
    if (emailofuser.indexOf("@") > -1) {
        domain = emailofuser.substr(emailofuser.indexOf("@") + 1);
    }
    return domain;
};

function onlyUnique(value, index, self) {
    return self.indexOf(value) === index;
}

function isFileExistById(id) {
    try {
        return DriveApp.getFileById(id);
    } catch (err) {
        return false;
    }
}

function LoadDefaultGradeScale() {
    var returnData = {};
    returnData.gradescaleID = '';
    returnData.gradescaleName = '';
    returnData.gradescaleURL = '';
    var defualtSheetValues = [
        ["Grade", "threshold"],
        ["A", 0.93],
        ["A-", 0.9],
        ["B+", 0.87],
        ["B", 0.83],
        ["B-", 0.8],
        ["C+", 0.77],
        ["C", 0.73],
        ["C-", 0.7],
        ["D+", 0.67],
        ["D", 0.63],
        ["D-", 0.6],
        ["F", 0]
    ];
    var ssNew = SpreadsheetApp.create("Default Gradescale", defualtSheetValues.length, defualtSheetValues[0].length);
    var sheet = ssNew.getSheets()[0];
    var range = sheet.getRange(1, 1, defualtSheetValues.length, defualtSheetValues[0].length);
    range.setValues(defualtSheetValues);
    if (ssNew !== false) {
        returnData.gradescaleID = ssNew.getId();
        returnData.gradescaleName = ssNew.getName();
        returnData.gradescaleURL = ssNew.getUrl();
    }
    return returnData;
}

function LoadPastSheets() {
    var userProperties = PropertiesService.getUserProperties();
    var gradebookID = userProperties.getProperty('gradebookID');
    var gradescaleID = userProperties.getProperty('gradescaleID');
    var gbI = isFileExistById(gradebookID);
    var gsI = isFileExistById(gradescaleID);
    var returnData = {};
    returnData.gradebookID = '';
    returnData.gradebookName = '';
    returnData.gradebookURL = '';
    returnData.gradescaleID = '';
    returnData.gradescaleName = '';
    returnData.gradescaleURL = '';
    if (gbI !== false) {
        if (walkSheet.checkFileType(gbI.getId(), MimeType.GOOGLE_SHEETS)) {
            if (!gbI.isTrashed()) {
                returnData.gradebookID = gbI.getId();
                returnData.gradebookName = gbI.getName();
                returnData.gradebookURL = gbI.getUrl();
            }
        }
    }
    if (gsI !== false) {
        if (walkSheet.checkFileType(gsI.getId(), MimeType.GOOGLE_SHEETS)) {
            if (!gsI.isTrashed()) {
                returnData.gradescaleID = gsI.getId();
                returnData.gradescaleName = gsI.getName();
                returnData.gradescaleURL = gsI.getUrl();
            }
        }
    }
    returnData.YourEmailAddress = Session.getActiveUser().getEmail().toString().trim();
    return returnData;
}


function isInt(n) {
    return parseFloat(n) == parseInt(n, 10) && !isNaN(n);
}

function isFloat(n) {
    return !isNaN(n);
}

function validateEmail(email) {
    var re = /^(([^<>()[\]\\.,;:\s@\"]+(\.[^<>()[\]\\.,;:\s@\"]+)*)|(\".+\"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
    return re.test(email);
}

function getProgress() {
    var userProperties = PropertiesService.getUserProperties();
    return " (" + userProperties.getProperty('studentnum') + "/" + userProperties.getProperty('studenttotal') + ")";
}

var interfaceClass = {};
interfaceClass.templateFile = false;
interfaceClass.gradebook = false;
interfaceClass.gradescale = false;
interfaceClass.eMail = false;
interfaceClass.points = 0;
interfaceClass.classTotal = 0;
interfaceClass.studentnum = 0;
interfaceClass.studenttotal = 0;
interfaceClass.noemail = false;
interfaceClass.percentage = false;
interfaceClass.random = false;
interfaceClass.quota = 1500;
interfaceClass.SavedstudentNameCol = -1;
interfaceClass.SavedstudentEMailCol = -1;
interfaceClass.Savedemailappend = "";
interfaceClass.Savedpointscollist = -1;

interfaceClass.createPercent = function (points, total) {
    if (interfaceClass.percentage === true) {
        var cper = 100 * points / total;
        var pars = cper.toFixed(0);
        return ' (' + pars + '%)';
    } else {
        return '';
    }
};


interfaceClass.cleanUpFile = function (rResult, rangecheck) {
    if (interfaceClass.SavedstudentNameCol > -1 && interfaceClass.Savedpointscollist > -1) {
        f = isFileExistById(interfaceClass.gradebook);
        if (f !== false) {
            if (interfaceClass.noemail && f.isTrashed()) {
                f.setTrashed(false);
            } else if (f.isTrashed() && rangecheck === false && rResult === true) {
                Drive.Files.remove(f.getId());
            }
        }
    }
};

interfaceClass.submitWindow = function (data) {
    interfaceClass.templateFile = false;
    interfaceClass.gradebook = false;
    interfaceClass.gradescale = false;
    interfaceClass.studentnum = 0;
    interfaceClass.studenttotal = 0;
    interfaceClass.noemail = false;
    interfaceClass.percentage = false;
    interfaceClass.random = false;
    interfaceClass.quota = MailApp.getRemainingDailyQuota();
    var assignmentPoints = data.points.trim();
    var numberOfPoints = data.tpoints.trim();
    var emailbody = data.emailbody.trim();
    var subjectline = data.subjectline.trim();

    var incomplete = (data.incomplete === "on") ? true : false;
    var gradeup = (data.gradeup === "on") ? true : false;
    var gradedown = (data.gradeup === "on") ? true : false;
    var randomize = false;

    var sharedrive = (data.delivery === 'share') ? true : false;
    var noemail = (data.delivery === 'noemail') ? true : false;
    var percentage = (data.percentage === "on") ? true : false;
    interfaceClass.noemail = noemail;
    interfaceClass.percentage = percentage;
    interfaceClass.random = randomize;

    if (isInt(data.SavedstudentNameCol)) {
        interfaceClass.SavedstudentNameCol = parseInt(data.SavedstudentNameCol, 10)
        if (interfaceClass.SavedstudentNameCol > -1) {
            if (isInt(data.SavedstudentEMailCol)) {
                interfaceClass.SavedstudentEMailCol = parseInt(data.SavedstudentEMailCol, 10);
            }
            if (isInt(data.Savedpointscollist)) {
                interfaceClass.Savedpointscollist = parseInt(data.Savedpointscollist, 10);
            }
            interfaceClass.Savedemailappend = data.Savedemailappend;
        }
    }


    var start = data.lStart;
    var end = data.lEnd;
    
    if (data.assignmentName){
        if (data.assignmentName.trim().length > 0){
            makeGrades.assignmentName = data.assignmentName.trim();
        }
    }
    
    var rangecheck = false;
    if (isInt(start) && isInt(end)) {
        var start = parseInt(start, 10);
        var end = parseInt(end, 10);
        if (start > 0 && end >= start) {
            var rangecheck = true;
        }
    }

    var notemp = true;

    if (data.template.trim().length > 0) {
        var notemp = false;
        var fileCheck = data.template.trim();
        var check = walkSheet.checkFileType(fileCheck, MimeType.GOOGLE_DOCS);
        if (check === true) {
            interfaceClass.templateFile = fileCheck;
        }
    }


    if (data.gradebook.trim().length > 0) {
        var fileCheck = data.gradebook.trim();
        var check = walkSheet.checkFileType(fileCheck, MimeType.GOOGLE_SHEETS);
        if (check === true) {
            interfaceClass.gradebook = fileCheck;
        }
    }

    if (data.gradescale.trim().length > 0) {
        var fileCheck = data.gradescale.trim();
        var check = walkSheet.checkFileType(fileCheck, MimeType.GOOGLE_SHEETS);
        if (check === true) {
            interfaceClass.gradescale = fileCheck;
        }
    }
    var go = false;

    if (isFloat(assignmentPoints) && isFloat(numberOfPoints)) {
        if (parseFloat(assignmentPoints) > 0 && parseFloat(numberOfPoints) > 0) {
            var go = true;
        }
    }

    if (go === true) {
        if (interfaceClass.gradebook !== false && interfaceClass.gradescale !== false) {
            var ggh = walkSheet.getGradeHeader(interfaceClass.gradescale);
            var gh = walkSheet.getHeader(interfaceClass.gradebook);
            var userProperties = PropertiesService.getUserProperties();

            if (ggh === false) {
                return "The gradescale does not follow a format that nudge understands.";
            }

            if (gh === false) {
                return "The gradebook does not follow a format that nudge understands.";
            }
            userProperties.setProperty('gradebookID', interfaceClass.gradebook);
            userProperties.setProperty('gradescaleID', interfaceClass.gradescale);
            if (interfaceClass.templateFile === false && notemp === true) {
                rlc = makeGrades.LoopOverClass(false, gh, ggh, parseFloat(assignmentPoints), parseFloat(numberOfPoints), sharedrive, incomplete, gradeup, gradedown, randomize, subjectline, emailbody, rangecheck, start, end, false);
                interfaceClass.cleanUpFile(rlc, rangecheck);
                return rlc;
            } else if (interfaceClass.templateFile !== false) {
                rlc = makeGrades.LoopOverClass(true, gh, ggh, parseFloat(assignmentPoints), parseFloat(numberOfPoints), sharedrive, incomplete, gradeup, gradedown, randomize, subjectline, emailbody, rangecheck, start, end, false);
                interfaceClass.cleanUpFile(rlc, rangecheck);
                return rlc;
            } else {
                return "The selected template file could not be used.";
            }
        } else {
            return "You must specify a gradebook and gradescale spreadsheet";
        }
    } else {
        return "I'm sorry, something is wrong with the assignment points.";
    }
    return "An unknown error occured";
};


walkSheet = {};

walkSheet.names = [];
walkSheet.emails = [];
walkSheet.points = [];


makeGrades = {};
makeGrades.assignmentName = 'this assignment';

makeGrades.randomize = function () {
    var rv = Math.random();
    if (rv >= 0.5) {
        return 1;
    } else {
        return 0;
    }
};

makeGrades.saveTreatment = function (gradebook, i, treatment) {
    if (walkSheet.treatmentCol > -1) {
        gradebook.getRange((i + 1), (walkSheet.treatmentCol + 1)).setValue(treatment);
    }
};

makeGrades.saveMessage = function (gradebook, i, message) {
    if (walkSheet.messageCol > -1) {
        gradebook.getRange((i + 1), (walkSheet.messageCol + 1)).setValue(message);
    }
};

makeGrades.LoopOverClass = function (maketemplate, gradebook, gradescale, assignmentPoints, ClassNumberOfPoints, sharedrive, incomplete, gradeup, gradedown, randomize, subjectline, emailbody, rangecheck, start, end, noscore) {
    if (maketemplate !== false) {
        var foldername = DriveApp.getFileById(interfaceClass.templateFile).getName();
        var folder = newFolderClass.createFolder(foldername);
        var dupfile = DriveApp.getFileById(interfaceClass.templateFile);
    }


    var initVal = 0;
    var endval = walkSheet.points.length;
    if (rangecheck === true) {
        if (start > 0 && start <= endval) {
            var initVal = start - 1;
        }
        if (end > 0 && end <= endval) {
            if (end >= (initVal + 1)) {
                var endval = end;
            } else {
                var endval = initVal + 1;
            }
        }
    }


    var userProperties = PropertiesService.getUserProperties();
    interfaceClass.studenttotal = endval;
    interfaceClass.studentnum = initVal + 1;
    userProperties.setProperty('studenttotal', interfaceClass.studenttotal.toString());
    userProperties.setProperty('studentnum', interfaceClass.studentnum.toString());
    numstudentstosend = endval - initVal;
    if (numstudentstosend > interfaceClass.quota && interfaceClass.noemail !== true) {
        return 'You are attempting to send ' + Math.round(numstudentstosend) + ' emails while your daily remaining Gmail sending quota is currently at ' + Math.round(interfaceClass.quota) + '. You can set a range of students you would like to send emails to under the advanced settings section.';
    }
    for (var i = initVal; i < endval; i++) {
        var numberOfPoints = ClassNumberOfPoints;
        interfaceClass.studentnum = i + 1;

        if (interfaceClass.studentnum % 10 === 0) {
            userProperties.setProperty('studentnum', interfaceClass.studentnum.toString());
        }

        var studentemail = walkSheet.emails[i];
        if (interfaceClass.Savedemailappend.length > 0 && studentemail.length > 0 && studentemail.indexOf("@") === -1) {
            studentemail += "@" + interfaceClass.Savedemailappend;
        }
        var studentname = walkSheet.names[i];
        var studentpoints = walkSheet.points[i];
        if (walkSheet.replacementCol > -1) {
            var studentreplacement = walkSheet.replacement[i];
        } else {
            var studentreplacement = 0;
        }
        if (walkSheet.possiblePointsCol > -1) {
            if (walkSheet.possiblePoints[i] !== false && isFloat(walkSheet.possiblePoints[i])) {
                numberOfPoints = walkSheet.possiblePoints[i];
            }
        }
        if (noscore === true) {
            var currentGrade = false;
        } else {
            var currentGrade = walkSheet.getGrade(studentpoints, numberOfPoints);
        }
        var lgrade = "";
        var ugrade = "";
        var igrade = "";

        if (interfaceClass.random === true) {
            var treatment = makeGrades.randomize();
        } else {
            var treatment = 1;
        }
        makeGrades.saveTreatment(gradebook, walkSheet.rowpos[i], treatment);
        if (currentGrade.grade !== false && currentGrade.id !== false && treatment === 1 && assignmentPoints > 0 && numberOfPoints > 0) {

            if (gradedown) {
                var lowerthreshold = walkSheet.threshold[currentGrade.id];
                if (walkSheet.replacementCol > -1) {
                    var lowergrade = false;
                } else {
                    var lowergrade = walkSheet.getGradeDown(currentGrade.id);
                }
                if (lowergrade !== false) {
                    var lgrade = lowergrade.grade;
                    var godownScore = walkSheet.calcGrade(studentpoints, assignmentPoints, numberOfPoints, lowerthreshold);
                } else {
                    var godownScore = false;
                }
            } else {
                var godownScore = false;
            }

            if (gradeup) {
                var uppergrade = walkSheet.getGradeUp(currentGrade.id);
                if (uppergrade !== false) {
                    var upperthreshold = uppergrade.threshold;
                    var ugrade = uppergrade.grade;
                    if (walkSheet.replacementCol > -1) {
                        var goupScore = walkSheet.calcGrade((studentpoints - studentreplacement), assignmentPoints, (numberOfPoints - assignmentPoints), upperthreshold);
                    } else {
                        var goupScore = walkSheet.calcGrade(studentpoints, assignmentPoints, numberOfPoints, upperthreshold);
                    }

                } else {
                    var goupScore = false;
                }
            } else {
                var goupScore = false;
            }

            if (incomplete) {
                if (walkSheet.replacementCol > -1) {
                    incompleteGrade = false;
                } else {
                    incompleteGrade = walkSheet.getGrade(studentpoints, (numberOfPoints + assignmentPoints));
                }
                if (incompleteGrade !== false) {
                    if (incompleteGrade.id !== false) {
                        if (incompleteGrade.id !== currentGrade.id) {
                            var igrade = incompleteGrade.grade;
                        }
                    }
                }
            }

            var message = makeGrades.writeMessage(currentGrade.grade, assignmentPoints, goupScore, godownScore, incomplete, ugrade, lgrade, igrade);

        } else {
            var message = "";
        }

        var emailcontent = "";
        if (studentname.length > 0 && (emailbody.length > 0 || message.length > 0)) {
            emailcontent += "Hi " + studentname + ",\r\n\r\n";
        }
        emailcontent += emailbody;

        if (message.length > 0 && emailbody.length > 0) {
            emailcontent += "\r\n\r\n";
        }
        if (interfaceClass.noemail === true) {
            makeGrades.saveMessage(gradebook, walkSheet.rowpos[i], (emailcontent + message));
        }
        if (maketemplate !== false) {
            var des = dupfile.makeCopy(dupfile.getName() + " - " + studentname + " - " + studentemail, folder);
            var dupid = des.getId();
            if (message.length > 0) {
                var doc = DocumentApp.openById(dupid);
                var body = doc.getBody();
                body.appendHorizontalRule();
                if (studentname.length > 0 && message.length > 0) {
                    body.appendParagraph("Hi " + studentname + ",");
                }
                body.appendParagraph(message);
                doc.saveAndClose();
            }

            if (sharedrive && validateEmail(studentemail)) {
                des.addEditor(studentemail);
                newEmailClass.sendEmail(subjectline, studentemail, (emailcontent + message));
            } else {
                newEmailClass.sendEmailFile(des, subjectline, studentemail, (emailcontent + message));
            }

        } else {
            if (interfaceClass.noemail === false) {
                newEmailClass.sendEmail(subjectline, studentemail, (emailcontent + message));
            }
        }
    }
    userProperties.setProperty('studenttotal', '');
    userProperties.setProperty('studentnum', '');
    return true;
};




makeGrades.writeMessage = function (grade, points, goupScore, godownScore, incomplete, ugrade, lgrade, igrade) {
    var message = "";
    var assignmentName = makeGrades.assignmentName;
    if (grade.length > 0) {
        message += "As of now, you have a(n) " + grade + " in the class.  " + capitalizeFirstLetter(assignmentName) + " is worth " + points.toFixed(2) + " points.";
    }
    if (goupScore !== false) {
        message += "  If you get more than " + goupScore.toFixed(2);
        message += interfaceClass.createPercent(goupScore, points);
        message += " on " + assignmentName + ", your class grade will increase to a(n) " + ugrade + ".";
    }
    if (godownScore !== false) {
        message += "  If you get less than " + godownScore.toFixed(2);
        message += interfaceClass.createPercent(godownScore, points);
        message += " on " + assignmentName + ", your grade will drop at least one grade.";
    }
    if (incomplete !== false && igrade.length > 0) {
        message += "  Not doing " + assignmentName + " will result in a(n) " + igrade + ".";
    }
    return message;
};

var newFolderClass = {};

newFolderClass.createFolder = function (folderName) {
    var folder = DriveApp.createFolder(folderName);
    return folder;
};


newFolderClass.shareFolder = function (folder, emailList) {
    folder.addEditors(emailList);
};

walkSheet = {};
walkSheet.emailCol = -1;
walkSheet.pointsCol = -1;
walkSheet.nameCol = -1;
walkSheet.replacementCol = -1;
walkSheet.treatmentCol = -1;
walkSheet.messageCol = -1;
walkSheet.gradeCol = -1;
walkSheet.thresholdCol = -1;
walkSheet.possiblePointsCol = -1;
walkSheet.names = [];
walkSheet.emails = [];
walkSheet.points = [];
walkSheet.grades = [];
walkSheet.rowpos = [];
walkSheet.threshold = [];
walkSheet.replacement = [];
walkSheet.possiblePoints = [];



walkSheet.checkFileType = function (myfile, reqType) {
    var file = DriveApp.getFileById(myfile);
    if (file.getMimeType() == reqType) {
        return true;
    } else {
        return false;
    }
};


walkSheet.sorter = function (A, B) {

    var all = [];

    for (var i = 0; i < B.length; i++) {
        all.push({
            'A': A[i],
            'B': B[i]
        });
    }

    all.sort(function (a, b) {
        return a.A - b.A;
    });

    A = [];
    B = [];

    for (var i = 0; i < all.length; i++) {
        A.push(all[i].A);
        B.push(all[i].B);
    }

    return {
        sortObj: A,
        sortedObj: B
    };
};



walkSheet.getGradeHeader = function (file) {
    var spreadsheet = SpreadsheetApp.openById(file);
    var sheets = spreadsheet.getSheets();
    if (sheets.length > 0) {
        var firstSheet = sheets[0];
        var data = firstSheet.getDataRange()
            .getValues();
        if (data.length > 0) {
            var firstRow = data[0];
            for (var i = 0; i < firstRow.length; i++) {
                if (firstRow[i].toString().toLowerCase().trim() == 'grade') {
                    walkSheet.gradeCol = i;
                } else if (firstRow[i].toString().toLowerCase().trim() == 'threshold') {
                    walkSheet.thresholdCol = i;
                }
            }
        }

        if (walkSheet.gradeCol > -1 && walkSheet.thresholdCol > -1) {
            for (var i = 1; i < data.length; i++) {
                walkSheet.grades.push(data[i][walkSheet.gradeCol].toString().trim());
                if (isNaN(parseFloat(data[i][walkSheet.thresholdCol]))) {
                    walkSheet.threshold.push(0.0);
                } else {
                    walkSheet.threshold.push(parseFloat(data[i][walkSheet.thresholdCol]));
                }
            }
        } else {
            return false;
        }

        var sortedList = walkSheet.sorter(walkSheet.threshold, walkSheet.grades);
        walkSheet.threshold = sortedList.sortObj;
        walkSheet.grades = sortedList.sortedObj;
        return firstSheet;
    }
    return false;
};


walkSheet.getGrade = function (score, pointsInClass) {
    var grade = false;
    var myid = false;
    if (pointsInClass > 0) {
        var avscore = score / pointsInClass;
        for (var i = 0; i < walkSheet.threshold.length; i++) {
            if (walkSheet.threshold[i] <= avscore) {
                var grade = walkSheet.grades[i];
                var myid = i;
            }
        }
    }
    return {
        grade: grade,
        id: myid
    };
};

walkSheet.getGradeUp = function (currentGradeId) {
    var nextgrade = currentGradeId + 1;
    if (nextgrade >= 0 && nextgrade < walkSheet.threshold.length) {
        return {
            threshold: walkSheet.threshold[nextgrade],
            grade: walkSheet.grades[nextgrade]
        };
    } else {
        return false;
    }
};



walkSheet.getGradeDown = function (currentGradeId) {
    var nextgrade = currentGradeId - 1;
    if (nextgrade >= 0 && nextgrade < walkSheet.threshold.length) {
        return {
            threshold: walkSheet.threshold[nextgrade],
            grade: walkSheet.grades[nextgrade]
        };
    } else {
        return false;
    }
};

walkSheet.calcGrade = function (points, apoints, cpoints, threshold) {
    var score = threshold * (cpoints + apoints) - points;
    if (score >= 0 && score <= apoints) {
        return score;
    } else {
        return false;
    }
};


walkSheet.getHeader = function (file) {
    var spreadsheet = SpreadsheetApp.openById(file);
    var sheets = spreadsheet.getSheets();
    if (sheets.length > 0) {
        var firstSheet = sheets[0];
        var data = firstSheet.getDataRange().getValues();
        if (data.length > 0) {
            var firstRow = data[0];

            if (interfaceClass.SavedstudentNameCol > -1 && interfaceClass.SavedstudentNameCol < firstRow.length) {
                walkSheet.nameCol = interfaceClass.SavedstudentNameCol;
            }
            if (interfaceClass.SavedstudentEMailCol > -1 && interfaceClass.SavedstudentEMailCol < firstRow.length) {
                walkSheet.emailCol = interfaceClass.SavedstudentEMailCol;
            }
            if (interfaceClass.Savedpointscollist > -1 && interfaceClass.Savedpointscollist < firstRow.length) {
                walkSheet.pointsCol = interfaceClass.Savedpointscollist;
            }

            for (var i = 0; i < firstRow.length; i++) {
                if ((commonSettings.fullName.indexOf(firstRow[i].toString().toLowerCase().trim()) > -1 || commonSettings.firstName.indexOf(firstRow[i].toString().toLowerCase().trim()) > -1) && walkSheet.nameCol === -1) {
                    walkSheet.nameCol = i;
                } else if ((commonSettings.email.indexOf(firstRow[i].toString().toLowerCase().trim()) > -1) && walkSheet.emailCol === -1) {
                    walkSheet.emailCol = i;
                } else if ((commonSettings.points.indexOf(firstRow[i].toString().toLowerCase().trim()) > -1 || firstRow[i].toString().toLowerCase().trim().substr(0, 5) == 'total') && walkSheet.pointsCol === -1) {
                    walkSheet.pointsCol = i;
                } else if ((commonSettings.message.indexOf(firstRow[i].toString().toLowerCase().trim()) > -1) && walkSheet.messageCol === -1) {
                    walkSheet.messageCol = i;
                } else if ((commonSettings.replace.indexOf(firstRow[i].toString().toLowerCase().trim()) > -1) && walkSheet.replacementCol === -1) {
                    walkSheet.replacementCol = i;
                } else if (commonSettings.fullPoints.indexOf(firstRow[i].toString().toLowerCase().trim()) > -1 && walkSheet.possiblePointsCol === -1) {
                    walkSheet.possiblePointsCol = i;
                }
            }
        }

        if (interfaceClass.noemail === true && walkSheet.messageCol < 0) {
            var lastCol = firstSheet.getLastColumn();
            firstSheet.getRange(1, (lastCol + 1)).setValue("message");
            walkSheet.messageCol = lastCol;
        }

        if (walkSheet.nameCol > -1 && (walkSheet.emailCol > -1 || interfaceClass.noemail === true) && walkSheet.pointsCol > -1) {
            for (var i = 1; i < data.length; i++) {
                var nameNoNumber = data[i][walkSheet.nameCol].toString().replace(/[0-9#]/g, '').trim();
                if (commonSettings.teststudents.indexOf(nameNoNumber.toLowerCase()) === -1 || (data.length - i) > 1) {
                    if (interfaceClass.noemail === true) {
                        if (nameNoNumber.length > 0) {
                            walkSheet.rowpos.push(i);
                            walkSheet.names.push(nameNoNumber);
                            walkSheet.emails.push("");
                            if (isNaN(parseFloat(data[i][walkSheet.pointsCol]))) {
                                walkSheet.points.push(0.0);
                            } else {
                                walkSheet.points.push(parseFloat(data[i][walkSheet.pointsCol]));
                            }
                            if (walkSheet.replacementCol > -1) {
                                if (isNaN(parseFloat(data[i][walkSheet.replacementCol]))) {
                                    walkSheet.replacement.push(0.0);
                                } else {
                                    walkSheet.replacement.push(parseFloat(data[i][walkSheet.replacementCol]));
                                }
                            }
                            if (walkSheet.possiblePointsCol > -1) {
                                if (isNaN(parseFloat(data[i][walkSheet.possiblePointsCol]))) {
                                    walkSheet.possiblePoints.push(false);
                                } else {
                                    walkSheet.possiblePoints.push(parseFloat(data[i][walkSheet.possiblePointsCol]));
                                }
                            }
                        }
                    } else {
                        if (nameNoNumber.length > 0 && data[i][walkSheet.emailCol].toString().trim().length > 0) {
                            walkSheet.rowpos.push(i);
                            walkSheet.names.push(nameNoNumber);
                            walkSheet.emails.push(data[i][walkSheet.emailCol].toString().trim());
                            if (isNaN(parseFloat(data[i][walkSheet.pointsCol]))) {
                                walkSheet.points.push(0.0);
                            } else {
                                walkSheet.points.push(parseFloat(data[i][walkSheet.pointsCol]));
                            }
                            if (walkSheet.replacementCol > -1) {
                                if (isNaN(parseFloat(data[i][walkSheet.replacementCol]))) {
                                    walkSheet.replacement.push(0.0);
                                } else {
                                    walkSheet.replacement.push(parseFloat(data[i][walkSheet.replacementCol]));
                                }
                            }
                            if (walkSheet.possiblePointsCol > -1) {
                                if (isNaN(parseFloat(data[i][walkSheet.possiblePointsCol]))) {
                                    walkSheet.possiblePoints.push(false);
                                } else {
                                    walkSheet.possiblePoints.push(parseFloat(data[i][walkSheet.possiblePointsCol]));
                                }
                            }
                        }
                    }
                }
            }
        } else {
            return false;
        }

        return firstSheet;
    }
    return false;
};

walkSheet.getValue = function (sheetObj, row, col) {
    var firstSheet = sheetObj;
    var data = firstSheet.getDataRange()
        .getValues();
    var r = parseInt(row);
    var c = parseInt(col);
    return data[r][c];
};


newEmailClass = {};

newEmailClass.sendEmailFile = function (file, subjectLine, email, emailContent) {
    if (validateEmail(email)) {
        subjectLine = subjectLine || file.getName();

        if (emailContent.length == 0) {
            emailContent = "Your assignment is attached to this e-mail."
        }

        MailApp.sendEmail(email, subjectLine, emailContent, {
            attachments: [file.getAs(MimeType.PDF)],
            name: file.getName()
        });
    }
};

newEmailClass.sendEmail = function (subjectLine, email, emailContent) {
    if (emailContent.length > 0 && validateEmail(email)) {
        subjectLine = subjectLine || "(No subject)";
        MailApp.sendEmail(email, subjectLine, emailContent);
    }
};