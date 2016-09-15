/* Common app functionality */

var app = (function () {
    "use strict";

    var app = {};
    var host = 'https://uprincecoreprodapi.azurewebsites.net';
    // Common initialization function (to be called from each page)
    app.initialize = function () {
        $('body').append(
            '<div id="notification-message">' +
                '<div class="padding">' +
                    '<div id="notification-message-close"></div>' +
                    '<div id="notification-message-header"></div>' +
                    '<div id="notification-message-body"></div>' +
                '</div>' +
            '</div>');

        $('#notification-message-close').click(function () {
            $('#notification-message').hide();
        });


        // After initialization, expose a common notification function
        app.showNotification = function (header, text) {
            $('#notification-message-header').text(header);
            $('#notification-message-body').text(text);
            $('#notification-message').slideDown('fast');
        };

        function onBindingNotFound() {
            showMessage("The binding object was not found. " +
            "Please return to previous step to create the binding");
        }
        $(document).on("click", "#Refresh", function () {
            //deleteTable('ProductDescription');
            localStorage.setItem("dailyLog", "true");
            //deleteTable('DailyLog');
            dailyLogGET();

        });

        $(document).on("click", "#Publish", function () {
            //publishRiskRegister();
            //publishProductDescription();
            publishDailyLog();
        });

        $(document).on("click", "#createSheet", function () {
            Excel.run(function (ctx) {
                ctx.workbook.worksheets.add("Values");
                return ctx.sync().then(function () {
                    //app.showNotification("Success! Sheet created");
                    //riskRegisterGET();
                });
            }).catch(function (error) {
                //app.showNotification(error);
            });
            Excel.run(function (ctx) {
                ctx.workbook.worksheets.add("DailyLog");
                return ctx.sync().then(function () {
                    //app.showNotification("Success! Sheet created");
                    //riskRegisterGET();
                });
            }).catch(function (error) {
                //app.showNotification(error);
            });
            Excel.run(function (ctx) {
                ctx.workbook.worksheets.add("RiskRegister");
                return ctx.sync().then(function () {
                    //app.showNotification("Success! Sheet created");
                    //riskRegisterGET();
                });
            }).catch(function (error) {
                //app.showNotification(error);
            });
            Excel.run(function (ctx) {
                ctx.workbook.worksheets.add("ProductDescription");
                return ctx.sync().then(function () {
                    //app.showNotification("Success! Sheet created");
                    //riskRegisterGET();
                });
            }).catch(function (error) {
                //app.showNotification(error);
            });

        });
    };
    function isNull(param) {
        if (param == null) return '';
        else return param;
    }

    function isZero(param) {
        if (param == 0) return null
        else return param;
    }

    function isTypeSeverity(severity) {
        if (severity == "TeamManager") return "Team Manager";
        else if (severity == "ProjectManager") return "Project Manager";
        else if (severity == "ProjectBoard") return "Project Board";
        else if (severity == "CooperateProgramManagement") return "Corporate / Program Management";
    }

    //if date is given as a regular date yyyy-mm-ddT..
    function formatDate(date) {
        if (date == null || date == "") { return '' }
            //app.showNotification();
        else {
            date = date.substring(0, 10);
            //app.showNotification(date);
            var year = date.substring(0, 4);
            var month = date.substring(5, 7);
            if (month.charAt(0) == "0") month = month.substring(1, 2);
            var day = date.substring(8, 10)
            if (day.charAt(0) == "0") day = day.substring(1, 2);
            //app.showNotification(year + "-" + month + "-" + day)
            return year + "-" + month + "-" + day;
        };
    }

    //if date is given in second since 01-01-1970
    function formatDate2(dateS) {
        if (dateS == null) return '';
        else {
            var mSeconds = dateS * 1000;
            var date = new Date(mSeconds);
            var day = date.getDate();
            var month = date.getMonth() + 1;
            var year = date.getFullYear();
            //app.showNotification(year + "-" + month + "-" + day);
            return year + "-" + month + "-" + day;
        };
    };

    function formatDate3(date2) {
        if (date2 == null || date2 == 0) return null;
        else {
            var date = new Date((date2 - (25567 + 2)) * 86400 * 1000);

            return formatDate(date.toJSON());
        }
    };

    //if date is asked in form yyyy-mm-dd
    function convertDate(days) {
        if (days == "" || days == null) { return null }
        else {
            var dateDays = days - 25569;
            var dateMS = dateDays * 24 * 60 * 60 * 1000;
            var date = new Date(dateMS);
            var day = date.getDate();
            var month = date.getMonth() + 1;
            var year = date.getFullYear();
            return (year + "-" + month + "-" + day + "T00:00:00.000");
        }
    }

    //if date is asked in seconds since 01-01-1970
    function convertDate2(days) {
        if (days == "" || days == null) { return null }
        else {
            var dateDays = days - 25569;
            var dateS = dateDays * 24 * 60 * 60;
            return dateS
        }
    }

    function activateWorksheet(name) {
        Excel.run(function (ctx) {
            var wSheetName = name;
            var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
            worksheet.activate();
            return ctx.sync();
        })
    .catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
    };

    function deleteTable(name) {
        Excel.run(function (ctx) {
            var tableName = name;
            var table = ctx.workbook.tables.getItem(tableName);
            var tableDataRange = table.getDataBodyRange();
            var values = "";
            tableDataRange.values = values;

            return ctx.sync().then(function () {
                console.log("delete is ok");
            });
        }).catch(function (error) {
            console.log("Error: " + error);
        });
    }

    function deleteValues() {
        Excel.run(function (ctx) {
            var values = "";
            var Project = ctx.workbook.tables.getItem("Projects");
            var tableDataRangeProject = Project.getDataBodyRange();
            tableDataRangeProject.values = values;
            var atContext = ctx.workbook.tables.getItem("atContext");
            var tableDataRangeatContext = atContext.getDataBodyRange();
            tableDataRangeatContext.values = values;
            var status = ctx.workbook.tables.getItem("Status");
            var tableDataRangestatus = status.getDataBodyRange();
            tableDataRangestatus.values = values;
            var responsible = ctx.workbook.tables.getItem("Responsible");
            var tableDataRangeresponsible = responsible.getDataBodyRange();
            tableDataRangeresponsible.values = values;
            var type = ctx.workbook.tables.getItem("Type");
            var tableDataRangetype = type.getDataBodyRange();
            tableDataRangetype.values = values;
            var time = ctx.workbook.tables.getItem("Time");
            var tableDataRangetime = time.getDataBodyRange();
            tableDataRangetime.values = values;
            var energy = ctx.workbook.tables.getItem("Energy");
            var tableDataRangeenergy = energy.getDataBodyRange();
            tableDataRangeenergy.values = values;
            return ctx.sync().then(function () {
                console.log("delete is ok");
            });
        }).catch(function (error) {
            console.log("Error: " + error);
        });
    };
    //Daily Log
    function dailyLogGET() {
        localStorage.setItem("dailyLog", "false")
        deleteTable("DailyLog");
        var projectId = localStorage.getItem('projectId');
        var userEmail = localStorage.getItem('email');
        var urlProject = host + "/api/DailyLog/GetDailyLog";
        var dataEmail =
            {
                "projectId": "",
                "project": null,
                "identifier": "",
                "title": "",
                "activity": "",
                "responsibleStatusString": "",
                "status": {
                    "All": true,
                    "New": false,
                    "Waiting": false,
                    "Completed": false
                },
                "activityType": {
                    "All": true,
                    "Problem": false,
                    "Action": false,
                    "Event": false,
                    "Comment": false,
                    "Decision": false,
                    "Reference": false
                },
                "priority": {
                    "All": true,
                    "High": false,
                    "Medium": false,
                    "Low": false
                },
                "targetDate": {
                    "All": true,
                    "-7": false,
                    "-1": false,
                    "0": false,
                    "+1": false,
                    "+7": false
                },
                "responsibleStatus": {
                    "All": true,
                    "Inbox": false,
                    "Next": false,
                    "Waiting": false,
                    "Schedule": false,
                    "Someday": false,
                    "Done": false
                },
                "energy": {
                    "All": true,
                    "Mild": false,
                    "Reasonable": false,
                    "Demanding": false,
                    "Very Demanding": false,
                    "Extreme": false
                },
                "responsible": "",
                "requester": "",
                "coreUserEmail": userEmail,
                "atContext": "",
                "startDate": "",
                "orderField": "",
                "sortOrder": ""
            };
        $.ajax({
            type: "POST",
            url: urlProject,
            dataType: "json",
            contentType: "application/json; charset=utf-8",
            data: JSON.stringify(dataEmail),
        })
            .done(function (str) {
                //app.showNotification(str.dailyLogListViewModel.length);
                getDailyLog(str.dailyLogListViewModel[1].identifier);
                var DlId = [str.dailyLogListViewModel.length];
                if (str.dailyLogListViewModel.length > 0) {
                    var matrix = [str.dailyLogListViewModel.length];
                    for (var i = 0; i < str.dailyLogListViewModel.length; i++) {
                        matrix[i] = [11];
                        DlId[i] = str.dailyLogListViewModel[i].id;
                        matrix[i][0] = isNull(str.dailyLogListViewModel[i].project);
                        var activity = isNull(str.dailyLogListViewModel[i].activity);
                        matrix[i][1] = activity;
                        var identifier = isNull(str.dailyLogListViewModel[i].identifier);
                        matrix[i][10] = str.dailyLogListViewModel[i].id;
                        matrix[i][3] = isNull(str.dailyLogListViewModel[i].atContext);
                        //matrix[i][4] = str.dailyLogListViewModel[i].targetDate;
                        matrix[i][4] = formatDate(str.dailyLogListViewModel[i].targetDate);
                        matrix[i][5] = isNull(str.dailyLogListViewModel[i].responsibleStatus);
                        matrix[i][6] = isNull(str.dailyLogListViewModel[i].responsible);
                        matrix[i][7] = isNull(str.dailyLogListViewModel[i].activityType);
                        matrix[i][8] = isNull(str.dailyLogListViewModel[i].time);
                        matrix[i][9] = isNull(str.dailyLogListViewModel[i].energy);
                        matrix[i][2] = isNull(str.dailyLogListViewModel[i].identifier);
                        //matrix[i][10] = 'hello';
                        storeLineDaily(matrix[i]);
                        //GTD > excel > Hyperlink 
                        matrix[i][2] = '=HYPERLINK("https://start.uprince.com/Home/Index/gdt/daily-log/description/"&[@Column1],[@Column1])'

                        //app.showNotification(matrix[i].toString())
                        //matrix[i][6] = isNull(str[i].Version);
                        //localStorage.setItem("ParentId" + str[i].Id, str[i].ParentId);
                    }
                } else {
                    var matrix = [["", "", "", "", "", "", "", "", "", "", ""]]
                };

                localStorage.setItem("DlId", DlId);

                Excel.run(function (ctx) {
                    var sheet = ctx.workbook.worksheets.getItem("DailyLog");
                    var range = sheet.getRange("A2:K" + (matrix.length + 1));
                    //tableDataRange.values = values;
                    range.values = matrix;
                    return ctx.sync().then(function () {
                        showMessage("Success! My monthly expense table created! Select the arrow button to see how to remove the table.");
                        //storeLineDaily(matrix[i]);
                    })
                     .catch(function (error) {
                         showMessage(JSON.stringify(error));
                     });
                });


            });
    };

    function addValues(rangeTable, row, ctx) {
        var sheet = ctx.workbook.worksheets.getItem("Values");
        var range = sheet.getRange(rangeTable)
        range.values = row;
    };

    function getDailyLog(dailyLogId) {
        deleteValues();
        var urlProject = host + '/api/DailyLog/GetDailyLog?logId=' + dailyLogId + '&email=' + localStorage.getItem("email");
        $.ajax({
            type: 'GET',
            url: urlProject,
            dataType: "json",
            contentType: "application/json; charset=utf-8",

        })
         .done(function (str) {
             //app.showNotification(str.impact[0].State);
             Excel.run(function (ctx) {
                 //var matrix = riskValuesImpact(str);
                 if (Object.keys(str.contextList).length != 0) {
                     //ctx.workbook.worksheets.getItem('Values').getRange("G1:G" + Object.keys(str.contextList).length).values = dailyLogContext(str);
                     var context = dailyLogContext(str);
                     addValues("B2:B" + (context.length + 1), context, ctx);
                 }
                 if (Object.keys(str.personnelContacts).length != 0) {
                     var contacts = dailyLogUsers(str);
                     addValues("D2:D" + (contacts.length + 1), contacts, ctx);
                     //ctx.workbook.worksheets.getItem('Values').getRange("I1:I" + Object.keys(str.personnelContacts).length).values = dailyLogUsers(str)
                 }
                 var status = [["Inbox"], ["Next"], ["Waiting"], ["Schedule"], ["Someday"], ["Done"]];
                 addValues("C2:C" + (status.length + 1), status, ctx);
                 var type = [["Problem"], ["Action"], ["Event"], ["Comment"], ["Decision"], ["Reference"]];
                 addValues("E2:E" + (type.length + 1), type, ctx);
                 var time = [["5 min"], ["15 min"], ["30 min"], ["1 hr"], ["2 hr"], ["4 hr"], ["8 hr"], [" - "]]/*[[1], [2], [1], [2], [1]] //str.impact[0].State*/;
                 addValues("F2:F" + (time.length + 1), time, ctx);
                 var energy = [["Mild"], ["Reasonable"], ["Demanding"], ["Very Demanding"], ["Extreme"], [" - "]]/*[[1], [2], [1], [2], [1]] //str.impact[0].State*/;
                 addValues("G2:G" + (energy.length + 1), energy, ctx);

                 if (Object.keys(str.project).length != 0) {
                     var projects = dailyLogProject(str);
                     addValues("A2:A" + (projects.length + 1), projects, ctx)
                     //ctx.workbook.worksheets.getItem('Values').getRange("M1:M" + Object.keys(str.project).length).values = dailyLogProject(str);
                 }

                 return ctx.sync().then(function () {
                     //console.log("Success! Insert range in A1:C3.");
                 });
             }).catch(function (error) {
                 //app.showNotification(error);
             });
         })
    };

    function dailyLogContext(str) {
        var val = [Object.keys(str.contextList).length];
        for (var i = 0; i < Object.keys(str.contextList).length; i++) {
            val[i] = [1];
            val[i][0] = str.contextList[i].description;
            localStorage.setItem('dailyLogContext' + str.contextList[i].description, "" + str.contextList[i].id);
            //val[i] = str.impact[i].State;
        }
        //app.showNotification(val[2][0]);
        return val;
    };

    function dailyLogUsers
        (str) {
        var val = [Object.keys(str.personnelContacts).length];
        for (var i = 0; i < Object.keys(str.personnelContacts).length; i++) {
            val[i] = [1];
            val[i][0] = str.personnelContacts[i].State;
            localStorage.setItem('dailyLogUsers' + str.personnelContacts[i].State, "" + str.personnelContacts[i].StateId);
            //val[i] = str.impact[i].State;
        }
        //app.showNotification(val[2][0]);
        return val;
    }

    function dailyLogProject(str) {
        var val = [Object.keys(str.project).length];
        for (var i = 0; i < Object.keys(str.project).length; i++) {
            val[i] = [1];
            val[i][0] = str.project[i].State;
            localStorage.setItem('dailyLogProject' + str.project[i].State, "" + str.project[i].StateId);
            //val[i] = str.impact[i].State;
        }
        //app.showNotification(val[2][0]);
        return val;
    }

    function publishDailyLog() {
        Excel.run(function (ctx) {
            var rows = ctx.workbook.tables.getItem("DailyLog").rows.load("values");
            var table = ctx.workbook.tables.getItem("DailyLog");
            table.clearFilters();
            /*for (var i = 0; i < 11; i++) {
                var column = ctx.workbook.tables.getItem("DailyLog").columns.getItemAt(i);
                column.filter.clear();
            }
            */
            return ctx.sync()
                .then(function () {
                    var DlId = localStorage.getItem('DlId');
                    var projectId = localStorage.getItem('projectId');
                    for (var i = 0; i < rows.items.length; i++) {
                        var row = rows.items[i];
                        if (isNewDaily(row)) {
                            commitNewDaily(row, i);
                        }
                        else {
                            if (editedLine(row)) {
                                commitUpdateDaily(row);
                            }
                        };
                    }
                })
                .then(ctx.sync)
                .then(function () {
                    console.log("Success! Format rows of 'Table1' with 2nd cell greater than 2 in green, other rows in red.");
                });
        }).catch(function (error) {
            console.log(error);
        });
        Excel.run(function (ctx) {
            var tableName = 'DailyLog';
            var table = ctx.workbook.tables.getItem(tableName);
            table.reapplyFilters();
            return ctx.sync();
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    };

    function addId(range, id) {
        Excel.run(function (ctx) {
            var count = (range + 2);
            //app.showNotification(typeof(range) + "  "+ typeof(range + 1));
            var excelRange = "K" + count + ":K" + count;
            //app.showNotification(excelRange);
            ctx.workbook.worksheets.getItem('DailyLog').getRange(excelRange).values = ("" + id)/*[[1], [2], [1], [2], [1]] //str.impact[0].State*/;

            return ctx.sync().then(function () {
                //console.log("Success! Insert range in A1:C3.");
            });
        }).catch(function (error) {
            // app.showNotification(error);
        });
    };

    function isActivityType(type) {
        if (type == "Problem") return "0";
        else if (type == "Action") return "1";
        else if (type == "Event") return "2";
        else if (type == "Comment") return "3";
        else if (type == "Decision") return "4";
        else if (type == "Reference") return "5";
        else return null
    }

    function isEnergy(type) {
        if (type == "Mild") return "0";
        else if (type == "Reasonable") return "1";
        else if (type == "Demanding") return "2";
        else if (type == "Very Demanding") return "3";
        else if (type == "Extreme") return "4";
        else return null;
    };

    function isTime(type) {
        if (type == "5 min") return "0";
        else if (type == "15 min") return "1";
        else if (type == "30 min") return "2";
        else if (type == "1 hr") return "3";
        else if (type == "2 hr") return "4";
        else if (type == "4 hr") return "5";
        else if (type == "8 hr") return "6";
        else return null;
    };

    function isResponseStatus(type) {
        if (type == "Inbox") return "0";
        else if (type == "Next") return "1";
        else if (type == "Waiting") return "2";
        else if (type == "Schedule") return "3";
        else if (type == "Someday") return "4";
        else if (type == "Done") return "5";
        else return null;
    };

    function commitNewDaily(line, i) {
        var email = localStorage.getItem("email");
        var urlProject = host + '/api/DailyLog/PostDailyLogHeader';
        var urlProject2 = host + '/api/DailyLog/PostDailyLogInvolvedTiming';
        var dataEmail = {
            "id": null,
            "projectId": localStorage.getItem("dailyLogProject" + line.values[0][0]),
            "activityTypeId": isActivityType(line.values[0][7]),
            "title": isNull(line.values[0][1]),
            "order": "0",
            "coreUserEmail": email,
            "authorEmail": email,
            "context": localStorage.getItem("dailyLogContext" + line.values[0][3]),
            "energy": isEnergy(line.values[0][9])
        };
        $.ajax({
            type: "POST",
            url: urlProject,
            dataType: "json",
            async: false,
            contentType: "application/json; charset=utf-8",
            data: JSON.stringify(dataEmail),
        })
        .done(function (result) {
            addId(i, result);
            var dataEmail2 = {
                "dailyLogId": result,
                "responsible": localStorage.getItem('dailyLogUsers' + line.values[0][6]),
                "responseStatus": isResponseStatus(line.values[0][5]),
                "targetDate": formatDate3(line.values[0][4]),
                "time": isTime(line.values[0][8])
            };
            $.ajax({
                type: "POST",
                url: urlProject2,
                dataType: "json",
                contentType: "application/json; charset=utf-8",
                data: JSON.stringify(dataEmail2),
            });
        })
    };

    function commitUpdateDaily(line) {
        //app.showNotification(line.values[0].toString());
        var urlProject = host + '/api/DailyLog/PostDailyLogHeader';
        var urlProject2 = host + '/api/DailyLog/PostDailyLogInvolvedTiming';
        var dataEmail = {
            "id": line.values[0][2],
            "projectId": localStorage.getItem("dailyLogProject" + line.values[0][0]),
            "activityTypeId": isActivityType(line.values[0][7]),
            "title": isNull(line.values[0][1]),
            "order": "0",
            "coreUserEmail": localStorage.getItem("email"),
            "authorEmail": localStorage.getItem("email"),
            "context": localStorage.getItem("dailyLogContext" + line.values[0][3]),
            "energy": isEnergy(line.values[0][9])
        };
        $.ajax({
            type: "POST",
            url: urlProject,
            dataType: "json",
            contentType: "application/json; charset=utf-8",
            data: JSON.stringify(dataEmail),
        })
        var dataEmail2 = {
            "dailyLogId": line.values[0][2],
            "responsible": localStorage.getItem('dailyLogUsers' + line.values[0][6]),
            "responseStatus": isResponseStatus(line.values[0][5]),
            "targetDate": formatDate3(line.values[0][4]),
            "time": isTime(line.values[0][8])
        };
        //app.showNotification(line.values[0][4]);
        //line.values[0][4] = "";
        $.ajax({
            type: "POST",
            url: urlProject2,
            dataType: "json",
            contentType: "application/json; charset=utf-8",
            data: JSON.stringify(dataEmail2),
        });
        //app.showNotification(line.values[0].toString());
    };

    function isNewDaily(line) {
        if (line.values[0][10] == null || line.values[0][10] == "") return true;
        else return false;
    }

    function storeLineDaily(line) {
        //line[4] = formatDate(line[4]);
        localStorage.setItem("line" + line[2], line.toString());
        //app.showNotification(line.toString());
        //app.showNotification(line[4]);
    }

    function editedLine(line) {
        var oldLine = localStorage.getItem("line" + line.values[0][10]);
        var temporary = line.values[0][4];
        line.values[0][4] = formatDate3(line.values[0][4]);
        //app.showNotification(oldLine + " / " + line.values[0].toString());
        if (line.values[0].toString() != oldLine) {
            line.values[0][4] = temporary;
            return true;
        }
        else {
            line.values[0][4] = temporary;
            return false
        };
    }

    return app;
})();