async function main(workbook: ExcelScript.Workbook) {
    let sht = workbook.getWorksheet('Sheet1');
    let sht2 = workbook.getWorksheet('Sheet2');
    let tHeads = {
    "Content-Type": "application/json",
    "Accept": "*/*",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive"
    };
    var d = new Date();
    if (new Date(String(sht2.getRange("F1").getValues())).getTime() < d.getTime())
    {
    //get token
    let tParams = {
    "ident": "ceavco",
    "password": “********”,
    "username": "richboulton@ceavco.com",
    };
    let response = await
    fetch('https://webapi1.ielightning.net/api/v1/Authentication/GetAccessToken',
    {
    method: "POST",
    headers: tHeads,
    body: JSON.stringify(tParams),
    }
    );
    
    let tToken: { accessToken: ""; refreshToken: ""; expirationDate: "" } = await response.json();
    //let token = "Bearer " + tToken.accessToken;
    sht2.getRange("E1").setValue(`${tToken.accessToken}`);
    sht2.getRange("F1").setValue(`${tToken.expirationDate}`);
    var nowtime = d
    console.log(new Date(tToken.expirationDate))
    console.log(d)
    console.log(new Date(tToken.expirationDate).getTime() - d.getTime())
    
    }
    let token = "Bearer " + sht2.getRange("E1").getValue()
    var sDate = "2023-01-01";
    var eDate = "2023-12-31";
    // set headers for all requests
    let heads = {
    "Authorization": token,
    "Content-Type": "application/json; charset=utf-8",
    "Accept": "*/*",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive"
    };
    
    // get all breakout coordinator positions
    let params = {
    "startDate": sDate,
    "endDate": eDate,
    "officeId": 0,
    "orderStatus": "Active,Quote Only,Option,Tentative,Invoiced,Accounting,Routing,Availability",
    "talentId": 498,
    "taskId": 0,
    "resourceId": 0,
    "workStatus": "",
    "jobTypeId": 0
    };
    let response = await
    fetch('https://webapi.ielightning.net/api/v1/Reports/Labor/LaborGanttView/List',
    {
    method: "POST",
    headers: heads,
    body: JSON.stringify(params),
    }
    );
    
    let sched1: [] = await response.json()
    //get all Onsite Coordinator/Manager
    params = {
    "startDate": sDate,
    "endDate": eDate,
    "officeId": 0,
    "orderStatus": "Active,Quote Only,Option,Tentative,Invoiced,Accounting,Routing,Availability",
    "talentId": 503,
    "taskId": 0,
    "resourceId": 0,
    "workStatus": "",
    "jobTypeId": 0
    };
    response = await fetch('https://webapi.ielightning.net/api/v1/Reports/Labor/LaborGanttView/List',
    {
    method: "POST",
    headers: heads,
    body: JSON.stringify(params),
    }
    );
    let sched2: [] = await response.json()
    
    // get all Project Manager
    params = {
    "startDate": sDate,
    
    "endDate": eDate,
    "officeId": 0,
    "orderStatus": "Active,Quote Only,Option,Tentative,Invoiced,Accounting,Routing,Availability",
    "talentId": 539,
    "taskId": 0,
    "resourceId": 0,
    "workStatus": "",
    "jobTypeId": 0
    };
    
    // get data from server
    response = await fetch('https://webapi.ielightning.net/api/v1/Reports/Labor/LaborGanttView/List',
    {
    method: "POST",
    headers: heads,
    body: JSON.stringify(params),
    }
    );
    let sched3: [] = await response.json()
    
    //Get Struble schedule
    params = {
    "startDate": sDate,
    "endDate": eDate,
    "officeId": 0,
    "orderStatus": "Active,Quote Only,Option,Tentative,Invoiced,Accounting,Routing,Availability",
    "talentId": 0,
    "taskId": 0,
    "resourceId": 1484,
    "workStatus": "",
    "jobTypeId": 0
    };
    
    // get data from server
    response = await fetch('https://webapi.ielightning.net/api/v1/Reports/Labor/LaborGanttView/List',
    {
    method: "POST",
    headers: heads,
    body: JSON.stringify(params),
    }
    );
    let sched4: [] = await response.json()
    //Get Safarik schedule
    params = {
    "startDate": sDate,
    "endDate": eDate,
    "officeId": 0,
    
    "orderStatus": "Active,Quote Only,Option,Tentative,Invoiced,Accounting,Routing,Availability",
    "talentId": 0,
    "taskId": 0,
    "resourceId": 1480,
    "workStatus": "",
    "jobTypeId": 0
    };
    
    // get data from server
    response = await fetch('https://webapi.ielightning.net/api/v1/Reports/Labor/LaborGanttView/List',
    {
    method: "POST",
    headers: heads,
    body: JSON.stringify(params),
    }
    );
    let sched5: [] = await response.json()
    //Get Rich B schedule
    params = {
    "startDate": sDate,
    "endDate": eDate,
    "officeId": 0,
    "orderStatus": "Active,Quote Only,Option,Tentative,Invoiced,Accounting,Routing,Availability",
    "talentId": 0,
    "taskId": 0,
    "resourceId": 1459,
    "workStatus": "",
    "jobTypeId": 0
    };
    
    // get data from server
    response = await fetch('https://webapi.ielightning.net/api/v1/Reports/Labor/LaborGanttView/List',
    {
    method: "POST",
    headers: heads,
    body: JSON.stringify(params),
    }
    );
    let sched6: [] = await response.json()
    //Get Emily S schedule
    params = {
    "startDate": sDate,
    "endDate": eDate,
    "officeId": 0,
    "orderStatus": "Active,Quote Only,Option,Tentative,Invoiced,Accounting,Routing,Availability",
    "talentId": 0,
    "taskId": 0,
    
    "resourceId": 1481,
    "workStatus": "",
    "jobTypeId": 0
    };
    
    // get data from server
    response = await fetch('https://webapi.ielightning.net/api/v1/Reports/Labor/LaborGanttView/List',
    {
    method: "POST",
    headers: heads,
    body: JSON.stringify(params),
    }
    );
    let sched7: [] = await response.json()
    //Get Andrew L schedule
    params = {
    "startDate": sDate,
    "endDate": eDate,
    "officeId": 0,
    "orderStatus": "Active,Quote Only,Option,Tentative,Invoiced,Accounting,Routing,Availability",
    "talentId": 0,
    "taskId": 0,
    "resourceId": 2692,
    "workStatus": "",
    "jobTypeId": 0
    };
    
    // get data from server
    response = await fetch('https://webapi.ielightning.net/api/v1/Reports/Labor/LaborGanttView/List',
    {
    method: "POST",
    headers: heads,
    body: JSON.stringify(params),
    }
    );
    let sched8: [] = await response.json()
    //Get Brett K schedule
    params = {
    "startDate": sDate,
    "endDate": eDate,
    "officeId": 0,
    "orderStatus": "Active,Quote Only,Option,Tentative,Invoiced,Accounting,Routing,Availability",
    "talentId": 0,
    "taskId": 0,
    "resourceId": 1469,
    "workStatus": "",
    "jobTypeId": 0
    
    };
    
    // get data from server
    response = await fetch('https://webapi.ielightning.net/api/v1/Reports/Labor/LaborGanttView/List',
    {
    method: "POST",
    headers: heads,
    body: JSON.stringify(params),
    }
    );
    let sched9: [] = await response.json()
    //Get Zak schedule
    params = {
    "startDate": sDate,
    "endDate": eDate,
    "officeId": 0,
    "orderStatus": "Active,Quote Only,Option,Tentative,Invoiced,Accounting,Routing,Availability",
    "talentId": 0,
    "taskId": 0,
    "resourceId": 2714,
    "workStatus": "",
    "jobTypeId": 0
    };
    
    // get data from server
    response = await fetch('https://webapi.ielightning.net/api/v1/Reports/Labor/LaborGanttView/List',
    {
    method: "POST",
    headers: heads,
    body: JSON.stringify(params),
    }
    );
    let sched10: [] = await response.json()
    //Get Stephen B schedule
    params = {
    "startDate": sDate,
    "endDate": eDate,
    "officeId": 0,
    "orderStatus": "Active,Quote Only,Option,Tentative,Invoiced,Accounting,Routing,Availability",
    "talentId": 0,
    "taskId": 0,
    "resourceId": 1456,
    "workStatus": "",
    "jobTypeId": 0
    };
    
    // get data from server
    response = await fetch('https://webapi.ielightning.net/api/v1/Reports/Labor/LaborGanttView/List',
    {
    method: "POST",
    headers: heads,
    body: JSON.stringify(params),
    }
    );
    let sched11: [] = await response.json()
    /*
    //Get Paul T schedule
    params = {
    "startDate": sDate,
    "endDate": eDate,
    "officeId": 0,
    "orderStatus": "Active,Quote Only,Option,Tentative,Invoiced,Accounting,Routing,Availability",
    "talentId": 0,
    "taskId": 0,
    "resourceId": 1486,
    "workStatus": "",
    "jobTypeId": 0
    };
    
    // get data from server
    response = await fetch('https://webapi.ielightning.net/api/v1/Reports/Labor/LaborGanttView/List',
    {
    method: "POST",
    headers: heads,
    body: JSON.stringify(params),
    }
    );
    let sched12: [] = await response.json()
    //Get Michael W schedule
    params = {
    "startDate": sDate,
    "endDate": eDate,
    "officeId": 0,
    "orderStatus": "Active,Quote Only,Option,Tentative,Invoiced,Accounting,Routing,Availability",
    "talentId": 0,
    "taskId": 0,
    "resourceId": 1491,
    "workStatus": "",
    "jobTypeId": 0
    };
    
    // get data from server
    response = await fetch('https://webapi.ielightning.net/api/v1/Reports/Labor/LaborGanttView/List',
    {
    
    method: "POST",
    headers: heads,
    body: JSON.stringify(params),
    }
    );
    let sched13: [] = await response.json()
    //Get Randy C schedule
    params = {
    "startDate": sDate,
    "endDate": eDate,
    "officeId": 0,
    "orderStatus": "Active,Quote Only,Option,Tentative,Invoiced,Accounting,Routing,Availability",
    "talentId": 0,
    "taskId": 0,
    "resourceId": 1461,
    "workStatus": "",
    "jobTypeId": 0
    };
    
    // get data from server
    response = await fetch('https://webapi.ielightning.net/api/v1/Reports/Labor/LaborGanttView/List',
    {
    method: "POST",
    headers: heads,
    body: JSON.stringify(params),
    }
    );
    let sched14: [] = await response.json()
    //Get Jeff P schedule
    params = {
    "startDate": sDate,
    "endDate": eDate,
    "officeId": 0,
    "orderStatus": "Active,Quote Only,Option,Tentative,Invoiced,Accounting,Routing,Availability",
    "talentId": 0,
    "taskId": 0,
    "resourceId": 2671,
    "workStatus": "",
    "jobTypeId": 0
    };
    
    // get data from server
    response = await fetch('https://webapi.ielightning.net/api/v1/Reports/Labor/LaborGanttView/List',
    {
    method: "POST",
    headers: heads,
    
    body: JSON.stringify(params),
    }
    );
    let sched15: [] = await response.json()*/
    let sched: {} = Array.from(new Set([...sched1, ...sched2, ...sched3, ...sched4, ...sched5,
    ...sched6, ...sched7, ...sched8, ...sched9, ...sched10, ...sched11/*, ...sched12, ...sched13,
    ...sched14, ...sched15*/]))
    //let schedx: {} = Array.from(new Set([...sched1, ...sched2,...sched3, ...sched4]))
    var rname = "name";
    var rstart = "2022-11-21T10:13:01";
    var rend = "2022-11-21T10:13:01";
    var rref = 12345;
    
    //sort by date
    sched.sort((a, b) => {
    return new Date(a.startDate) - new Date(b.startDate);
    })
    //parse through and add row nums to TBD Jobs
    var rname = sched[0].resourceName;
    var rstart = sched[0].startDate.substring(0, 10);
    var rend = sched[0].endDate.substring(0, 10);
    var rref = sched[0].rowReference;
    var rtal: string = sched[0].talentName;
    //work out TBD job stacking
    sched.forEach((s) => {
    if (
    s.resourceName == "TBD TBD" && s.disprow == undefined
    ) {
    var i = 1
    while (sched.filter(function (f) {
    return Date.parse(f.startDate.substring(0, 10)) == Date.parse(s.startDate.substring(0, 10)) &&
    f.disprow == i
    }).length > 0) { i = i + 1 }
    sched.forEach((z) => {
    if (
    z.rowReference == s.rowReference && z.talentName == s.talentName && z.resourceName ==
    s.resourceName
    ) { z.disprow = i }
    })
    }
    });
    //sort by name
    sched.sort((a, b) => {
    let fa: string = a.resourceName.toLowerCase(),
    fb: string = b.resourceName.toLowerCase();
    if (fa < fb) {
    return -1;
    
    }
    if (fa > fb) {
    return 1;
    }
    return 0;
    });
    //Sort by job# (rowreference)
    sched.sort((a, b) => {
    return a.rowReference - b.rowReference;
    })
    //sort by role(talent)
    sched.sort((a, b) => {
    let fa: string = a.talentName.toLowerCase(),
    fb: string = b.talentName.toLowerCase();
    if (fa < fb) {
    return -1;
    }
    if (fa > fb) {
    return 1;
    }
    return 0;
    });
    var index = 0
    sched.forEach((s) => {
    s.index = index
    index++
    });
    //parse through and remove duplicates
    var rname = sched[0].resourceName;
    var rstart = sched[0].startDate.substring(0, 10);
    var rend = sched[0].endDate.substring(0, 10);
    var rref = sched[0].rowReference;
    var rtal: string = sched[0].talentName;
    sched.forEach((s) => {
    if (s.index > 0) {
    //console.log("tindex:" + s.index);
    if (
    (s.resourceName == rname &&
    s.rowReference == rref &&
    s.talentName == rtal && /*
    s.startDate.substring(0, 10) == sched[s.index - 1].startDate.substring(0, 10) &&
    s.endDate.substring(0, 10) == sched[s.index - 1].endDate.substring(0, 10) */
    Date.parse(s.startDate.substring(0, 10)) == Date.parse(sched[s.index - 1].startDate.substring(0,
    10)) &&
    Date.parse(s.endDate.substring(0, 10)) == Date.parse(sched[s.index - 1].endDate.substring(0,
    
    10))) || s.resourceName == "DO NOT FILL" || s.taskName == ("Pre-Production" || "Post-
    Production") /*|| s.resourceId == */ || s.rowReference == 0
    
    ) {
    s.filterout = true
    }
    
    rname = s.resourceName;
    rstart = s.startDate.substring(0, 10);
    rend = s.endDate.substring(0, 10);
    rref = s.rowReference;
    rtal = s.talentName
    }
    });
    //parse through and combine multi line jobs into one line
    var rname = sched[0].resourceName;
    var rstart = sched[0].startDate.substring(0, 10);
    var rend = sched[0].endDate.substring(0, 10);
    var rref = sched[0].rowReference;
    var rtal = sched[0].talentName;
    sched.filter(function (f) {
    return f.filterout !== true;
    }).forEach((s) => {
    if (s.index > 0) {
    //console.log("tindex:" + s.index);
    if (
    s.resourceName == rname &&
    s.rowReference == rref &&
    s.talentName == rtal &&
    Date.parse(s.startDate.substring(0, 10)) == Date.parse(sched[s.index - 1].endDate.substring(0,
    10)) + 86400000
    ) {
    s.filterout = true
    sched[
    sched.indexOf(
    sched.find(r =>
    r.resourceName === s.resourceName &&
    r.rowReference === s.rowReference &&
    r.talentName === s.talentName &&
    Date.parse(r.endDate.substring(0, 10)) == (Date.parse(s.startDate.substring(0, 10)) - 86400000)
    )
    )
    ].endDate = s.endDate
    }
    rname = s.resourceName;
    rstart = s.startDate.substring(0, 10);
    rend = s.endDate.substring(0, 10);
    rref = s.rowReference;
    rtal = s.talentName;
    }
    });
    /* for (var i = 1; i < sched.length; i++) {
    var rname = sched[i - 1].resourceName;
    var rstart = sched[i - 1].startDate.substring(0, 10);
    var rend = sched[i - 1].endDate.substring(0, 10);
    var rref = sched[i - 1].rowReference;
    if (
    sched[i].resourceName == rname &&
    
    sched[i].rowReference == rref
    //sched[i].startDate.substring(0,10) == rstart + 1 &&
    //sched[i].endDate.substring(0,10) == sched[i].startDate.substring(0,10)
    ) {
    // sched[i-1].endDate = sched[i].endDate
    // delete sched[i]
    // how to find a specific row
    // sched.find(({ resourceName }) => resourceName === "EMILY SENDELBACH")
    }
    } */
    let row = 1
    sht.getRange("A:K").clear(ExcelScript.ClearApplyTo.contents);
    // sht.getRange("A:T").clear
    //sched.filter(function (f) {
    // return f.filterout !== true;
    // }).forEach((s) => {
    //console.log(`${s.index} ${s.rowReference} ${s.resourceName} ${s.startDate.substring(0, 10)}
    ${s.endDate.substring(0, 10)} ${s.talentName}`);
    // });
    sched.filter(function (f) {
    return f.filterout !== true;
    }).forEach((s) => {
    sht.getRange("A" + row).setValue(`${s.index}`);
    sht.getRange("B" + row).setValue(`${s.rowReference}`);
    sht.getRange("C" + row).setValue(`${s.resourceName}`);
    sht.getRange("D" + row).setValue(`${s.startDate.substring(0, 10)}`);
    sht.getRange("E" + row).setValue(`${s.endDate.substring(0, 10)}`);
    sht.getRange("F" + row).setValue(`${s.talentName}`);
    sht.getRange("G" + row).setValue(`${s.rowDescription}`);
    sht.getRange("H" + row).setValue(`${s.jobBarBackgroundColor}`);
    sht.getRange("I" + row).setValue(`${s.orderStatus}`);
    sht.getRange("j" + row).setValue("=E" + row + "-D" + row);
    sht.getRange("k" + row).setValue(`${s.disprow}`);
    sht.getRange("l" + row).setValue(`${s.taskName}`);
    row++
    });
    sht.getRange("M1").setValue(row - 1);
    console.log(sched)
    }