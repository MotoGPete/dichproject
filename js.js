async function main(workbook) {
    let sht = workbook.getWorksheet('Sheet1');
    let sht2 = workbook.getWorksheet('Sheet2');
    let tHeads = {
        "Content-Type": "application/json",
        "Accept": "*/*",
        "Accept-Encoding": "gzip, deflate, br",
        "Connection": "keep-alive"
    };

    var d = new Date();
    if (new Date(String(sht2.getRange("F1").getValues())).getTime() < d.getTime()) {
      // get token
    let tParams = {
        "ident": "ceavco",
        "password": "********",
        "username": "richboulton@ceavco.com",
    };

    let response = await fetch('https://webapi1.ielightning.net/api/v1/Authentication/GetAccessToken', {
        method: "POST",
        headers: tHeads,
        body: JSON.stringify(tParams),
    });
    let tToken = await response.json();
    sht2.getRange("E1").setValue(`${tToken.accessToken}`);
    sht2.getRange("F1").setValue(`${tToken.expirationDate}`);
    var nowtime = d;
    console.log(new Date(tToken.expirationDate));
    console.log(d);
    console.log(new Date(tToken.expirationDate).getTime() - d.getTime());
    }

    let token = "Bearer " + sht2.getRange("E1").getValue();
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
    params = {
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

    response = await fetch('https://webapi.ielightning.net/api/v1/Reports/Labor/LaborGanttView/List', {
        method: "POST",
        headers: heads,
        body: JSON.stringify(params),
    });

    sched1 = await response.json();

    // get all Onsite Coordinator/Manager
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
    response = await fetch('https://webapi.ielightning.net/api/v1/Reports/Labor/LaborGanttView/List', {
        method: "POST",
        headers: heads,
        body: JSON.stringify(params),
    });

    let sched2 = await response.json();

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
let response = await fetch('https://webapi.ielightning.net/api/v1/Reports/Labor/LaborGanttView/List', {
    method: "POST",
    headers: heads,
    body: JSON.stringify(params),
});
let sched3 = await response.json();

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
response = await fetch('https://webapi.ielightning.net/api/v1/Reports/Labor/LaborGanttView/List', {
    method: "POST",
    headers: heads,
    body: JSON.stringify(params),
});
let sched4 = await response.json();

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
response = await fetch('https://webapi.ielightning.net/api/v1/Reports/Labor/LaborGanttView/List', {
    method: "POST",
    headers: heads,
    body: JSON.stringify(params),
});
let sched5 = await response.json();

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
response = await fetch('https://webapi.ielightning.net/api/v1/Reports/Labor/LaborGanttView/List', {
    method: "POST",
    headers: heads,
    body: JSON.stringify(params),
});
let sched6 = await response.json();

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
let response = await fetch('https://webapi.ielightning.net/api/v1/Reports/Labor/LaborGanttView/List', {
    method: "POST",
    headers: heads,
    body: JSON.stringify(params)
});

let sched7 = await response.json();

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
response = await fetch('https://webapi.ielightning.net/api/v1/Reports/Labor/LaborGanttView/List', {
    method: "POST",
    headers: heads,
    body: JSON.stringify(params)
});

let sched8 = await response.json();

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
response = await fetch('https://webapi.ielightning.net/api/v1/Reports/Labor/LaborGanttView/List', {
    method: "POST",
    headers: heads,
    body: JSON.stringify(params)
});

let sched9 = await response.json();

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
response = await fetch('https://webapi.ielightning.net/api/v1/Reports/Labor/LaborGanttView/List', {
    method: "POST",
    headers: heads,
    body: JSON.stringify(params)
});

let sched10 = await response.json();

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
response = await fetch('https://webapi.ielightning.net/api/v1/Reports/Labor/LaborGanttView/List', {
    method: "POST",
    headers: heads,
    body: JSON.stringify(params)
});
let sched11 = await response.json();

// Get Paul T schedule
let params = {
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

// Get data from server
let response = await fetch('https://webapi.ielightning.net/api/v1/Reports/Labor/LaborGanttView/List', {
  method: "POST",
  headers: heads,
  body: JSON.stringify(params),
});
let sched1 = await response.json();

// Get Michael W schedule
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

// Get data from server
response = await fetch('https://webapi.ielightning.net/api/v1/Reports/Labor/LaborGanttView/List', {
  method: "POST",
  headers: heads,
  body: JSON.stringify(params),
});
let sched2 = await response.json();

// Get Randy C schedule
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

// Get data from server
response = await fetch('https://webapi.ielightning.net/api/v1/Reports/Labor/LaborGanttView/List', {
  method: "POST",
  headers: heads,
  body: JSON.stringify(params),
});
let sched3 = await response.json();

// Get Jeff P schedule
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

// Get data from server
response = await fetch('https://webapi.ielightning.net/api/v1/Reports/Labor/LaborGanttView/List', {
  method: "POST",
  headers: heads,
  body: JSON.stringify(params),
});
let sched4 = await response.json();

let sched = Array.from(new Set([...sched1, ...sched2, ...sched3, ...sched4, ...sched5, ...sched6, ...sched7, ...sched8, ...sched9, ...sched10, ...sched11/*, ...sched12, ...sched13, ...sched14, ...sched15*/]));

// Sort by date
sched.sort((a, b) => {
    return new Date(a.startDate) - new Date(b.startDate);
  });
  
  let rname = sched[0].resourceName;
  let rstart = sched[0].startDate.substring(0, 10);
  let rend = sched[0].endDate.substring(0, 10);
  let rref = sched[0].rowReference;
  let rtal = sched[0].talentName;
  
  sched.forEach((s) => {
    if (s.resourceName == "TBD TBD" && s.disprow == undefined) {
      let i = 1;
      while (sched.filter((f) => {
        return (
          Date.parse(f.startDate.substring(0, 10)) ==
            Date.parse(s.startDate.substring(0, 10)) && f.disprow == i
        );
      }).length > 0) {
        i++;
      }
      sched.forEach((z) => {
        if (
          z.rowReference == s.rowReference &&
          z.talentName == s.talentName &&
          z.resourceName == s.resourceName
        ) {
          z.disprow = i;
        }
      });
    }
  });
  
  sched.sort((a, b) => {
    let fa = a.resourceName.toLowerCase(),
      fb = b.resourceName.toLowerCase();
    if (fa < fb) {
      return -1;
    }
    if (fa > fb) {
      return 1;
    }
    return 0;
  });
  
  sched.sort((a, b) => {
    return a.rowReference - b.rowReference;
  });
  
  sched.sort((a, b) => {
    let fa = a.talentName.toLowerCase(),
      fb = b.talentName.toLowerCase();
    if (fa < fb) {
      return -1;
    }
    if (fa > fb) {
      return 1;
    }
    return 0;
  });
  
  let index = 0;
  sched.forEach((s) => {
    s.index = index;
    index++;
  });
  
  let rname = sched[0].resourceName;
  let rstart = sched[0].startDate.substring(0, 10);
  let rend = sched[0].endDate.substring(0, 10);
  let rref = sched[0].rowReference;
  let rtal = sched[0].talentName;
  
  sched.forEach((s) => {
    if (s.index > 0) {
      if (
        (s.resourceName == rname &&
          s.rowReference == rref &&
          s.talentName == rtal &&
          Date.parse(s.startDate.substring(0, 10)) ==
            Date.parse(sched[s.index - 1].startDate.substring(0, 10)) &&
          Date.parse(s.endDate.substring(0, 10)) ==
            Date.parse(sched[s.index - 1].endDate.substring(0, 10))) ||
        s.resourceName == "DO NOT FILL" ||
        s.taskName == ("Pre-Production" || "Post-Production") ||
        s.rowReference == 0
      ) {
        s.filterout = true;
      }
  
      rname = s.resourceName;
      rstart = s.startDate.substring(0, 10);
      rend = s.endDate.substring(0, 10);
      rref = s.rowReference;
      rtal = s.talentName;
    }
  });
  
  let rname = sched[0].resourceName;
  let rstart = sched[0].startDate.substring(0, 10);
  let rend = sched[0].endDate.substring(0, 10);
  let rref = sched[0].rowReference;
  let rtal = sched[0].talent
  