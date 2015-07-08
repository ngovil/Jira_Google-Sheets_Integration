C_MAX_RESULTS = 1000;
vals = new Array();
vals2 = new Array();
colNums = new Array();
jiraHeads = new Array();
colHeads = new Array();
pCol=0;
sCol=0;
rCol=0;
dCol=0;
ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
lRow=ss.getLastRow();
sheeturl = SpreadsheetApp.getActiveSpreadsheet().getUrl();

function onOpen(){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  jiraConfigure();
  var menuEntries = [{name: "Configure Headings", functionName: "doGet"}, {name: "Refresh Now", functionName: "jiraPullManual"}, {name: "Credentials", functionName: "userandpass"}]; 
  ss.addMenu("Jira", menuEntries);
}

function jiraConfigure() {
  PropertiesService.getUserProperties().setProperty("host", "jira.naehas.com");
  PropertiesService.getUserProperties().setProperty("digest", "No");
}

function userandpass(){
  var userAndPassword = Browser.inputBox("Enter your Jira User id and Password in the form Username:Password. e.g. John.Doe:ilovejira", "Userid:Password", Browser.Buttons.OK_CANCEL);
  var x = Utilities.base64Encode(userAndPassword);
  PropertiesService.getUserProperties().setProperty("digest", "Basic " + x);
}

function doGet() {
    var jiraKey = Browser.inputBox("What is the heading of your JIRA key column? e.g. JIRA", "", Browser.Buttons.OK_CANCEL);
    PropertiesService.getUserProperties().setProperty("jira" + sheeturl, jiraKey);
    var LIST = ["Priority", "Status", "fixVersions", "issueType", "Summary", "Description", "dueDate", "Resolution", "Reporter", "Assignee"];
    var mydoc = SpreadsheetApp.getActiveSpreadsheet();
    var app = UiApp.createApplication().setTitle("Jira Headings").setWidth(100).setHeight(200);
    var panel = app.createVerticalPanel().setId('panel');
    // Store the number of items in the array (LIST)
    panel.add(app.createHidden('checkbox_total', LIST.length));
    // add 1 checkbox + 1 hidden field per item 
    for(var i = 0; i < LIST.length; i++){
      var checkbox = app.createCheckBox().setName('checkbox_isChecked_'+i).setText(LIST[i]);
      var hidden = app.createHidden('checkbox_value_'+i, LIST[i]);
      panel.add(checkbox).add(hidden);
    }
    var handler = app.createServerHandler('submit').addCallbackElement(panel);
    panel.add(app.createButton('Submit', handler));
    var scroll = app.createScrollPanel().setPixelSize(100, 200);
    scroll.add(panel);
    app.add(scroll);
    mydoc.show(app);
}

function submit(e){
  var numberOfItems = e.parameter.checkbox_total;
  var itemsSelected = new Array();
  // for each item, if it is checked / selected, add it to itemsSelected
  for(var i = 0; i < numberOfItems; i++){
    if(e.parameter['checkbox_isChecked_'+i] == 'true'){
      itemsSelected.push(e.parameter['checkbox_value_'+i]);
    }
  }
  jiraHeads = itemsSelected;
  for(var i=0; i<jiraHeads.length; i++){
     colHeads.push(Browser.inputBox("What is the heading of your " + jiraHeads[i] + " column?", "", Browser.Buttons.OK_CANCEL));
  }
  PropertiesService.getUserProperties().setProperty("jiraHeads" + sheeturl, jiraHeads.toString());
  PropertiesService.getUserProperties().setProperty("colHeads" + sheeturl, colHeads.toString());
  /*var app = UiApp.getActiveApplication();
  app.getElementById('panel').clear().add(app.createLabel(itemsSelected));
  return app;*/
}

function configureIt(){
  var jiraKey = Browser.inputBox("What is the heading of your JIRA key column? e.g. JIRA", "", Browser.Buttons.OK_CANCEL);
  PropertiesService.getUserProperties().setProperty("jira" + sheeturl, jiraKey);
  var includeP = Browser.inputBox("Do you want to include a priority column?", "Yes or No", Browser.Buttons.OK_CANCEL);
  if(includeP.toLowerCase() == "yes"){
    colHeads.push(Browser.inputBox("What is the heading of your priority column? e.g. Priority", "", Browser.Buttons.OK_CANCEL));
    jiraHeads.push("Priority");
  }
  var includeS = Browser.inputBox("Do you want to include a status column?", "Yes or No", Browser.Buttons.OK_CANCEL);
  if(includeS.toLowerCase() == "yes"){
      colHeads.push(Browser.inputBox("What is the heading of your status column? e.g. Status", "", Browser.Buttons.OK_CANCEL));
      jiraHeads.push("Status");
  }
  var includeR = Browser.inputBox("Do you want to include a fix version column?", "Yes or No", Browser.Buttons.OK_CANCEL);
  if(includeR.toLowerCase() == "yes"){
      colHeads.push(Browser.inputBox("What is the heading of your fix version column? e.g. Release", "", Browser.Buttons.OK_CANCEL));
      jiraHeads.push("fixVersions");
  }
  var includeD = Browser.inputBox("Do you want to include a type column?", "Yes or No", Browser.Buttons.OK_CANCEL);
  if(includeD.toLowerCase() == "yes"){
      colHeads.push(Browser.inputBox("What is the heading of your type column? e.g. Dev Category", "", Browser.Buttons.OK_CANCEL));
      jiraHeads.push("Type");
  }
  PropertiesService.getUserProperties().setProperty("jiraHeads" + sheeturl, jiraHeads.toString());
  PropertiesService.getUserProperties().setProperty("colHeads" + sheeturl, colHeads.toString());
  /*var includeP = Browser.inputBox("Do you want to include a priority column?", "Yes or No", Browser.Buttons.OK_CANCEL);
  if(includeP.toLowerCase() == "yes"){
    var pHead = Browser.inputBox("What is the heading of your priority column? e.g. Priority", "", Browser.Buttons.OK_CANCEL);
    PropertiesService.getUserProperties().setProperty("priority" + sheeturl, pHead);
  }
  var includeS = Browser.inputBox("Do you want to include a status column?", "Yes or No", Browser.Buttons.OK_CANCEL);
  if(includeS.toLowerCase() == "yes"){
      var sHead = Browser.inputBox("What is the heading of your status column? e.g. Status", "", Browser.Buttons.OK_CANCEL);
       PropertiesService.getUserProperties().setProperty("status" + sheeturl, sHead);
  }
  var includeR = Browser.inputBox("Do you want to include a fix version column?", "Yes or No", Browser.Buttons.OK_CANCEL);
  if(includeR.toLowerCase() == "yes"){
      var rHead = Browser.inputBox("What is the heading of your fix version column? e.g. Release", "", Browser.Buttons.OK_CANCEL);
      PropertiesService.getUserProperties().setProperty("release" + sheeturl, rHead);
  }
  var includeD = Browser.inputBox("Do you want to include a type column?", "Yes or No", Browser.Buttons.OK_CANCEL);
  if(includeD.toLowerCase() == "yes"){
      var dHead = Browser.inputBox("What is the heading of your type column? e.g. Dev Category", "", Browser.Buttons.OK_CANCEL);
      PropertiesService.getUserProperties().setProperty("dev category" + sheeturl, dHead);
  }*/
}


function jiraPullManual() {
    if(PropertiesService.getUserProperties().getProperty("colHeads" + sheeturl)!=null){
       jiraHeads=PropertiesService.getUserProperties().getProperty("jiraHeads" + sheeturl).split(",");
       colHeads=PropertiesService.getUserProperties().getProperty("colHeads" + sheeturl).split(",");
    }
    if(PropertiesService.getUserProperties().getProperty("digest") == "No"){
        var userAndPassword = Browser.inputBox("Enter your Jira User id and Password in the form Username:Password. e.g. John.Doe:ilovejira", "Userid:Password", Browser.Buttons.OK_CANCEL);
        var x = Utilities.base64Encode(userAndPassword);
        PropertiesService.getUserProperties().setProperty("digest", "Basic " + x);
    }
    if(PropertiesService.getUserProperties().getProperty("colHeads" + sheeturl)==null){
        Browser.msgBox("Please configure the applet");
    }else{
        var start = new Date().getTime();
        var key=0;
        var jira = PropertiesService.getUserProperties().getProperty("jira"+ sheeturl);
        for(var a=0; a<jiraHeads.length; a++){
            for(var r=0; r<ss.getLastColumn(); r++){
              if(ss.getRange(1,r+1,1,r+1).getCell(1, 1).getValue().toLowerCase() == colHeads[a].toLowerCase()){
                colNums.push(r+1);
                r=ss.getLastColumn()-1;
              }
            }
        }
      for(var a=0; a<ss.getLastColumn(); a++){ 
        if(ss.getRange(1,a+1,1,a+1).getCell(1, 1).getValue().toLowerCase() == jira.toLowerCase()) key=a+1;
      }
        var inter = 2;
        var temp2 = ss.getRange(2, key, ss.getLastRow()+1, key).getValues();
        for(var i=0; i<ss.getLastRow(); i++){
            if(temp2[i][0].indexOf("FRED")>-1){
                inter=i+2;
            }
        }
        lRow = inter-1;
        vals = ss.getRange(2, key, lRow, key).getValues();
        vals2 = vals;
        var temp = new Array();
        for(var i=0; i<lRow; i++){
            if(vals[i][0] != ""){
                if(vals[i][0].indexOf(" ")>-1){
                    vals[i][0] = vals[i][0].replace(/\s/g, '');
                }
                if(vals[i][0].indexOf("jira.naehas.com")>-1){
                    vals[i][0] = vals[i][0].replace("https://jira.naehas.com/browse/", "");
                }
                if(vals[i][0].indexOf(",")>-1){
                    var t = vals[i][0].split(",");
                    for(var k=0; k<t.length-1; k++){
                        ss.insertRowAfter(i+2);
                        var temprange = ss.getRange(i+2, 1, i+2, ss.getLastColumn()).getValues();
                        for(var l=0; l<temprange[0].length; l++) ss.getRange(i+3+k, l+1, i+3+k, l+1).getCell(1,1).setValue(temprange[0][l]);
                    }
                    for(var k=0; k<t.length; k++){ 
                        var work = new Array();
                        work.push(t[k]);
                        for(var l=1; l<vals[0].length; l++) work.push("");
                        temp.push(work);
                    }
                    for(var k=0; k<t.length; k++) ss.getRange(i+2+k, key, i+2+k, key).getCell(1,1).setValue(t[k]);
                    var test = ss.getRange(i+2, key, i+2+t.length-1, key).getValues();
                }else{
                     temp.push(vals[i]);
                }
            }
        }
        var temp2 = new Array();
        for(var i=0; i<lRow; i++){
            if(vals2[i][0].indexOf(" ")>-1){
                vals2[i][0] = vals2[i][0].replace(/\s/g, '');
            }
            if(vals2[i][0].indexOf("jira.naehas.com")>-1){
                vals2[i][0] = vals2[i][0].replace("https://jira.naehas.com/browse/", "");
            }
            if(vals2[i][0].indexOf(",")>-1){
                var t = vals2[i][0].split(",");
                for(var k=0; k<t.length; k++){ 
                    var work = new Array();
                    work.push(t[k]);
                    for(var l=1; l<vals2[0].length; l++) work.push("");
                    temp2.push(work);
                }
            }else{
                temp2.push(vals2[i]);
            }
        }
        vals=temp;
        vals2=temp2;
        jiraPull();
        //ss.getRange(40,1,40,1).getCell(1,1).setValue(new Date().getTime() - start);
    }
}  

function jiraPull() {
    var data = getStories();
    if (data == "") {
        Browser.msgBox("Error pulling data from Jira - aborting now. \\n \\n Please check if any of your JIRA keys include: \\n • Misspellings \\n • A key that doesn't exist \\n \\n If you still recieve an error, please try re-entering your credentials.");
        return;
    }
    var headings = jiraHeads;
    var y = new Array();
    for (var i=0;i<vals2.length;i++) {
        if(vals2[i][0].indexOf("FRED")==-1){
            var temp = new Array();
            for(var a=0; a<colHeads.length; a++) temp.push("");
            y.push(temp);
        }else{
            var inter = data.issues[0].key;
            var o=0;
            while(inter != vals2[i][0]){
                 o++;
                 inter = data.issues[o].key;
            }
            var d=data.issues[o];
            var test = getStory(d,headings);
            y.push(getStory(d,headings));
        }
    } 
    var last = lRow;
    if (y.length>0) {
        for(var k=0; k<y.length; k++){
            var x=0;
            for(var a=0; a<colHeads.length; a++) ss.getRange(k+2, colNums[a], k+2, colNums[a]).getCell(1,1).setValue(y[k][a]);
        }
    }
}

function getStories() {
    var allData = {issues:[]};
    var data = {startAt:0,maxResults:0,total:1};
    var startAt = 0;
    while (data.startAt + data.maxResults < data.total) {
        Logger.log("Making request for %s entries", C_MAX_RESULTS);
        var inter = ["search?jql="];
        for (var i = 0; i < vals.length; i++) {
            if(i==vals.length-1) inter.push("issue%20%3D%20", vals[i][0], "%20order%20by%20rank%20&maxResults=", C_MAX_RESULTS, "&startAt=", startAt);
            else inter.push("issue%20%3D%20", vals[i][0], "%20or%20");
        }
        var interStr = inter.join("");
        var temp = getDataForAPI(interStr);
        if(temp == "") return temp;
        data =  JSON.parse(temp);  
        allData.issues = allData.issues.concat(data.issues);
        startAt = data.startAt + data.maxResults;
    }  
    return allData;
}  

function getDataForAPI(path) {
    var url = "https://" + PropertiesService.getUserProperties().getProperty("host") + "/rest/api/2/" + path;
    var digestfull = PropertiesService.getUserProperties().getProperty("digest");
    var headers = { "Accept":"application/json", "Content-Type":"application/json", "method": "GET", "headers": {"Authorization": digestfull}, "muteHttpExceptions": true};
    var resp = UrlFetchApp.fetch(url,headers );
    var test = resp.getResponseCode();
    if (resp.getResponseCode() != 200) {
        //Browser.msgBox("Error retrieving data for url" + url + ":" + resp.getContentText());
        return "";
    } else {
        return resp.getContentText();
    }  
} 

function getStory(data,headings) {
    var story = [];
    for (var i = 0;i<headings.length;i++) {
        if (headings[i] !== "" && getDataForHeading(data,headings[i].toLowerCase()) != null) {
            if(headings[i].toLowerCase() == "priority") story.push(getDataForHeading(data,headings[i].toLowerCase()).name + " - " + getDataForHeading(data,headings[i].toLowerCase()).id);
            else if(headings[i].toLowerCase() == "status") story.push(getDataForHeading(data,headings[i].toLowerCase()).name);
            else if(headings[i].toLowerCase() == "fixversions"){
                var inter = "";
                for(var l=0; l<getDataForHeading(data,"fixVersions").length; l++){
                    inter = inter+getDataForHeading(data,"fixVersions")[l].name;
                    if(l!=getDataForHeading(data,"fixVersions").length-1) inter = inter + ", ";
                }
                story.push(inter);
            }
            else if(headings[i].toLowerCase() == "issuetype"){
                story.push(getDataForHeading(data,headings[i].toLowerCase()).name);
            }
            else if(headings[i].toLowerCase() == "resolution"){
              story.push(getDataForHeading(data,headings[i].toLowerCase()).name);
            }
          else if(headings[i].toLowerCase() == "reporter"){
              story.push(getDataForHeading(data,headings[i].toLowerCase()).displayName);
            }
          else if(headings[i].toLowerCase() == "assignee"){
              story.push(getDataForHeading(data,headings[i].toLowerCase()).displayName);
            }
            else story.push(getDataForHeading(data,headings[i].toLowerCase()));
        }else if(headings[i] != ""){
            story.push("N/A");
        }
    }        
    return story;
}  

function getDataForHeading(data,heading) {
    if (data.hasOwnProperty(heading)) {
        return data[heading];
    }else if (data.fields.hasOwnProperty(heading)) {
        return data.fields[heading];
    }  
    var fieldName = heading;
    if (fieldName !== "") {
        if (data.hasOwnProperty(fieldName)) {
        return data[fieldName];
        } else if (data.fields.hasOwnProperty(fieldName)) {
            return data.fields[fieldName];
        }  
    }
    var splitName = heading.split(" ");
    if (splitName.length == 2) {
        if (data.fields.hasOwnProperty(splitName[0]) ) {
            if (data.fields[splitName[0]] && data.fields[splitName[0]].hasOwnProperty(splitName[1])) {
                return data.fields[splitName[0]][splitName[1]];
            }
            return "";
        }  
    }  
    return "N/A";
}
