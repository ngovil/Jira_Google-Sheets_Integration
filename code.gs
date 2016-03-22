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
columns = new Array();
bool=true;
batchsize = 30;


function onInstall(e){
  onOpen(e); 
}

function onOpen(e){

  //var menu = SpreadsheetApp.getUi().createAddonMenu();
  var menu = SpreadsheetApp.getUi().createMenu("JIRA");
    menu.addItem("Configure Headings", "configureApp");
    menu.addItem("Credentials", "configurePassword");
    menu.addItem("Refresh Now", "jiraPullManual");
    menu.addToUi();
    jiraConfigure();
    findColHeads();
}

function findColHeads(){
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var cols = ss.getRange(1,1,1,ss.getLastColumn()).getValues();
  columns.length = 0;
  for(var i=0; i<cols[0].length; i++){
      if(cols[0][i].replace(" ", "") != "") columns.push(cols[0][i]);
  }
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

function configureApp(){
    var mydoc = SpreadsheetApp.getActiveSpreadsheet();
    var app = UiApp.createApplication().setTitle("What is the heading of your JIRA key column? e.g. JIRA").setHeight(75);
    var panel = app.createVerticalPanel().setId('panel');
    panel.add(app.createTextBox().setName('jira').setText('JIRA'));
    var handler = app.createServerHandler('submitJIRA').addCallbackElement(panel);
    panel.add(app.createButton('Submit', handler));
    app.add(panel);
    mydoc.show(app);
}

function submitJIRA(e){
  var label = e.parameter.jira;
  var sheeturl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  PropertiesService.getUserProperties().setProperty("jira" + sheeturl, label);
  var app = UiApp.getActiveApplication();
  app.close();
  doGet();
  return app;
}
function configurePassword(){
    var mydoc = SpreadsheetApp.getActiveSpreadsheet();
    var app = UiApp.createApplication().setTitle("Please enter your username and password: ").setHeight(120);
    var panel = app.createVerticalPanel().setId('panel');
    panel.add(app.createLabel("Username: "));
    panel.add(app.createTextBox().setName("username"));
    panel.add(app.createLabel("Password: "));
    panel.add(app.createPasswordTextBox().setName("password"));
    var handler = app.createServerHandler('submitPassword').addCallbackElement(panel);
    panel.add(app.createButton('Submit', handler));
    app.add(panel);
    mydoc.show(app);
}

function submitPassword(e){
  var username = e.parameter.username;
  var password = e.parameter.password;
  var x = Utilities.base64Encode(username + ":" + password);
  PropertiesService.getUserProperties().setProperty("digest", "Basic " + x);
  sleep(100);
  var app = UiApp.getActiveApplication();
  app.close();
  return app;
}

function doGet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    findColHeads();
    var sheeturl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
    var LIST = ["Priority", "Status", "fixVersions", "issueType", "Summary", "Description", "dueDate", "Resolution", "Reporter", "Assignee", "Labels", "Customers Impacted"];
    var mydoc = SpreadsheetApp.getActiveSpreadsheet();
    var app = UiApp.createApplication().setTitle("Jira Headings").setWidth(250).setHeight(420);
    var panel = app.createVerticalPanel().setId('panel');
    app.add(app.createLabel("Please check the JIRA headings you wish to include and type in the corresponding column headings."));
    app.add(app.createHTML("<br/>"));
    panel.add(app.createHidden('checkbox_total', LIST.length));
    for(var i = 0; i < LIST.length; i++){
      var checkbox = app.createCheckBox().setName('checkbox_isChecked_'+i).setText(LIST[i]);
      var hidden = app.createHidden('checkbox_value_'+i, LIST[i]);
      var dropDown = app.createListBox().setName('column_value_'+i);
      dropDown.addItem("");
      for(var a=0; a<columns.length; a++){
        var temp = columns[a];
         dropDown.addItem(temp);
      }
      panel.add(checkbox).add(hidden).add(dropDown);
    }
    var handler = app.createServerHandler('submit').addCallbackElement(panel);
    
    var scroll = app.createScrollPanel().setPixelSize(250, 300);
    scroll.add(panel);
    app.add(scroll);
    app.add(app.createHTML("<br/>"));
    app.add(app.createButton('Submit', handler));
    mydoc.show(app);
}

function submit(e){
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var sheeturl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
  var numberOfItems = e.parameter.checkbox_total;
  var itemsSelected = new Array();
  var columnsSelected = new Array();
  var app = UiApp.getActiveApplication();
  for(var i = 0; i < numberOfItems; i++){
    if(e.parameter['checkbox_isChecked_'+i] == 'true'){
      if(e.parameter['checkbox_value_'+i] == "Customers Impacted") itemsSelected.push("customfield_10024");
      else itemsSelected.push(e.parameter['checkbox_value_'+i]);
      if(e.parameter['column_value_'+i] != null && e.parameter['column_value_'+i] != "") columnsSelected.push(e.parameter['column_value_'+i]);
    }
  }
  if(columnsSelected.length == itemsSelected.length){
  app.close();
  jiraHeads = itemsSelected;
  colHeads = columnsSelected;

  PropertiesService.getUserProperties().setProperty("jiraHeads" + sheeturl, jiraHeads.toString());
  PropertiesService.getUserProperties().setProperty("colHeads" + sheeturl, colHeads.toString());
  return app;
  }else{
    Browser.msgBox("Please enter all the heading names.");
  }
}
function sleep(milliseconds) {
  var start = new Date().getTime();
  for (var i = 0; i < 1e7; i++) {
    if ((new Date().getTime() - start) > milliseconds){
      break;
    }
  }
}


function jiraPullManual() {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lRow=ss.getLastRow();
    var sheeturl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
    if(PropertiesService.getUserProperties().getProperty("colHeads" + sheeturl)!=null){
       jiraHeads=PropertiesService.getUserProperties().getProperty("jiraHeads" + sheeturl).split(",");
       colHeads=PropertiesService.getUserProperties().getProperty("colHeads" + sheeturl).split(",");
    }
    if(PropertiesService.getUserProperties().getProperty("digest") != "No"){
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
        var temp2 = new Array();
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
                    }
                    for(var k=0; k<t.length; k++){ 
                        var work = new Array();
                        work.push(t[k]);
                        for(var l=1; l<vals[0].length; l++) work.push("");
                        temp.push(work);
                        temp2.push(work);
                    }
                    for(var k=0; k<t.length; k++) ss.getRange(i+2+k, key, i+2+k, key).getCell(1,1).setValue(t[k]);
                }else{
                     temp.push(vals[i]);
                     temp2.push(vals[i]);
                }
            }else{
              temp2.push(vals[i]);
            }
        }
        vals=temp;
        vals2=temp2;
        jiraPull();
    }
    }else{
      Browser.msgBox("Please enter your credentials.");
    }
}  

function jiraPull() {
    var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var sheeturl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
    var y = new Array();
    var headings = jiraHeads;
    if(colHeads.length != colNums.length){
        Browser.msgBox("Error pulling data from Jira - aborting now. \\n \\n Please check if you have correctly named your headings.");
        return;
    }
   var temporary = new Array();
   for(var divideby = 0; divideby < Math.ceil(vals.length/batchsize); divideby++){
     var data = getStories(divideby);
     temporary.push(data);
     if (data == "") {
         return;
     }
  } 
  var issueData = temporary[0];
  for(var number = 1; number<temporary.length; number++){
     for(var i=0; i<temporary[number]['issues'].length; i++){
       issueData['issues'].push(temporary[number]['issues'][i]); 
    }
  }
   for (var i=0;i<vals2.length;i++) {
        if(vals2[i][0].indexOf("FRED")==-1){
            var temp = new Array();
            for(var a=0; a<colHeads.length; a++) temp.push("");
            y.push(temp);
        }else{
            var inter = issueData.issues[0].key;
            var o=0;
            while(inter != vals2[i][0]){
                 o++;
                 inter = issueData.issues[o].key;
            }
            var d=issueData.issues[o];
            var test = getStory(d,headings);
            y.push(getStory(d,headings));
        }
    }
    if (y.length>0) {
        for(var k=0; k<y.length; k++){
            var x=0;
            for(var a=0; a<colHeads.length; a++){
              if(y[k][a].length > 0) ss.getRange(k+2, colNums[a], k+2, colNums[a]).getCell(1,1).setValue(y[k][a]);
            }
        } 
    }
}

function getStories(divideby) {
    var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var sheeturl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
    var allData = {issues:[]};
    var data = {startAt:0,maxResults:0,total:1};
    var startAt = 0;
    Logger.log("Making request for %s entries", C_MAX_RESULTS);
    var inter = ["search?jql="];
    if(divideby*batchsize+batchsize>=vals.length) var numberofrepeats = vals.length;
    else numberofrepeats = divideby*batchsize + batchsize;
    for (var i = divideby*batchsize; i < numberofrepeats; i++) {
       if(i==numberofrepeats-1) inter.push("issue%20%3D%20", vals[i][0], "%20order%20by%20rank%20&maxResults=", C_MAX_RESULTS, "&startAt=", startAt);
       else inter.push("issue%20%3D%20", vals[i][0], "%20or%20");
    }
   
    var interStr = inter.join("");
    var temp = getDataForAPI(interStr);
    if(temp == "") return temp;
    data =  JSON.parse(temp);  
    allData.issues = allData.issues.concat(data.issues);
    startAt = data.startAt + data.maxResults;
    return allData;
}  

function getDataForAPI(path) {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var sheeturl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
    var url = "https://" + PropertiesService.getUserProperties().getProperty("host") + "/rest/api/2/" + path;
    var digestfull = PropertiesService.getUserProperties().getProperty("digest");
    var headers = { "Accept":"application/json", "Content-Type":"application/json", "method": "GET", "headers": {"Authorization": digestfull}, "muteHttpExceptions": true};
    var resp = UrlFetchApp.fetch(url,headers );
    var test = resp.getResponseCode();
    var test3 = resp.getContentText();
    
    if (resp.getResponseCode() != 200) {
        if(test3.indexOf("Unauthorized")>-1) Browser.msgBox("Please check your credentials.");
        else {
          var test2 = JSON.parse(resp.getContentText()).errorMessages;
          var inter = new Array();
          for(var i=0; i<test2.length; i++){
             inter.push(test2[i].split("'")[1]);
          }
          var inter2 = "";
          for(var i=0; i<test2.length; i++){
             inter2 = inter2 + "• " + inter[i] + "\\n";
          }
          Browser.msgBox("Please check the following key(s): \\n" + inter2);
        }
        return "";
    } else {
        return resp.getContentText();
    }  
} 

function getStory(data,headings) {
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var sheeturl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
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
            else if(headings[i].toLowerCase() == "labels"){
                var inter = "";
                for(var l=0; l<getDataForHeading(data,"labels").length; l++){
                    inter = inter+getDataForHeading(data,"labels")[l];
                    if(l!=getDataForHeading(data,"labels").length-1) inter = inter + ", ";
                }
                story.push(inter);
            }
            else if(headings[i].toLowerCase() == "customfield_10024"){
                var inter = "";
                for(var l=0; l<getDataForHeading(data,"customfield_10024").length; l++){
                    inter = inter+getDataForHeading(data,"customfield_10024")[l].value;
                    if(l!=getDataForHeading(data,"customfield_10024").length-1) inter = inter + ", ";
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
  var ss = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var sheeturl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
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
