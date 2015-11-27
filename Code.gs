// Script-as-app template.

//Trial to get to update

function doGet() {
  Logger.log(Session.getActiveUser().getEmail());
  return HtmlService.createHtmlOutputFromFile('index')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
}



//Process Time Card - Add Time
function processForm(formObject) {
  //TODO need to catch whether this is the end of something and add it to the calendar.
  var desc = formObject.Description;
  var workType = formObject.workType;
  var date = new Date();
  var timeStamp = date.toLocaleString();
  var location = formObject.location;
  var locationString = "=HYPERLINK(\"http://maps.google.com/?ll=" + location + "\",\""+ location + "\")";
  var sheet =  SpreadsheetApp.openById(findOrCreateTimeSheet());
  sheet.appendRow([timeStamp, desc, locationString, workType]);
  if(workType == "offWork"){
	  digest2();
  }
  return sheet.getUrl();
}

//Process Phone calls
function processForm2(formObject) {
  var cust = formObject.customer;
  var desc = formObject.description;
  var date = new Date();
  var timeStamp = date.toLocaleString();
  var sheet = SpreadsheetApp.openById(findOrCreatePhoneSheet());
  sheet.appendRow([timeStamp, cust, desc]);
  return sheet.getUrl();
}

//Returns a brief summary of the weeks events.  Meant to be viewed as web page.
function digest(){
	digest2();
	var currentWeekNo = new Date().getWeek();
	var currentWeekList = [];
	var currentWeekData = [];
	var mondayList = [];
	var tuesdayList = [];
	var wednesdayList = [];
	var thursdayList = [];
	var fridayList = [];
	var sheet =  SpreadsheetApp.openById(findOrCreateTimeSheet());
	var data = sheet.getDataRange().getValues(); // read all data in the sheet
	for(n=0;n<data.length;++n){//This will make a list of everything this week.
		var eqDate = new Date(data[n][0]);
		if(eqDate.getWeek() == currentWeekNo){
			currentWeekList.push(eqDate);
			currentWeekData.push(data[n]);
		}
	}
	for(n=0;n<currentWeekList.length;++n){//go through the list to sort to days...
		var today = new Date().getDay();
		var eqDate = new Date(currentWeekList[n]);
		if (eqDate.getDay() == 1){
			mondayList.push(currentWeekData[n]);
		}
		if (eqDate.getDay() == 2){
			tuesdayList.push(currentWeekData[n]);
		}
		if (eqDate.getDay() == 3){
			wednesdayList.push(currentWeekData[n]);
		}
		if (eqDate.getDay() == 4){
			thursdayList.push(currentWeekData[n]);
		}
		if (eqDate.getDay() == 5){
			fridayList.push(currentWeekData[n]);
		}
	}
	var mondayHours = findDayHours(mondayList);
	var tuesdayHours = findDayHours(tuesdayList);
	var wednesdayHours = findDayHours(wednesdayList);
	var thursdayHours = findDayHours(thursdayList);
	var fridayHours = findDayHours(fridayList);
	var soFar = parseFloat(mondayHours) + parseFloat(tuesdayHours) + 
	parseFloat(wednesdayHours) + parseFloat(thursdayHours) +
	parseFloat(fridayHours);
	var printString  = "Monday: " + convertToHHMM(mondayHours) + "<br>" +
	"Tuesday: " + convertToHHMM(tuesdayHours) + "<br>" +
	"Wednesday: " + convertToHHMM(wednesdayHours) + "<br>" +
	"Thursday: " + convertToHHMM(thursdayHours) + "<br>" +
	"Friday: " + convertToHHMM(fridayHours) + "<br><br>"+  
	"Total so far: " + convertToHHMM(soFar);
	return printString;

}
  
//Returns a longer summary of the weeks events.  Meant to be viewed as text.
//Also enters days activities into calendar.
function digest2(){
	var currentWeekNo = new Date().getWeek();
	var currentWeekList = [];
	var currentWeekData = [];
	var mondayList = [];
	var tuesdayList = [];
	var wednesdayList = [];
	var thursdayList = [];
	var fridayList = [];
	var sheet =  SpreadsheetApp.openById(findOrCreateTimeSheet());
	var data = sheet.getDataRange().getValues(); // read all data in the sheet
	for(n=0;n<data.length;++n){//This will make a list of everything this week.
		var eqDate = new Date(data[n][0]);
		if(eqDate.getWeek() == currentWeekNo){
			currentWeekList.push(eqDate);
			currentWeekData.push(data[n]);
		}
	}
	for(n=0;n<currentWeekList.length;++n){//go through the list to sort to days...
		var today = new Date().getDay();
		var eqDate = new Date(currentWeekList[n]);
		if (eqDate.getDay() == 1){
			mondayList.push(currentWeekData[n]);
		}
		if (eqDate.getDay() == 2){
			tuesdayList.push(currentWeekData[n]);
		}
		if (eqDate.getDay() == 3){
			wednesdayList.push(currentWeekData[n]);
		}
		if (eqDate.getDay() == 4){
			thursdayList.push(currentWeekData[n]);
		}
		if (eqDate.getDay() == 5){
			fridayList.push(currentWeekData[n]);
		}   
	}
	var mondayHours = findDayHours(mondayList);
	var tuesdayHours = findDayHours(tuesdayList);
	var wednesdayHours = findDayHours(wednesdayList);
	var thursdayHours = findDayHours(thursdayList);
	var fridayHours = findDayHours(fridayList);

	//Here is where we add...
	var mondayHoursList = findDayHoursList(mondayList);
	var tuesdayHoursList = findDayHoursList(tuesdayList);
	var wednesdayHoursList = findDayHoursList(wednesdayList);
	var thursdayHoursList = findDayHoursList(thursdayList);
	var fridayHoursList = findDayHoursList(fridayList);

	var soFar = parseFloat(mondayHours) + parseFloat(tuesdayHours) + 
	parseFloat(wednesdayHours) + parseFloat(thursdayHours) +
	parseFloat(fridayHours);
	var printString  = "Monday "+ mondayHoursList + " " + convertToHHMM(mondayHours) + " hours\n" +
	"Tuesday "+ tuesdayHoursList+ " " + convertToHHMM(tuesdayHours) + " hours\n" +
	"Wednesday "+ wednesdayHoursList+ " " + convertToHHMM(wednesdayHours) + " hours\n" +
	"Thursday "+ thursdayHoursList+ " " + convertToHHMM(thursdayHours) + " hours\n" +
	"Friday "+ fridayHoursList+ " " + convertToHHMM(fridayHours) + " hours\n\n"+  
	"Total: " + convertToHHMM(soFar);

	//Need to only send email if its saturday...
	if(today == 6){
		MailApp.sendEmail(Session.getActiveUser().getEmail(),
				"Week " + currentWeekNo,
				printString);
	}
}



//Gets the time worked of the list of that days activities.
function findDayHours(dayList){
  var startTime;
  var endTime;
  for(n=0; n<dayList.length;++n){

    //Logger.log("endTime: "+ endTime);
    if(dayList[n][3]=="atWork"){
      startTime = new Date(dayList[n][0]);
    }
    if(dayList[n][3]=="offWork"){
      endTime = new Date(dayList[n][0]); 
    }
  }
  if(startTime!=undefined && endTime == undefined ){
    endTime = new Date();
  }
  var timeDiff = Math.abs(endTime - startTime);
  if(isNaN(timeDiff) == true){
    timeDiff = 0;
  }
  var diffDays = timeDiff / (1000 * 3600); 
  return(diffDays.toFixed(2));
}


//Gets the time worked of the list of that days activities.  Returns everything printed out pretty-like.
function findDayHoursList(dayList){
  var numOfTasks = 0;
  var startTime;
  var endTime;
  var startTimeLoc;
  var endTimeLoc;
  var startService;
  var endService;
  var serviceLoc;
  var serviceDescrip;
  var descrip;
  var xtraDesc;
  var xtraServiceDesc;
  var today = false;
  for(n=0; n<dayList.length;++n){
    if(dayList[n][3]=="atWork"){
      startTime = new Date(dayList[n][0]);
      descrip = dayList[n][1];
      startTimeLoc = dayList[n][2];
      //Logger.log("startTime: " + startTime);
    }
    if(dayList[n][3]=="offWork"){
      endTime = new Date(dayList[n][0]); 
      xtraDesc = dayList[n][1];
      endTimeLoc = dayList[n][2];
      //Logger.log("endTime: " + endTime+","+xtraDesc);
    }
    //Working on more than day event!!!!
    
    if(dayList[n][3]=="startTask"){
        numOfTasks ++;
        //Logger.log("Number of Tasks " + numOfTasks);
    	startService = new Date(dayList[n][0]);
    	serviceDescrip = dayList[n][1];
        serviceLoc = dayList[n][2];
    }
    if(dayList[n][3]=="endTask"){
    	endService = new Date(dayList[n][0]);
    	xtraServiceDesc = dayList[n][1];
    	
    }
    //Trying to process each event
    var timeDiff = Math.abs(endTime - startTime);
    if(isNaN(timeDiff) == true){
        timeDiff = 0;
    }
    //Logger.log(startService+endService+serviceDescrip);
    
    if(today == false && startService != undefined && endService != undefined){
        addToCal(serviceDescrip, startService, endService, serviceLoc, "");
        //Horrible hack to allow events already added to not effect new events.  ie old end service, new start service
        endService = undefined;
    }
  //End test session
  }
  
  if(startTime!=undefined && endTime == undefined ){
    today = true;
    endTime = new Date();
    xtraDesc = "";
  }
  
  if(startService!=undefined && endService == undefined ){
	    endService = new Date();
	    xtraServiceDesc = "";
	  }
  
  var timeDiff = Math.abs(endTime - startTime);
  if(isNaN(timeDiff) == true){
    timeDiff = 0;
  }
  
    Logger.log("JUNK: "+descrip+","+startTime+","+endTime+","+startTimeLoc+","+xtraDesc)
  
  //add to calendar (if its not today's time, and there is a start and stop.)
  if(today == false && startTime != undefined && endTime != undefined){
      addToCal(descrip, startTime, endTime, startTimeLoc, xtraDesc);
      //Logger.log("Days time added.");
    
  }
  
  
  if(startTime == undefined || endTime == undefined){
    return "No data";
  }else{
    return startTime.toLocaleTimeString() + " to " +  endTime.toLocaleTimeString()+ " working at: "+ descrip;
  }
  
 }
  
//Helper function to convert date objects into time. "8:00"
function convertToHHMM(info) {
  var hrs = parseInt(Number(info));
  var min = Math.round((Number(info)-hrs) * 60);
  if(hrs < 10){
    hrs = hrs.toString();
  }
  else{
    hrs = hrs.toString();
  }
  if(min < 10){
    min = "0" + min.toString();
  }
  else{
    min = min.toString();
  }
  return hrs + ':' + min;
  
}

//Helper function to find or create spreadsheet.  Returns spreadsheet ID
function findOrCreateTimeSheet(){
	   var sheetThingList = DriveApp.getFilesByName("timeSheet");
	   if (sheetThingList.hasNext() == false){
	        var sheet = SpreadsheetApp.create("timeSheet");
	        sheet.appendRow(["Date / Time", "Description", "Locations", "Status"]);
	        var range = sheet.getRange("A1:D1");
	        range.setBackground("grey");
	        var sheetId = sheet.getId();
	   }else{
	       var driveFile = DriveApp.getFilesByName("timeSheet");
	     if (driveFile.hasNext()== true){
	       var fileID = driveFile.next().getId();
	       var sheet = SpreadsheetApp.openById(fileID);
	       var sheetId = sheet.getId();
	     }
	       
	   }
	   return sheetId;
}

//Helper function to create phone log spreadsheet
function findOrCreatePhoneSheet(){
	   var sheetThingList = DriveApp.getFilesByName("Phone Calls");
	   if (sheetThingList.hasNext() == false){
	        var sheet = SpreadsheetApp.create("Phone Calls");
	        sheet.appendRow(["Date / Time", "Customer", "Description"]);
	        var range = sheet.getRange("A1:D1");
	        range.setBackground("grey");
	        var sheetId = sheet.getId();
	   }else{
	       var driveFile = DriveApp.getFilesByName("Phone Calls");
	     if (driveFile.hasNext()== true){
	       var fileID = driveFile.next().getId();
	       var sheet = SpreadsheetApp.openById(fileID);
	       var sheetId = sheet.getId();
	     }
	       
	   }
	   return sheetId;
}


//Helper function to find calendar or create one.  Returns calendar ID
function findOrCreateCal(){
  var calList = CalendarApp.getCalendarsByName("Time Sheet");
  if (calList.length == 0){
        var cal = CalendarApp.createCalendar("Time Sheet");
        var calId = cal.getId();
   }else{
       var cal = CalendarApp.getCalendarsByName("Time Sheet");
       var calId = cal[0].getId();
   
     }
  return calId;
}

//Helper function that adds a calendar event
function addToCal(descrip, startTime, endTime, eventLoc, xtraInfo){
  Logger.log("Trying to add to Cal: "+ descrip+ "," + startTime+","+ endTime+"," +eventLoc+","+ xtraInfo);
   // Determines how many events are happening today.	
   //Sanitize special characters before searching...
   var desired = descrip.replace("\'", "");
   var today = new Date(startTime);
   var todayPlus = new Date(startTime + 1);
   var calId = findOrCreateCal();
   var cal = CalendarApp.getCalendarById(calId);
   //hmmm...  going to need to check for event name...
   // Determines how many events are happening today and contain the description.
   //var events = cal.getEventsForDay(today, {search: descrip});
   //******Found an issue here...  If the service and day title are the same, only one gets built...**************
   var events = cal.getEvents(today, todayPlus, {search: descrip});
   //var events = cal.getEventsForDay(today, {search: desired});
   //TODO  Catch the service info.. It doesn't make sense here...
   Logger.log("events.length: "+ events.length);
   if (events.length == 0){
	if (xtraInfo != ""){
		var event = cal.createEvent(desired, startTime, endTime,{location: eventLoc, description: "Started at: " + descrip + ". Ended at: " + xtraInfo + "."});
	}else{
      var event = cal.createEvent(desired, startTime, endTime,{location: eventLoc, description: "Task: " + descrip});
	}
    
   }
  

  
}

//Helper function to get the week number from date passed to it.  (index by sunday)  Returns week no
Date.prototype.getWeek = function() {
  //It looks like I need to force the day!!  Like onejan.getDay()+1
  var onejan = new Date(this.getFullYear(),0,1);
  var firstSunIndex = 6 - onejan.getDay();
  var sunjan =  new Date(onejan.getTime()+(firstSunIndex*24*60*60*1000));
  return Math.ceil((((this - sunjan) / 86400000) + onejan.getDay()+1)/7);
  
}
