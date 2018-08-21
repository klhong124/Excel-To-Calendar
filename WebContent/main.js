var appointments = [];

function ExportToTimeTable() {


// prepare the data
var source =
{
    dataType: "array",
    dataFields: [
        { name: 'id', type: 'string' },
        { name: 'description', type: 'string' },
        { name: 'location', type: 'string' },
        { name: 'subject', type: 'string' },
        { name: 'calendar', type: 'string' },
        { name: 'start', type: 'date' },
        { name: 'end', type: 'date' }
    ],
    id: 'id',
    localData: appointments
};
var adapter = new $.jqx.dataAdapter(source);
$("#scheduler").jqxScheduler({
    date: new $.jqx.date(2018, 8, 3),
    width: 1450,
    height: 850,
    source: adapter,
    view: 'monthView',
    showLegend: true,
    ready: function () {
        $("#scheduler").jqxScheduler('ensureAppointmentVisible', 'id1');
    },
    resources:
    {
        colorScheme: "scheme05",
        dataField: "calendar",
        source:  new $.jqx.dataAdapter(source)
    },
    appointmentDataFields:
    {
        from: "start",
        to: "end",
        id: "id",
        description: "description",
        location: "place",
        subject: "subject",
        resourceId: "calendar"
    },
    views:
    [
        'dayView',
        'weekView',
        'monthView',
    ]
});
};

function ExportToTable() {  
     var regex = /^([a-zA-Z0-9\s_\\.\-:])+(.xlsx|.xls)$/; 
     /*Checks whether the file is a valid excel file*/  
     if (regex.test($("#excelfile").val().toLowerCase())) {  
         var xlsxflag = false; 
         /*Flag for checking whether excel is .xls format or .xlsx format*/  
         if ($("#excelfile").val().toLowerCase().indexOf(".xlsx") > 0) {  
             xlsxflag = true;  
         }  
         /*Checks whether the browser supports HTML5*/  
         if (typeof (FileReader) != "undefined") {
             var reader = new FileReader();  
             reader.onload = function (e) {
                 var data = e.target.result;  
                 /*Converts the excel data in to object*/  
                 if (xlsxflag) {  
                     var workbook = XLSX.read(data, { type: 'binary' });  
                 }  
                 else {  
                     var workbook = XLS.read(data, { type: 'binary' });  
                 }  
                 /*Gets all the sheetnames of excel in to a variable*/  
                 var sheet_name_list = workbook.SheetNames;  
                 var cnt = 0; /*This is used for restricting the script to consider only first sheet of excel*/  
                 sheet_name_list.forEach(function (y) { 
                	 /*Iterate through all sheets*/  
                     /*Convert the cell value to Json*/  
                     if (xlsxflag) {  
                         var exceljson = XLSX.utils.sheet_to_json(workbook.Sheets[y]);  
                     }  
                     else {  
                         var exceljson = XLS.utils.sheet_to_row_object_array(workbook.Sheets[y]);  
                     }  
                     if (exceljson.length > 0 && cnt == 0) {  
                         BindTable(exceljson);
                         cnt++;  
                     }
                 });
               ExportToTimeTable();
             }
             if (xlsxflag) {/*If excel file is .xlsx extension than creates a Array Buffer from excel*/  
                 reader.readAsArrayBuffer($("#excelfile")[0].files[0]);  
             }  
             else {  
                 reader.readAsBinaryString($("#excelfile")[0].files[0]);  
             }  
         }  
         else {  
             alert("Sorry! Your browser does not support HTML5!");  
         }  
     }  
     else {  
         alert("Please upload a valid Excel file!");  
     }  
 }  

function BindTable(jsondata) {/*Function used to convert the JSON array to Html Table*/  
    var columns = BindTableHeader(jsondata); /*Gets all the column headings of Excel*/ 
    for (var i = 0; i < jsondata.length; i++) {
    	var data = [];
        for (var colIndex = 0; colIndex < columns.length; colIndex++) {  
            var cellValue = jsondata[i][columns[colIndex]];  
            if (cellValue == null)  
                cellValue = "";  
            data.push(cellValue);
        }  
        /*"0 StartDate", "1 EndDate", "2 Subject", "3 Location", "4 ResourceID", "5 Attendance", "6 StaffId", "7 StaffDetail",
         *  "8 PurposeValue", "9 MeetingNature", "10 FoodDrinks", "11 Equipment", "12 CompanyName*/ 
        var dataSet = {
        	    id: data[6],
        	    description: data[7],
        	    location: data[3],
        	    subject: data[2],
        	    calendar: data[4],
        	    start: new Date(data[0].substring(6, 10),
        	    				data[0].substring(3, 5),
        	    				data[0].substring(0, 2),
        	    				data[0].substring(11, 13),
        	    				data[0].substring(14, 16),
        	    				data[0].substring(17, 19)),
        	    end: new Date(data[1].substring(6, 10),
			    			  data[1].substring(3, 5),
			    			  data[1].substring(0, 2),
				    		  data[1].substring(11, 13),
				    		  data[1].substring(14, 16),
				    		  data[1].substring(17, 19))
        }
        appointments.push(dataSet);
    } 
    return dataSet;
}  
function BindTableHeader(jsondata) {/*Function used to get all column names from JSON and bind the html table header*/  
    var columnSet = []; 
    for (var i = 0; i < jsondata.length; i++) {
        var rowHash = jsondata[i];  
        for (var key in rowHash) {
            if (rowHash.hasOwnProperty(key)) {  
                if ($.inArray(key, columnSet) == -1) {/*Adding each unique column names to a variable array*/  
                    columnSet.push(key); 
                }  
            }  
        }  
    }
    return columnSet;  
}  