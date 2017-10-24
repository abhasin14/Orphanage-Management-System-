<html>
	<head>
		<title > Cyanic Writers  </title>	
		<link rel="stylesheet" type="text/css" href="2-css/website.css">
		
	</head>
	<body>
	<iframe name=" my_frame" width="700" height="282" frameborder="0" scrolling="no" 
	src="https://onedrive.live.com/embed?cid=91DCFB34FF4037D3&resid=91DCFB34FF4037D3%211003&authkey=AOWX3qXDYU_tC2Q&em=2&Item='Sheet1'!A1%3AE13">
	</iframe>
	
	<div id="myExcelDiv" style="width: 700px; height: 302px"></div>
<script type="text/javascript" src="https://r.office.microsoft.com/r/rlidExcelWLJS?v=1&kip=1"></script>
<script type="text/javascript">
	/*
	 * This code uses the Microsoft Office Excel Javascript object model to programmatically insert the
	 * Excel Web App into a div with id=myExcelDiv. The full API is documented at
	 * https://msdn.microsoft.com/en-US/library/hh315812.aspx. There you can find out how to programmatically get
	 * values from your Excel file and how to use the rest of the object model. 
	 */

	// Use this file token to reference Book.xlsx in Excel's APIs
	var fileToken = "SD91DCFB34FF4037D3!1003/-7936192238294386733/t=0&s=0&v=!AOWX3qXDYU_tC2Q";

	// run the Excel load handler on page load
	if (window.attachEvent) {
		window.attachEvent("onload", loadEwaOnPageLoad);
	} else {
		window.addEventListener("DOMContentLoaded", loadEwaOnPageLoad, false);
	}

	function loadEwaOnPageLoad() {
		var props = {
			item: "'Sheet1'!A1:E13",
			uiOptions: {
				showDownloadButton: false,
				showParametersTaskPane: false
			},
			interactivityOptions: {
				allowTypingAndFormulaEntry: false,
				allowParameterModification: false,
				allowSorting: false,
				allowFiltering: false,
				allowPivotTableInteractivity: false
			}
		};

		Ewa.EwaControl.loadEwaAsync(fileToken, "myExcelDiv", props, onEwaLoaded);
	}

	function onEwaLoaded(result) {
		/*
		 * Add code here to interact with the embedded Excel web app.
		 * Find out more at https://msdn.microsoft.com/en-US/library/hh315812.aspx.
		 */
	}
</script>
	</body>
	