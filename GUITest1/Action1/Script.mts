'importing data from excel file to data table
DataTable.ImportSheet "C:\Users\rugve\OneDrive\Documents\UFT One\data.xlsx", 1, "Global"
'storing number of rows in n ie 3
n = DataTable.GetSheet("Global").GetRowCount
'iterating 3 times by step 1
For i = 1 To n Step 1
'setting the current row to i
DataTable.SetCurrentRow(i) @@ hightlight id_;_1953829432_;_script infofile_;_ZIP::ssf71.xml_;_
'running the FlightGUI application
systemutil.Run "C:\Program Files (x86)\Micro Focus\UFT One\samples\Flights Application\FlightsGUI.exe"
'screenshot before login @@ hightlight id_;_1969648_;_script infofile_;_ZIP::ssf47.xml_;_
WpfWindow("Micro Focus MyFlight Sample").CaptureBitmap "C:\Users\rugve\OneDrive\Documents\UFT One\ssFlightGUI\preLogin"&i&".png" , True
'entering username and password	
WpfWindow("Micro Focus MyFlight Sample").WpfEdit("Name").Set DataTable.Value("Name","Global")
WpfWindow("Micro Focus MyFlight Sample").WpfEdit("password").SetSecure "61688c24743faa4b4ce6" @@ hightlight id_;_2029691848_;_script infofile_;_ZIP::ssf74.xml_;_
'screenshot after entering details
WpfWindow("Micro Focus MyFlight Sample").CaptureBitmap "C:\Users\rugve\OneDrive\Documents\UFT One\ssFlightGUI\postLogin"&i&".png" , True
'checkpoint on OK button to check if it is enabled and check the text
WpfWindow("Micro Focus MyFlight Sample").WpfButton("OK").Check CheckPoint("OK")
'clicking on OK button
WpfWindow("Micro Focus MyFlight Sample").WpfButton("OK").Click @@ hightlight id_;_3082080_;_script infofile_;_ZIP::ssf76.xml_;_
'screenshot of the flight itenary
WpfWindow("Micro Focus MyFlight Sample").CaptureBitmap "C:\Users\rugve\OneDrive\Documents\UFT One\ssFlightGUI\preFlightSelection"&i&".png" , True
'Entering the source, destination, date of flight, class, number of tickets
WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("fromCity").Select datatable("fromCity")
WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("toCity").Select datatable("toCity") @@ hightlight id_;_1927607936_;_script infofile_;_ZIP::ssf55.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfImage("WpfImage").Click 75,55 @@ hightlight id_;_1929419160_;_script infofile_;_ZIP::ssf56.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfCalendar("datePicker").SetDate datatable("date")
WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("Class").Select datatable("class") @@ hightlight id_;_1929413160_;_script infofile_;_ZIP::ssf59.xml_;_
WpfWindow("Micro Focus MyFlight Sample").WpfComboBox("numOfTickets").Select datatable("numOfTickets") @@ hightlight id_;_1929767472_;_script infofile_;_ZIP::ssf61.xml_;_
'clicking on find tickets
WpfWindow("Micro Focus MyFlight Sample").WpfButton("FIND FLIGHTS").Click @@ hightlight id_;_1928415344_;_script infofile_;_ZIP::ssf62.xml_;_
'checkpoint to check if 'SELECT FLIGHT' button is disabled before clicking on a flight
WpfWindow("Micro Focus MyFlight Sample").WpfButton("SELECT FLIGHT").Check CheckPoint("SELECT FLIGHT")
'selct a desired flight
WpfWindow("Micro Focus MyFlight Sample").WpfTable("flightsDataGrid").SelectCell 0,0 @@ hightlight id_;_1930436344_;_script infofile_;_ZIP::ssf63.xml_;_
'screenshot after selecting flight
WpfWindow("Micro Focus MyFlight Sample").CaptureBitmap "C:\Users\rugve\OneDrive\Documents\UFT One\ssFlightGUI\postFlightSelection"&i&".png" , True
'clicking on SELECT FLIGHT button
WpfWindow("Micro Focus MyFlight Sample").WpfButton("SELECT FLIGHT").Click @@ hightlight id_;_1984716712_;_script infofile_;_ZIP::ssf64.xml_;_
'Screenshot of selected itenary
WpfWindow("Micro Focus MyFlight Sample").CaptureBitmap "C:\Users\rugve\OneDrive\Documents\UFT One\ssFlightGUI\prePassengerDetail"&i&".png" , True
'checkpoint on ORDER button to check if it is enabled before entering the passenger name in the third iteration (failing checkpoint)
If i = 3 Then
		WpfWindow("Micro Focus MyFlight Sample").WpfButton("ORDER").Check CheckPoint("ORDER")
End If
'Entering the passenger name
WpfWindow("Micro Focus MyFlight Sample").WpfEdit("passengerName").Set datatable("passengerName") @@ hightlight id_;_1929413016_;_script infofile_;_ZIP::ssf65.xml_;_
'screenshot after entering passenger name
WpfWindow("Micro Focus MyFlight Sample").CaptureBitmap "C:\Users\rugve\OneDrive\Documents\UFT One\ssFlightGUI\postPassengerDetail"&i&".png" , True
'ckeckpoint on ORDER button to check if it is enabled after entering the passenger details
WpfWindow("Micro Focus MyFlight Sample").WpfButton("ORDER").Check CheckPoint("ORDER")
'clicking on the ORDER button
WpfWindow("Micro Focus MyFlight Sample").WpfButton("ORDER").Click
'Close the FlightGUI application
WpfWindow("Micro Focus MyFlight Sample").Close
Next
