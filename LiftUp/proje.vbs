Set e3 = CreateObject("CT.Application")
Set prj = e3.CreateJobObject
Set sht = prj.CreateSheetObject
Set sym = prj.CreateSymbolObject
Set con = prj.CreateConnectionObject
Set dev = prj.CreateDeviceObject
Set pin1 = prj.CreatePinObject
Set pin2 = prj.CreatePinObject
Set pin = prj.CreatePinObject
Set uplist = CreateObject("Scripting.Dictionary")
Set lplist = CreateObject("Scripting.Dictionary")
Set con = prj.CreateConnectionObject

' Create the XML DOM object
Set xmlDoc = CreateObject("Msxml2.DOMDocument.6.0")

' Load the XML file
xmlDoc.Async = "False"
xmlDoc.Load("C:\Users\SEFA\Desktop\LiftUp-Electron-UI\LiftUp\src\components\Output.xml")

' Check to ensure the XML file was loaded successfully
If xmlDoc.ParseError.ErrorCode <> 0 Then
    WScript.Echo "XML loading error: " & xmlDoc.ParseError.Reason
    WScript.Quit -1
End If

' Get the list of row elements
Set rows = xmlDoc.getElementsByTagName("row")
row_count = rows.Length

' Initialize separate arrays for each attribute
Dim numberArray(), signalNameArray(), sourceConnectorArray(),sourcePinsArray(), destinationLocationArray()
ReDim numberArray(row_count - 1)
ReDim signalNameArray(row_count - 1)
ReDim sourceConnectorArray(row_count - 1)
ReDim sourcePinsArray(row_count - 1)
ReDim destinationLocationArray(row_count - 1)

' Iterate through each row element
For i = 0 To row_count - 1
    Set row = rows.item(i)
    ' Extract various details and store them in corresponding arrays
    numberArray(i) = row.selectSingleNode("Number").Text
    signalNameArray(i) = row.selectSingleNode("Signal/Name").Text
    sourceConnectorArray(i) = row.selectSingleNode("Source/Connector").Text
    sourcePinsArray(i) = row.selectSingleNode("Source/Pin_No").Text
    destinationLocationArray(i) = row.selectSingleNode("Destination/LOCATION").Text
Next

' Optional: Display contents of the arrays (for verification)
WScript.Echo "Numbers: " & Join(numberArray, ", ")
WScript.Echo "Signal Names: " & Join(signalNameArray, ", ")
WScript.Echo "Source Locations: " & Join(sourceConnectorArray, ", ")
WScript.Echo "Destination Locations: " & Join(destinationLocationArray, ", ")
WScript.Echo "Pin No: " & Join(sourcePinsArray, ", ")
WScript.Echo "Pin Counts: " & UBound(sourcePinsArray)+1


' Clean up
Set xmlDoc = Nothing





Dim xarr(2)
Dim yarr(2)

If prj.GetId = 0 Then
    prj.Create "Test"
End If
sht.create 0, "1", "DINA3", 0, 0





sym.Load "DEFBLOCK", "1"
sym.PlaceBlock Sht.GetId, 80, 200, 30,(UBound(sourcePinsArray)+1)*6
    
dev.Create "-X3", "", "", sourceConnectorArray(0), "", 0 
If (dev.IsConnector) Then
        pincnt = dev.GetAllPinIds(pinids) 
        WScript.Echo "Pin Count: " & pincnt
        For i = 1 To UBound(sourcePinsArray)+1
            sym.PlacePins pinids(i), "", 0, Sht.GetId, 110, 198 + i*6, 0
            WScript.Echo "Pins: " & pinids(i)
            uplist.Add i, sym.GetId
            
        Next
        
End If





sym.Load "DEFBLOCK", "1"
sym.PlaceBlock Sht.GetId, 150, 200, 30, (UBound(sourcePinsArray)+1)*6

dev.Create "-X4", "", "", "282108-1", "", 0 
If (dev.IsConnector) Then
        pincnt = dev.GetAllPinIds(pinids) 
        For i = 1 To UBound(sourcePinsArray)+1
            sym.PlacePins pinids(i), "", 0, Sht.GetId, 150, 198 + i*6, 0
            lplist.Add i, sym.GetId
        Next
        
End If




pincnt = UBound(sourcePinsArray)+1
For i = 0 To pincnt
    pin1.SetId uplist(i)
    pin2.SetId lplist(i)



    pin1.GetSchemaLocation xarr(1), yarr(1), grid
    pin2.GetSchemaLocation xarr(2), yarr(2), grid
    con.Create sht.GetId, 2, xarr, yarr


    'pin1Id = pin1.SetId(uplist(i))
    'pin2Id = pin2.SetId(lplist(i))
    'cableId = dev.SetId( cableIds( cableIndex ) )
    'pin.CreateWire "", "FLY", "2.0-BK", cableId, 0, 0
    'pin.SetEndPinId(1, pin1Id) 
    'pin.CreateWire "", "FLY", "2.0-BK", cableId, 0, 0
    'pin.SetEndPinId(2, pin2Id)


        
Next





sht.display