' Create the XML DOM object
Set xmlDoc = CreateObject("Msxml2.DOMDocument.6.0")

' Load the XML file
xmlDoc.Async = "False"
xmlDoc.Load("C:\Users\SEFA\Desktop\LiftUp-Electron-UI\LiftUp\src\components\Output.xml")

' Check to ensure the XML file was loaded successfully
If xmlDoc.ParseError.ErrorCode <> 0 Then
    WScript.Echo "XML loading error: " & xmlDoc.ParseError.Reason
Else
    ' Get the list of row elements
    Set rows = xmlDoc.getElementsByTagName("row")

    ' Iterate through each row element
    For Each row In rows
        ' Extract and display various details
        number = row.selectSingleNode("Number").Text
        signalName = row.selectSingleNode("Signal/Name").Text
        sourceLocation = row.selectSingleNode("Source/LOCATION").Text
        destinationLocation = row.selectSingleNode("Destination/LOCATION").Text

        WScript.Echo "Number: " & number
        WScript.Echo "Signal Name: " & signalName
        WScript.Echo "Source Location: " & sourceLocation
        WScript.Echo "Destination Location: " & destinationLocation
        WScript.Echo "--------------------------------"
    Next
End If

' Clean up
Set xmlDoc = Nothing
