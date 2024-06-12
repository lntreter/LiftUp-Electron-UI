' Main script
Set fso = CreateObject("Scripting.FileSystemObject")

' Read the contents of the SignalRow.vbs file
Set file = fso.OpenTextFile(fso.GetParentFolderName(WScript.ScriptFullName) & "\Classes.vbs", 1)
classDefinition = file.ReadAll
file.Close

' Execute the class definition globally
ExecuteGlobal classDefinition

Set e3 = CreateObject("CT.Application")
Set prj = e3.CreateJobObject
Set sht = prj.CreateSheetObject
Set sym = prj.CreateSymbolObject
Set con = prj.CreateConnectionObject
Set dev = prj.CreateDeviceObject
Set pin1 = prj.CreatePinObject
Set pin2 = prj.CreatePinObject
Set pin = prj.CreatePinObject
Set con = prj.CreateConnectionObject

' Create the XML DOM object
Set xmlDoc = CreateObject("Msxml2.DOMDocument.6.0")

' Load the XML file
xmlDoc.Async = "False"
xmlDoc.Load("LiftUp\src\components\Output.xml")

' Check to ensure the XML file was loaded successfully
If xmlDoc.ParseError.ErrorCode <> 0 Then
    WScript.Echo "XML loading error: " & xmlDoc.ParseError.Reason
    WScript.Quit -1
End If

' Get the list of row elements
Set rows = xmlDoc.getElementsByTagName("row")
row_count = rows.Length

Set hello = New signalRow

' Clean up
Set xmlDoc = Nothing

' Initialize the list of signalRow objects
Dim signalRows()
ReDim signalRows(row_count - 1)

WScript.Echo row_count
' Iterate through each row element
For i = 0 To row_count - 1

    Set row = rows.item(i)
    Set newSignalRow = New signalRow

    ' Extract various details and assign them to the signalRow properties
    newSignalRow.Number = row.selectSingleNode("Number").Text
    newSignalRow.SignalName = row.selectSingleNode("Signal/Name").Text
    newSignalRow.SignalType = row.selectSingleNode("Signal/TYPE").Text
    newSignalRow.SignalCategory = row.selectSingleNode("Signal/CATEGORY").Text
    newSignalRow.CurrentMax = row.selectSingleNode("Signal/CURRENT_Max").Text

    newSignalRow.CableType = row.selectSingleNode("CABLE/TYPE").Text
    newSignalRow.CableAwg = row.selectSingleNode("CABLE/AWG").Text

    newSignalRow.SourceAta = row.selectSingleNode("Source/ATA_CHAPTER").Text
    newSignalRow.SourcePinName = row.selectSingleNode("Source/PIN_NAME").Text
    newSignalRow.SourceLocation = row.selectSingleNode("Source/LOCATION").Text
    newSignalRow.SourceLRU = row.selectSingleNode("Source/LRU").Text
    newSignalRow.SourceRdNumber = row.selectSingleNode("Source/RD_NUMBER").Text
    newSignalRow.SourceConnector = row.selectSingleNode("Source/Connector").Text
    newSignalRow.SourcePinNo = row.selectSingleNode("Source/Pin_No").Text

    newSignalRow.DestinationAta = row.selectSingleNode("Destination/ATA_CHAPTER").Text
    newSignalRow.DestinationPinName = row.selectSingleNode("Destination/PIN_NAME").Text
    newSignalRow.DestinationLocation = row.selectSingleNode("Destination/LOCATION").Text
    newSignalRow.DestinationLRU = row.selectSingleNode("Destination/LRU").Text
    newSignalRow.DestinationRdNumber = row.selectSingleNode("Destination/RD_NUMBER").Text
    newSignalRow.DestinationConnector = row.selectSingleNode("Destination/Connector").Text
    newSignalRow.DestinationPinNo = row.selectSingleNode("Destination/Pin_No").Text

    ' Add the new signalRow object to the list
    Set signalRows(i) = newSignalRow

    ' Display the details of the new signalRow object
    'WScript.Echo "Number: " & signalRows(i).Number & ", Signal Name: " & signalRows(i).SignalName & ", Signal Type: " & signalRows(i).SignalType & ", Signal Category: " & signalRows(i).SignalCategory & ", Current Max: " & signalRows(i).CurrentMax & ", Cable Type: " & signalRows(i).CableType & ", Cable AWG: " & signalRows(i).CableAwg & ", Source ATA: " & signalRows(i).SourceAta & ", Source Pin Name: " & signalRows(i).SourcePinName & ", Source Location: " & signalRows(i).SourceLocation & ", Source LRU: " & signalRows(i).SourceLRU & ", Source RD Number: " & signalRows(i).SourceRdNumber & ", Source Connector: " & signalRows(i).SourceConnector & ", Source Pin No: " & signalRows(i).SourcePinNo & ", Destination ATA: " & signalRows(i).DestinationAta & ", Destination Pin Name: " & signalRows(i).DestinationPinName & ", Destination Location: " & signalRows(i).DestinationLocation & ", Destination LRU: " & signalRows(i).DestinationLRU & ", Destination RD Number: " & signalRows(i).DestinationRdNumber & ", Destination Connector: " & signalRows(i).DestinationConnector & ", Destination Pin No: " & signalRows(i).DestinationPinNo
Next

' Initialize a dictionary to store Block objects by name
Set blocks = CreateObject("Scripting.Dictionary")

' Initialize a dictionary to store pins by connector groups dynamically
Set connectorGroups = CreateObject("Scripting.Dictionary")

' Define a function to determine or create the group for the connector
Function GetOrCreateConnectorGroup(connectorName)
    If Not connectorGroups.Exists(connectorName) Then
        Set pinList = CreateObject("Scripting.Dictionary")
        connectorGroups.Add connectorName, pinList
    End If
    Set GetOrCreateConnectorGroup = connectorGroups(connectorName)
End Function

' Iterate through the signalRows to create Blocks and Pins
For i = 0 To UBound(signalRows)
    Set sr = signalRows(i)

    ' Create or get the SourceBlock
    If Not blocks.Exists(sr.SourceLRU) Then
        Set sourceBlock = New Block
        sourceBlock.Name = sr.SourceLRU
        blocks.Add sr.SourceLRU, sourceBlock
    Else
        Set sourceBlock = blocks(sr.SourceLRU)
    End If

    ' Create or get the DestinationBlock
    If Not blocks.Exists(sr.DestinationLRU) Then
        Set destinationBlock = New Block
        destinationBlock.Name = sr.DestinationLRU
        blocks.Add sr.DestinationLRU, destinationBlock
    Else
        Set destinationBlock = blocks(sr.DestinationLRU)
    End If

    ' Create a new Pin and set its properties
    Set cPin = New ConnectorPin
    Set cPin.SourceBlock = sourceBlock
    cPin.SourcePinID = sr.SourcePinNo
    Set cPin.DestinationBlock = destinationBlock
    cPin.DestinationPinID = sr.DestinationPinNo

    ' Add the Pin to the appropriate connector group list dynamically
    connectorName = sr.SourceConnector
    Set pinList = GetOrCreateConnectorGroup(connectorName)
    pinList.Add cPin.SourcePinID, cPin

    ' Add the Pin to the SourceBlock
    sourceBlock.AddPin cPin
    destinationBlock.AddPin cPin
Next

' Function to create connector and place pins
Sub CreateConnectorAndPlacePins(connectorName, pinList, x, y, pinIdList, unique)
    dev.Create connectorName, "", "", connectorName, "CCCC", 0 
    If dev.IsConnector Then
        pincnt = dev.GetAllPinIds(pinids)
        ' Use only the number of pins in pinList
        For i = 0 To pinList.Count
            sym.PlacePins pinids(i), "", 0, Sht.GetId, x, y + i * 8, 0
            pinIdList.Add unique, pinids(i) ' Add to pinIdList for connection
            unique = unique + 1
        Next
    End If
End Sub

' Initial coordinates
Dim xCoord, yCoord, unique
xCoord = 110
unique = 1

If prj.GetId = 0 Then
    prj.Create "Test"
End If
sht.create 0, "1", "DINA3", 0, 0

blocklocator = 1

' Initialize lists for pin connections
Set uplist = CreateObject("Scripting.Dictionary")
Set lplist = CreateObject("Scripting.Dictionary")

Dim xarr(2)
Dim yarr(2)

unique = 1
' Iterate through each block and place connectors and pins
For Each key In blocks.Keys
    yCoord = 150
    Set blk = blocks(key)
    WScript.Echo "Block Name: " & blk.Name
    ' Calculate the height of the block based on pin count
    blockHeight = blk.Pins.Count * 10

    customName = "MyCustomName" & blocklocator

    ' Place the block at the current coordinates
    sym.Load "DEFBLOCK", "1"
    sym.PlaceBlock Sht.GetId, blocklocator * 80, 150, 30, blockHeight
    sym.SetDeviceCompleteName blk.Name, "", ""

    ' Iterate through each connector group and place pins
    For Each connectorName In connectorGroups.Keys
        
        Set pinList = connectorGroups(connectorName)
        ' Adjust the coordinates for each connector group
        
        If blocklocator = 1 Then
            CreateConnectorAndPlacePins connectorName, pinList, xCoord, yCoord, uplist, unique
        ElseIf blocklocator = 2 Then
            CreateConnectorAndPlacePins connectorName, pinList, xCoord, yCoord, lplist, unique
        End If

        yCoord = yCoord + pinList.Count * 8
    Next

    ' Update the coordinates for the next block
    xCoord = xCoord + 50
    yCoord = yCoord + blockHeight + 50
    blocklocator = blocklocator + 1
Next

' Populate lplist from existing pins (example approach)
' Assuming lplist should be populated similar to uplist but for destination pins

' Print uplist and lplist elements
For Each key In uplist.Keys
    WScript.Echo "uplist: " & uplist(key)
Next

For Each key In lplist.Keys
    WScript.Echo "lplist: " & lplist(key)
Next

' Create connections between pins
For i = 0 To uplist.Count - 1
    WScript.Echo i
    pin1.SetId uplist.Items()(i)
    pin2.SetId lplist.Items()(i)

    pin1.GetSchemaLocation xarr(1), yarr(1), grid
    pin2.GetSchemaLocation xarr(2), yarr(2), grid
    con.Create sht.GetId, 2, xarr, yarr
Next

sht.display

' Display connector groups information for verification
For Each connectorName In connectorGroups.Keys
    WScript.Echo "Connector: " & connectorName
    Set pinList = connectorGroups(connectorName)
    For Each pinID In pinList.Keys
        Set pin = pinList(pinID)
        WScript.Echo "  Pin " & pinID & ": " & pin.SourceBlock.Name & " - " & pin.SourcePinID & " to " & pin.DestinationBlock.Name & " - " & pin.DestinationPinID
    Next
Next

