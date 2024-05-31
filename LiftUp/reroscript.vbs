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
Set uplist = CreateObject("Scripting.Dictionary")
Set lplist = CreateObject("Scripting.Dictionary")
Set con = prj.CreateConnectionObject

' Create the XML DOM object
Set xmlDoc = CreateObject("Msxml2.DOMDocument.6.0")

' Load the XML file
xmlDoc.Async = "False"
xmlDoc.Load("C:\Users\Cemil\Documents\Code\LiftUp-Electron-UI\LiftUp\src\components\Output.xml")

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
    WScript.Echo "Number: " & signalRows(i).Number & ", Signal Name: " & signalRows(i).SignalName & ", Signal Type: " & signalRows(i).SignalType & ", Signal Category: " & signalRows(i).SignalCategory & ", Current Max: " & signalRows(i).CurrentMax & ", Cable Type: " & signalRows(i).CableType & ", Cable AWG: " & signalRows(i).CableAwg & ", Source ATA: " & signalRows(i).SourceAta & ", Source Pin Name: " & signalRows(i).SourcePinName & ", Source Location: " & signalRows(i).SourceLocation & ", Source LRU: " & signalRows(i).SourceLRU & ", Source RD Number: " & signalRows(i).SourceRdNumber & ", Source Connector: " & signalRows(i).SourceConnector & ", Source Pin No: " & signalRows(i).SourcePinNo & ", Destination ATA: " & signalRows(i).DestinationAta & ", Destination Pin Name: " & signalRows(i).DestinationPinName & ", Destination Location: " & signalRows(i).DestinationLocation & ", Destination LRU: " & signalRows(i).DestinationLRU & ", Destination RD Number: " & signalRows(i).DestinationRdNumber & ", Destination Connector: " & signalRows(i).DestinationConnector & ", Destination Pin No: " & signalRows(i).DestinationPinNo
Next

' Initialize a dictionary to store Block objects by name
Set blocks = CreateObject("Scripting.Dictionary")

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
    ' Add the Pin to the SourceBlock
    sourceBlock.AddPin cPin
    destinationBlock.AddPin cPin
Next

' Display blocks and pins information for verification
For Each key In blocks.Keys
    Set blk = blocks(key)
    WScript.Echo "Block Name: " & blk.Name
    For Each pinKey In blk.Pins.Keys
        Set pin = blk.Pins(pinKey)
        WScript.Echo "  Pin " & pinKey & ": " & pin.SourceBlock.Name & " - " & pin.SourcePinID & " to " & pin.DestinationBlock.Name & " - " & pin.DestinationPinID
    Next
Next

pcnt = 1

Dim xarr(2)
Dim yarr(2)

If prj.GetId = 0 Then
    prj.Create "Test"
End If
sht.create 0, "1", "DINA3", 0, 0

' Initialize a variable to track the vertical position of the blocks
Dim verticalPosition
verticalPosition = 200 ' Initial vertical position

' Iterate through the blocks to position them and calculate height
blocklocator = 1



For Each key In blocks.Keys
    Set blk = blocks(key)
    ' Calculate the height of the block based on pin count
    blockHeight = blk.Pins.Count * 10

    customName = "MyCustomName" & blocklocator

    WScript.Echo blk.Name
    ' Place the block at the current vertical position
    sym.Load "DEFBLOCK", "1"
    sym.PlaceBlock Sht.GetId, blocklocator * 80, 150, 30, blockHeight
    sym.SetDeviceCompleteName blk.Name, "", ""
    
    ' Update the vertical position for the next block
    verticalPosition = verticalPosition + blockHeight + 50
    blocklocator = blocklocator + 1
Next

sht.display