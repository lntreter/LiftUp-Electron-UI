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




pcnt = 1

Dim xarr(2)
Dim yarr(2)

If prj.GetId = 0 Then
    prj.Create "Test"
End If
sht.create 0, "1", "DINA3", 0, 0

sym.Load "DEFBLOCK", "1"
sym.PlaceBlock Sht.GetId, 80, 200, 30, 20
    
dev.Create "-X3", "", "", "282108-1", "CCCC", 0 
If (dev.IsConnector) Then
        pincnt = dev.GetAllPinIds(pinids) 
        For i = 0 To UBound(pinids)
            sym.PlacePins pinids(i), "W_BU", 0, Sht.GetId, 110, 198 + i*6, 0
            uplist.Add pcnt, sym.GetId
            pcnt = pcnt + 1
        Next
        
End If

pcnt = 1
sym.Load "DEFBLOCK", "1"
sym.PlaceBlock Sht.GetId, 150, 200, 30, 20

dev.Create "-X4", "", "", "282108-1", "CCCC", 0 
If (dev.IsConnector) Then
        pincnt = dev.GetAllPinIds(pinids) 
        For i = 0 To UBound(pinids)
            sym.PlacePins pinids(i), "W_BU", 0, Sht.GetId, 150, 198+ i*6, 0
            lplist.Add pcnt, sym.GetId
            pcnt = pcnt + 1
        Next
        
End If

For i = 0 To UBound(pinids)
    pin1.SetId uplist(i)
    pin2.SetId lplist(i)



    pin1.GetSchemaLocation xarr(1), yarr(1), grid
    pin2.GetSchemaLocation xarr(2), yarr(2), grid
    con.Create sht.GetId, 2, xarr, yarr

    pin1Id = pin1.SetId(uplist(i))
    pin2Id = pin2.SetId(lplist(i))

    'cableId = dev.SetId( cableIds( cableIndex ) )
    'pin.CreateWire "", "FLY", "2.0-BK", cableId, 0, 0
    'pin.SetEndPinId(1, pin1Id) 
    'pin.CreateWire "", "FLY", "2.0-BK", cableId, 0, 0
    'pin.SetEndPinId(2, pin2Id)


        
Next





sht.display
