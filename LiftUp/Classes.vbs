Class signalRow
    Public Number
    Public SignalName
    Public SignalType
    Public SignalCategory
    Public CurrentMax

    Public CableType
    Public CableAwg

    Public SourceAta
    Public SourcePinName
    Public SourceLocation
    Public SourceLRU
    Public SourceRdNumber
    Public SourceConnector
    Public SourcePinNo

    Public DestinationAta
    Public DestinationPinName
    Public DestinationLocation
    Public DestinationLRU
    Public DestinationRdNumber
    Public DestinationConnector
    Public DestinationPinNo

    Public Notes
End Class

Class ConnectorPin
    Private pSourceBlock
    Private pSourcePinID
    Private pDestinationBlock
    Private pDestinationPinID

    ' Property Let for SourceBlock
    Public Property Set SourceBlock(obj)
        Set pSourceBlock = obj
    End Property

    ' Property Get for SourceBlock
    Public Property Get SourceBlock()
        Set SourceBlock = pSourceBlock
    End Property

    ' Property Let for SourcePinID
    Public Property Let SourcePinID(value)
        pSourcePinID = value
    End Property

    ' Property Get for SourcePinID
    Public Property Get SourcePinID()
        SourcePinID = pSourcePinID
    End Property

    ' Property Let for DestinationBlock
    Public Property Set DestinationBlock(obj)
        Set pDestinationBlock = obj
    End Property

    ' Property Get for DestinationBlock
    Public Property Get DestinationBlock()
        Set DestinationBlock = pDestinationBlock
    End Property

    ' Property Let for DestinationPinID
    Public Property Let DestinationPinID(value)
        pDestinationPinID = value
    End Property

    ' Property Get for DestinationPinID
    Public Property Get DestinationPinID()
        DestinationPinID = pDestinationPinID
    End Property
End Class

Class Block
    Public Name
    Public Pins

    Private Sub Class_Initialize()
        Set Pins = CreateObject("Scripting.Dictionary")
    End Sub

    Public Sub AddPin(pin)
        Pins.Add Pins.Count + 1, pin
    End Sub
End Class
