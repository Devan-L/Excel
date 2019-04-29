Attribute VB_Name = "PensionElement"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private pElementDate As Date
Private pValue As Double
Private pElementID As String


Public Property Get ElementDate() As Date
    ElementDate = pElementDate
End Property

Public Property Let ElementDate(NewValue As Date)
    pElementDate = NewValue
End Property


Public Property Get Value() As Double
    Value = pValue
End Property

Public Property Let Value(NewValue As Double)
    pValue = NewValue
End Property


Public Property Get ElementID() As String
    ElementID = pElementID
End Property

Public Property Let ElementID(NewValue As String)
    pElementID = NewValue
End Property


Public Property Get ElementType() As String
    If Me.ElementID Like "*spouse*" Then
        ElementType = "Spouse"
    Else
        ElementType = "Member"
    End If
End Property

