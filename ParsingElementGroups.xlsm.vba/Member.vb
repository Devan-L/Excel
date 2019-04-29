Attribute VB_Name = "Member"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private pMemberNumber As String
Private pNationalInsuranceNumber As String
Private pElements As Collection


Public Property Get MemberNumber() As String
    MemberNumber = pMemberNumber
End Property

Public Property Let MemberNumber(NewValue As String)
    pMemberNumber = NewValue
End Property


Public Property Get NationalInsuranceNumber() As String
    NationalInsuranceNumber = pNationalInsuranceNumber
End Property

Public Property Let NationalInsuranceNumber(NewValue As String)
    pNationalInsuranceNumber = NewValue
End Property


Public Property Get Elements() As Collection
    Set Elements = pElements
End Property


Private Sub Class_Initialize()
    Set pElements = New Collection
End Sub
