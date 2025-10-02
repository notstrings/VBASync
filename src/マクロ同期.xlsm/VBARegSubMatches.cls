VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VBARegSubMatches"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Description = "This class provides the means to access SubMatch (Group) values."
'@ModuleDescription "This class provides the means to access SubMatch (Group) values."
'@Exposed
Option Explicit

Private colSubMatchesVal As Collection
Private colSubMatchesKey As Collection

Friend Sub Add(ByVal Val As String, Optional ByVal key As String)
    If key <> vbNullString Then
        colSubMatchesVal.Add Val, key
        colSubMatchesKey.Add key
    Else
        colSubMatchesVal.Add Val
    End If
End Sub

'@Description("Returns the number of SubMatch objects found in the source string.")
Public Property Get Count() As Long
Attribute Count.VB_Description = "Returns the number of SubMatch objects found in the source string."
   Count = colSubMatchesVal.Count
End Property

'@Description("Returns the SubMatch string associated with the provided index. Named groups are allowed.")
'@DefaultMember
Public Property Get Item(ByVal pItemIndex As Variant) As String
Attribute Item.VB_Description = "Returns the SubMatch string associated with the provided index. Named groups are allowed."
Attribute Item.VB_UserMemId = 0
    If Not IsNumeric(pItemIndex) Then
        Dim elm
        For Each elm In colSubMatchesKey
            If elm = pItemIndex Then
                Item = colSubMatchesVal(pItemIndex)
            End If
        Next
    Else
        'VBS is zero-based, but Collection is 1-based
        If pItemIndex + 1 >= 1 And pItemIndex + 1 <= colSubMatchesVal.Count Then
            Item = colSubMatchesVal.Item(pItemIndex + 1) 'to match VBScript_RegExp_55 zero-based
        Else
            Err.Raise vbObjectError + 9, , "Requested item not found in SubMatches collection"
        End If
    End If
End Property

'@Enumerator
Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = colSubMatchesVal.[_NewEnum]
End Function

Private Sub Class_Terminate()
    Set colSubMatchesVal = Nothing
    Set colSubMatchesKey = Nothing
End Sub

Private Sub Class_Initialize()
    Set colSubMatchesVal = New Collection
    Set colSubMatchesKey = New Collection
End Sub
