Attribute VB_Name = "modFun"
'====================================================================
'MODFUN
'====================================================================
'Functions for managing schema
'====================================================================
'By Anoop M. All Rights Reserved
'http://www.logicmatrixonline.com/anoop
'====================================================================
'License: This project comes with GNU GENERAL PUBLIC
'LICENSE. See License.htm


'====================================================================
Sub LoadPropb(MyFld As Field, PropBlock As String)
'====================================================================

'Loads the property from a string block to the field

Dim Props() As String, i, K
Dim Param, Value
On Error Resume Next

Props = Split(PropBlock, vbCrLf)

For i = LBound(Props()) To UBound(Props())
    Param = ""
    Value = ""
    
    GetValue Props(i), Param, Value
    
    For K = 0 To FldPropCount
        If VBA.LCase(FieldProperty(K)) = VBA.LCase(Param) Then
            MyFld.Properties(K) = Value
        End If
    Next
    
    
Next
    
End Sub


'====================================================================
Public Function GetValue(Str As String, Param, Value)
'====================================================================

  Dim i, Pos

    'Gets value from string. eg. font=arial. then returns arial
    On Error Resume Next
      Pos = InStr(1, Str, "=")
      If Pos = 0 Then
          Param = Str
          Value = ""
        Else
          Param = Trim$(Left$(Str, Pos - 1))
          Value = Trim$(Right$(Str, Len(Str) - Pos))
      End If

End Function



Sub InitArrays()
'Initialize the Arrays


FieldProperty(0) = "Attributes"
FieldProperty(1) = "Type"
FieldProperty(2) = "OrdinalPosition"
FieldProperty(3) = "Size"
FieldProperty(4) = "DefaultValue"
FieldProperty(5) = "ValidationRule"
FieldProperty(6) = "ValidationText"
FieldProperty(7) = "Required"
FieldProperty(8) = "AllowZeroLength"
FieldProperty(9) = "DecimalPlaces"
FieldProperty(10) = "Format"
FieldProperty(11) = "Caption"
FieldProperty(12) = "Min"
FieldProperty(13) = "Max"
FieldProperty(14) = "Lookup"
FieldProperty(15) = "LookupQuery"

FldPropCount = 16

Dim i, K, L

Erase SH.Entities
Erase SH.Relations

ReDim SH.Entities(MAXENT)
ReDim SH.Relations(MAXREL)


For i = 0 To MAXENT
    
    ReDim SH.Entities(i).Fields(MAXFLD)
    
    For K = 0 To MAXFLD
        Erase SH.Entities(i).Fields(K).Properties
        ReDim SH.Entities(i).Fields(K).Properties(20)
    Next K
Next i

SH.EntCount = 0
SH.RelCount = 0

End Sub

