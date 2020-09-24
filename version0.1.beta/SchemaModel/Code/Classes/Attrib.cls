VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Attrib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit

'local variable(s) to hold property value(s)
Private mvarAttribName As Variant 'local copy
Private mvarAttribValue As Variant 'local copy
Public Property Let AttribValue(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AttribValue = 5
    mvarAttribValue = vData
End Property


Public Property Set AttribValue(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.AttribValue = Form1
    Set mvarAttribValue = vData
End Property


Public Property Get AttribValue() As Variant
Attribute AttribValue.VB_UserMemId = 0
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AttribValue
    If IsObject(mvarAttribValue) Then
        Set AttribValue = mvarAttribValue
    Else
        AttribValue = mvarAttribValue
    End If
End Property



Public Property Let AttribName(ByVal vData As Variant)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AttribName = 5
    mvarAttribName = vData
End Property


Public Property Set AttribName(ByVal vData As Variant)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.AttribName = Form1
    Set mvarAttribName = vData
End Property


Public Property Get AttribName() As Variant
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AttribName
    If IsObject(mvarAttribName) Then
        Set AttribName = mvarAttribName
    Else
        AttribName = mvarAttribName
    End If
End Property



