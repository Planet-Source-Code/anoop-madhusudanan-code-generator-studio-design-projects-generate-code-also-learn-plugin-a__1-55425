VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Field"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"FieldAttributes"
Attribute VB_Ext_KEY = "Member1" ,"FieldHeaderAttributes"
Option Explicit

Private mvarFieldAttributes As FieldAttributes
Private mvarFieldHeaderAttributes As FieldHeaderAttributes


Public Property Get FieldHeaderAttributes() As FieldHeaderAttributes
    If mvarFieldHeaderAttributes Is Nothing Then
        Set mvarFieldHeaderAttributes = New FieldHeaderAttributes
    End If


    Set FieldHeaderAttributes = mvarFieldHeaderAttributes
End Property


Public Property Set FieldHeaderAttributes(vData As FieldHeaderAttributes)
    Set mvarFieldHeaderAttributes = vData
End Property



Public Property Get FieldAttributes() As FieldAttributes
    If mvarFieldAttributes Is Nothing Then
        Set mvarFieldAttributes = New FieldAttributes
    End If


    Set FieldAttributes = mvarFieldAttributes
End Property


Public Property Set FieldAttributes(vData As FieldAttributes)
    Set mvarFieldAttributes = vData
End Property
Private Sub Class_Terminate()
  Set mvarFieldHeaderAttributes = Nothing
    Set mvarFieldAttributes = Nothing
End Sub



