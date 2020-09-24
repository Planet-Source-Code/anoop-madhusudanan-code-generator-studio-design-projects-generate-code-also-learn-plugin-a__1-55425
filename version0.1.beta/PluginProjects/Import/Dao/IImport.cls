VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IImport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Public Function PlugName()
'Called to get the name of plug in

PlugName = "Access 97 Database"

End Function



Public Function Import(Sch As Object)
'Function to initiate the Import

Dim FileName
Import = False

FileName = frmImport.GetFile()

If Trim(FileName) = "" Then
    Exit Function
End If

Dim Importer As New clsDAOImport
On Error Resume Next

Importer.DBToSchema CStr(FileName), Sch
If Not Err Then Import = False

Import = True

Set Importer = Nothing

End Function
