VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IExport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit



Public Function PlugName()
'Called to get the name of plug in

PlugName = "Active Server Pages"

End Function

Public Function Export(Sch As Object, AppObject As Object)

Dim NewForm As New frmProject

Export = False

Load NewForm
Set NewForm.Sch = Sch
Set NewForm.AppObj = AppObject

NewForm.Show vbModal

If Not Err Then Export = True


End Function
