Attribute VB_Name = "modPlugins"
Option Explicit

'====================================================================
'Plugins
'====================================================================
'Smart Code Engine Plugin Interface
'====================================================================

'====================================================================
Function LoadImportPlugins()
'Function to load all input plugins
'====================================================================

  Dim s As String, sPlugin As String, ObjTemp As Object, sCaption As String
  Dim IfBreak As Boolean
  IfBreak = False

    s = Dir$(App.Path & "\Import\*.*")
    
    Do Until LenB(s) = 0
        
            If Right(s, 4) = ".dll" Or Right(s, 4) = ".map" Then
              
              If Right(s, 4) = ".dll" Or Right(s, 4) = ".map" Then
                sPlugin = Mid(s, 1, Len(s) - 4) & ".IImport"
              End If
              
              On Error GoTo NoObj
              Set ObjTemp = CreateObject(sPlugin)
              sCaption = ObjTemp.PlugName()
              
              If Trim(sCaption) <> "" Then
                AddInMenu sCaption, sPlugin
              End If
              
NoObj:
              Set ObjTemp = Nothing
            End If
        
        
        s = Dir
    Loop
    
    If frmMain.mnuImportSub(frmMain.mnuImportSub.Count - 1).Caption = "-" Then
        frmMain.mnuImportSub(frmMain.mnuImportSub.Count - 1).Visible = False
    End If


'====================================================================
End Function
'====================================================================


Public Function AddInMenu(PluginCaption As String, PluginTag As String) As Integer
'Function to add input plugins to the menu


On Error Resume Next
Dim iIndex As Integer

iIndex = (frmMain.mnuImportSub.Count - 1) ' Get the position (Index) of where the mnuImportSub must go.
If frmMain.mnuImportSub(0).Enabled = True Then iIndex = iIndex + 1
With frmMain
  If iIndex <> 0 Then Load .mnuImportSub(iIndex)
  .mnuImportSub(iIndex).Caption = PluginCaption
  .mnuImportSub(iIndex).Visible = True
  .mnuImportSub(iIndex).Enabled = True
  .mnuImportSub(iIndex).Tag = PluginTag
End With

End Function


'====================================================================
Function LoadExportPlugins()
'Function to load all input plugins
'====================================================================

  Dim s As String, sPlugin As String, ObjTemp As Object, sCaption As String
  Dim IfBreak As Boolean
  IfBreak = False

    s = Dir$(App.Path & "\Export\*.*")
    
    Do Until LenB(s) = 0
        
            If Right(s, 4) = ".dll" Or Right(s, 4) = ".map" Then
              
              If Right(s, 4) = ".dll" Or Right(s, 4) = ".map" Then
                sPlugin = Mid(s, 1, Len(s) - 4) & ".IExport"
              End If
              
              On Error GoTo NoObj
              Set ObjTemp = CreateObject(sPlugin)
              sCaption = ObjTemp.PlugName()
              
              If Trim(sCaption) <> "" Then
                AddOutMenu sCaption, sPlugin
              End If
              
NoObj:
              Set ObjTemp = Nothing
            End If
        
        
        s = Dir
    Loop
    
    If frmMain.mnuExportsub(frmMain.mnuExportsub.Count - 1).Caption = "-" Then
        frmMain.mnuExportsub(frmMain.mnuExportsub.Count - 1).Visible = False
    End If


'====================================================================
End Function
'====================================================================


Public Function AddOutMenu(PluginCaption As String, PluginTag As String) As Integer
'Function to add input plugins to the menu


On Error Resume Next
Dim iIndex As Integer

iIndex = (frmMain.mnuExportsub.Count - 1) ' Get the position (Index) of where the mnuExportSub must go.
If frmMain.mnuExportsub(0).Enabled = True Then iIndex = iIndex + 1
With frmMain
  If iIndex <> 0 Then Load .mnuExportsub(iIndex)
  .mnuExportsub(iIndex).Caption = PluginCaption
  .mnuExportsub(iIndex).Visible = True
  .mnuExportsub(iIndex).Enabled = True
  .mnuExportsub(iIndex).Tag = PluginTag
End With

End Function
