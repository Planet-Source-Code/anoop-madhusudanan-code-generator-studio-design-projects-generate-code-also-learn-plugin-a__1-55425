VERSION 5.00
Object = "*\A..\..\..\SchemaControl\Schema.vbp"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "Entity Studio"
   ClientHeight    =   5190
   ClientLeft      =   930
   ClientTop       =   1110
   ClientWidth     =   5460
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5190
   ScaleWidth      =   5460
   WindowState     =   2  'Maximized
   Begin SchemaControl.Schema Schema1 
      Height          =   4455
      Left            =   -30
      TabIndex        =   1
      Top             =   630
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   7858
   End
   Begin MSComctlLib.ImageList imgMain 
      Left            =   1365
      Top             =   2625
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":09DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0F76
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1510
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1AAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1C04
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":219E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2738
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2892
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":29EC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tbMain 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   1005
      ButtonWidth     =   1931
      ButtonHeight    =   953
      Wrappable       =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "imgMain"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Import From"
            Key             =   "import"
            Description     =   "Import Schema"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Open Schema"
            Key             =   "open"
            Description     =   "Open Schema From XML File"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Save Schema"
            Key             =   "save"
            Description     =   "Save Schema To XML File"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Export To"
            Key             =   "generate"
            Description     =   "Generate Code"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Get Help"
            Key             =   "help"
            Description     =   "Get Help"
            ImageIndex      =   7
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   2190
      Top             =   5730
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuImport 
         Caption         =   "Import &From"
         Begin VB.Menu mnuImportSub 
            Caption         =   "[No plugins]"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu mnuFB2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExport 
         Caption         =   "Export & To"
         Begin VB.Menu mnuExportsub 
            Caption         =   "[No plugins]"
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu mnuFB1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSchema 
         Caption         =   "&Open Schema"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuClose 
         Caption         =   "&Close Schema"
      End
      Begin VB.Menu mnuFB3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save Schema"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "&Save Schema As.."
      End
      Begin VB.Menu mnuFB4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "&Edit"
      Begin VB.Menu mnuFieldProp 
         Caption         =   "&Field Properties"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuArrange 
         Caption         =   "&Arrange Now"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "&View"
      Begin VB.Menu mnuEntity 
         Caption         =   "&Entity"
         Begin VB.Menu mnuEntMin 
            Caption         =   "&Minimize All Entities"
         End
         Begin VB.Menu mnuEntRes 
            Caption         =   "&Restore All Entities"
         End
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpContents 
         Caption         =   "&Contents"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuWelcome 
         Caption         =   "&Show Welcome Dialog"
      End
      Begin VB.Menu mnuHb1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuLic 
         Caption         =   "&GNU General Public License"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "&About Entity Studio"
         Shortcut        =   ^{F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'====================================================================
'Main form
'====================================================================
'By Anoop M. All Rights Reserved
'http://www.logicmatrixonline.com/anoop
'====================================================================

'License: This project comes with GNU GENERAL PUBLIC
'LICENSE. See License.htm



Public CurFile As String, IsDirty As Boolean
Private Sub Form_Load()
CurFile = ""
Me.Caption = "Entity Studio"
IsDirty = False

modPlugins.LoadExportPlugins
modPlugins.LoadImportPlugins

On Error Resume Next
frmTip.Show , Me


End Sub

Private Sub Form_Resize()
'Resize
On Error Resume Next
Schema1.Move 0, Schema1.Top, Me.ScaleWidth, Me.ScaleHeight - Schema1.Top
Schema1.Arrange

End Sub


Function AskSave()
'Asks whether or not to save the file

Dim Ret

If IsDirty = False Then
    AskSave = vbNo
    Exit Function
End If

Ret = MsgBox("Do you want to save changes to the schema?", vbYesNoCancel + vbQuestion)

If Ret = vbCancel Then
AskSave = vbCancel
Exit Function
End If

If Ret = vbYes Then
    mnuSave_Click
    AskSave = vbYes
Else
    AskSave = vbNo
End If


End Function

Private Sub Form_Unload(Cancel As Integer)

If AskSave() <> vbCancel Then
    CurFile = ""
   Me.Caption = "Entity Studio - Unsaved"
    Schema1.Flush
Else
    Cancel = 1
End If

End Sub

Private Sub mnuAbout_Click()
frmHelp.LoadDoc App.Path & "\docs\about.htm"

End Sub

Private Sub mnuArrange_Click()
    
    Schema1.Arranged = False
    Schema1.Arrange

End Sub

Private Sub mnuClose_Click()
If CurFile = "" Then Exit Sub

If AskSave() <> vbCancel Then
    CurFile = ""
   Me.Caption = "Entity Studio - Unsaved"
    Schema1.Flush
End If


End Sub


Private Sub mnuEntities_Click()
'Schema1.EditEntity
End Sub

Private Sub mnuEntMin_Click()
Schema1.Restore True

End Sub

Private Sub mnuEntRes_Click()
    Schema1.Restore False
End Sub

Private Sub mnuExportsub_Click(Index As Integer)

On Error Resume Next

Dim s As String, sResult

s = mnuExportsub(Index).Tag

              On Error GoTo NoObj
              Set ObjTemp = CreateObject(s)
              'Load the schema to our object
              sResult = ObjTemp.Export(Schema1.GetSchema(), App)
              
                If sResult = False Then
                    MsgBox "Unable to export schema using the plugin", vbInformation + vbOKOnly, "Cannot Import"
                End If
              
NoObj:
                    
                    Set ObjTemp = Nothing
                




End Sub

Private Sub mnuFieldProp_Click()
On Error Resume Next
    Schema1.EditField
End Sub

Private Sub mnuHelpContents_Click()
    frmHelp.LoadDoc App.Path & "\docs\readme.htm"
End Sub



Private Sub mnuRelations_Click()
'
End Sub


Private Sub mnuImportSub_Click(Index As Integer)
On Error Resume Next
If AskSave() = vbCancel Then Exit Sub

Dim s As String, sResult
Dim Sch As New SchemaModel.Schema


s = mnuImportSub(Index).Tag

              On Error GoTo NoObj
              Set ObjTemp = CreateObject(s)
              'Load the schema to our object
              sResult = ObjTemp.Import(Sch)
              
                If sResult = False Then
                    MsgBox "Unable to import schema using the plugin", vbInformation + vbOKOnly, "Cannot Import"
                Else
                    'Load schema to our control
                        Schema1.LoadSchema Sch
                    'Arrange the control's entities automatically
                        mnuArrange_Click
                        IsDirty = False
                        Me.Caption = "Entity Studio - New Schema"
                End If
              
NoObj:
                    Set ObjTemp = Nothing
                



End Sub

Private Sub mnuLic_Click()
frmHelp.LoadDoc App.Path & "\docs\license.htm"

End Sub

Private Sub mnuQuickStart_Click()
frmHelp.LoadDoc App.Path & "\docs\Readme.htm#B. Quick Start"

End Sub

Private Sub mnuSave_Click()
If CurFile = "" Then
    mnuSaveAs_Click
    Exit Sub
End If

Dim RetStr
RetStr = Schema1.DumpSchema()

On Error Resume Next
Open CurFile For Output As #1

If Err Then
    MsgBox "File error. Unable to save changes. Try saving the file with a different name", vbCritical
    mnuSaveAs_Click
    Exit Sub
Else
    IsDirty = False
End If



Print #1, RetStr
Close #1


End Sub

Private Sub mnuSaveAs_Click()

On Error Resume Next
cd.CancelError = True
cd.Filter = "Entity Studio Schema|*.xml"
cd.Flags = cdlOFNHideReadOnly + cdlOFNPathMustExist
cd.FileName = "schema"
cd.DefaultExt = "xml"
cd.ShowSave

If Err Then Exit Sub

Dim RetStr
'Get the current schema as xml.
RetStr = Schema1.DumpSchema()

On Error Resume Next
Open cd.FileName For Output As #1

If Err Then
MsgBox "File error. Unable to save changes. Try saving the file with a different name", vbCritical
Else
CurFile = cd.FileName
Me.Caption = "Entity Studio - " & cd.FileName
IsDirty = False
End If

Print #1, RetStr
Close #1

End Sub

Private Sub mnuSchema_Click()
On Error Resume Next
If AskSave() = vbCancel Then Exit Sub

cd.CancelError = True
cd.Filter = "Entity Studio Schemas|*.xml"
cd.Flags = cdlOFNFileMustExist + cdlOFNHideReadOnly
cd.InitDir = App.Path & "\Schema"

cd.ShowOpen

If Err Then Exit Sub

If Schema1.LoadSchemaFromFile(cd.FileName) = False Then
    MsgBox "Invalid Entity Studio Schema. Kindly specify a valid Entity Studio Schema", vbCritical + vbOKOnly, "Cannot Open"
Else
    'Arrange the control's entities automatically
    Schema1.Arrange
    CurFile = cd.FileName
    Me.Caption = "Entity Studio - " & cd.FileName
    IsDirty = False
End If



End Sub

Private Sub mnuUsers_Click()
    '
End Sub


Private Sub mnuWelcome_Click()
On Error Resume Next
SaveSetting App.EXEName, "Options", "Show Tips at Startup", 1
frmTip.Show , Me
End Sub

Private Sub Schema1_Dirty()
IsDirty = True

End Sub

Private Sub tbMain_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case VBA.LCase(Button.Key)
    Case "import"
        
        PopupMenu mnuImport, , tbMain.Left + Button.Left, tbMain.Top + Button.Top + Button.Height
        
    Case "open"
         mnuSchema_Click
    Case "save"
        mnuSave_Click
    Case "generate"
        PopupMenu mnuExport, , tbMain.Left + Button.Left, tbMain.Top + Button.Top + Button.Height
    Case "users"
        mnuUsers_Click
    Case "entities"
        mnuEntities_Click
    Case "relations"
        mnuRelations_Click
    Case "help"
        mnuHelpContents_Click
End Select

End Sub
