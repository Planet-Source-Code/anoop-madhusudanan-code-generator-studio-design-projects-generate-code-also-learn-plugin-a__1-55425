VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form frmTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Schema Model Test"
   ClientHeight    =   6015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   7485
   StartUpPosition =   3  'Windows Default
   Begin SHDocVwCtl.WebBrowser wbMain 
      Height          =   4845
      Left            =   195
      TabIndex        =   3
      Top             =   900
      Width           =   7170
      ExtentX         =   12647
      ExtentY         =   8546
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin MSComDlg.CommonDialog cdMain 
      Left            =   240
      Top             =   4440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   90
      Left            =   150
      TabIndex        =   2
      Top             =   675
      Width           =   7320
   End
   Begin VB.TextBox txtRead 
      Height          =   285
      Left            =   1335
      TabIndex        =   1
      Top             =   240
      Width           =   5760
   End
   Begin VB.CommandButton cmdRead 
      Caption         =   "Read From"
      Height          =   330
      Left            =   150
      TabIndex        =   0
      Top             =   225
      Width           =   1065
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'====================================================================
'Schma Model Demo.
'====================================================================

'====================================================================
Private Sub cmdRead_Click()
'====================================================================
On Error Resume Next
Dim Sch As New SchemaModel.Schema

    cdMain.Filter = "Entity Studio XML Files|*.xml"
    cdMain.ShowOpen
    Me.txtRead = cdMain.FileName
    
    If Sch.LoadSchema(Me.txtRead) = False Then
         MsgBox "Invalid xml file"
    Else
        'Demonstrates iterating the schema
        IterateAll Sch
        wbMain.Navigate2 App.Path & "\temp.xml"
    End If
    
Set Sch = Nothing

'====================================================================
End Sub
'====================================================================

Private Sub cmdWrite_Click()
    'Sch.SaveSchema Filename

End Sub


Function Space(Optional C = 1)
Dim Rets
Rets = ""
Dim i
For i = 0 To C
    Rets = Rets & "     "
Next
Space = Rets
End Function


'====================================================================
Function IterateAll(Sch As Schema)
'Shows iterating a schema. Try to create some vb classes
'out of this.
'====================================================================
    
Dim thisRel As Relation
Dim thisFld As Field
Dim thisEnt As Entity
Dim thisAttrib As Attrib

Open App.Path & "\temp.xml" For Output As #1
    
        Print #1, "<schema>"
        
        'Read the attributes of the schema
        For Each thisAttrib In Sch.SchemaAttributes
            Print #1, RetXML(thisAttrib.AttribName, thisAttrib.AttribValue)
        Next
    
        'Read all entities
        For Each thisEnt In Sch.SchemaEntities
            
            Print #1, "<entity>" & vbCrLf
        
            'Read all entity attributes
            For Each thisAttrib In thisEnt.EntityAttributes
                Print #1, RetXML(thisAttrib.AttribName, thisAttrib.AttribValue)
            Next
            
            'Read all field fields
            For Each thisFld In thisEnt.Fields
            Print #1, "<field>" & vbCrLf
            
                'Read all field header attributes
                For Each thisAttrib In thisFld.FieldHeaderAttributes
                    Print #1, RetXML(thisAttrib.AttribName, thisAttrib.AttribValue)
                Next
            
                Print #1, "<attributes>" & vbCrLf
                
                'Read all field attributes
                For Each thisAttrib In thisFld.FieldAttributes
                    Print #1, RetXML(thisAttrib.AttribName, thisAttrib.AttribValue)
                Next
            
                Print #1, "</attributes>" & vbCrLf
            
            Print #1, "</field>" & vbCrLf
            Next
            
            
            
        
        
            Print #1, "</entity>" & vbCrLf
        Next
        
        
        'Read all relations
        For Each thisRel In Sch.SchemaRelations
        
            Print #1, "<relation>" & vbCrLf
            
            'Read all relation attributes
                For Each thisAttrib In thisRel.RelationAttributes
                    Print #1, RetXML(thisAttrib.AttribName, thisAttrib.AttribValue)
                Next
            
            Print #1, "</relation>" & vbCrLf
        
        
        Next
        
            Print #1, "</schema>" & vbCrLf
    
Close #1
    
'====================================================================
End Function
'====================================================================

Function RetXML(Param, Value)
Dim oXml As New SchemaModel.SchemaXML
RetXML = oXml.RetXML(Param, Value) & vbCrLf
End Function
