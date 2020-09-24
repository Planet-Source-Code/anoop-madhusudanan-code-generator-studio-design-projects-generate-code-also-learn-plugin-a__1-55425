VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsASPEngine"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'====================================================================
'ASPEngine
'====================================================================
'Built In Engine System For Creating ASP Code
'====================================================================
'By Anoop M. All Rights Reserved
'http://www.logicmatrixonline.com/anoop
'====================================================================
'License: This project comes with GNU GENERAL PUBLIC
'LICENSE. See License.htm

'Notes: Some part of this code was found from PSC, I made
'changes to it and adapted it for ASP code generation from a
'schema

'Important: The ASP generator is NOT using the
'flexible schema model. This is depreciated. Use the SchemaModel
'for generating code, as in VB Export engine.




Dim FIELDPROP(17) As String
Dim TableBGColor As String
Dim FirstHalfSQL As String
Dim SecondHalfSQL As String
Dim EndSQL As String


Dim VAR_TITLE As String
Dim VAR_SUBTITLE As String
Dim VAR_NAVBAR As String
Dim VAR_SUBNAVBAR As String


Dim Tables As Collection
Dim Fields As Collection
Dim Rels As Collection

Dim Fieldcount

Dim gMAPSTR As String
'Dim Tables() As String, Fields() As String, Rels() As String, FieldCount

Dim HeaderString
Dim FrontEndForm As frmProject

Public Function Create(Path, ProjectName, Sch As SchemaModel.Schema, AppObject As Object, FormObject)

'Function to create the project
Set FrontEndForm = FormObject

On Error Resume Next


PutStatus "Loading the schema to project.."

FrontEndForm.pBar.Max = Sch.SchemaEntities.Count + 5

PutStatus "Starting project creation.."


If Not CreateProject(Sch.SchemaAttributes("name"), Path, ProjectName, AppObject.Path & "\Export\Asp\templates\style.htm", AppObject.Path & "\Export\Asp\core", Sch) Then
    MsgBox "Unable to create the project. Kindly make sure that you provided a valid name and directory", vbInformation + vbOKOnly, "Unable To Create"
    GoTo ErrorExit
Else
    PutStatus "Project Created To The Path."
    MsgBox "Project created to the specified path. You may point your browser to the location and view the project. Eg: http://localhost/" & FrontEndForm.txtName, vbInformation + vbOKOnly, "Project Created"
    FrontEndForm.Hide
End If

ErrorExit:


End Function


Public Sub PutStatus(Status)
    FrontEndForm.lblStatus.Refresh
    FrontEndForm.lblStatus = Status
End Sub



Public Function CreateProject(iDbPath, iUserRootPath, iProjectName, iTemplate, iCorePath, iSch As SchemaModel.Schema) As Boolean
    
    Dim intProjectID As Integer
    Dim intTableID As Integer
    Dim intColumnID As Integer
    Dim UserName As String
    Dim iProjectRootPath As String
    Dim iProjectTablePath As String
    Dim fsoObject As New FileSystemObject
    Dim bDoesFolderExist As Boolean
    Dim xCount As Integer
    Dim iProjectID As String
    Dim CurrentPrimaryKey As String
    Dim tfolders As Variant
    Dim driveRoot As String
    Dim zCount As Integer
    Dim CurrentPath As String
    
    
    
    
    VAR_TITLE = ""
    VAR_SUBTITLE = ""
    VAR_NAVBAR = ""
    VAR_SUBNAVBAR = ""
    gMAPSTR = ""
    
    VAR_TITLE = iProjectName
    
    
    
    
    
'on error GoTo PError


    If VBA.Right(iUserRootPath, 1) = "\" Then
        iUserRootPath = Left(iUserRootPath, Len(iUserRootPath) - 1)
    End If
    
    
If Not fsoObject.FolderExists(iProjectRootPath) Then
    CreateProject = False
End If
    
    
    tfolders = Split(iUserRootPath, "\", , vbTextCompare)
    driveRoot = tfolders(0)
    
    
    'Creating folder for project
    
    For xCount = 1 To UBound(tfolders)
        CurrentPath = driveRoot
        For zCount = 1 To xCount
            CurrentPath = CurrentPath & "\" & tfolders(zCount)
        Next zCount
        If Not fsoObject.FolderExists(CurrentPath) Then
            Call fsoObject.CreateFolder(CurrentPath)
        End If
    Next
    
    xCount = 1
    iProjectRootPath = iUserRootPath & "\" & iProjectName
    
    'Creates a unique path
    If fsoObject.FolderExists(iProjectRootPath) Then
        bDoesFolderExist = True
        While bDoesFolderExist = True
            xCount = xCount + 1
            iProjectRootPath = iUserRootPath & "\" & iProjectName & xCount
            bDoesFolderExist = fsoObject.FolderExists(iProjectRootPath)
        Wend
        iProjectRootPath = iUserRootPath & "\" & iProjectName & xCount
    End If
    
    iProjectID = xCount
    
    
    'Adds the project folder
    Call AddProjectFolder(iDbPath, iUserRootPath, iProjectID, iProjectName, iProjectRootPath)
    
    
    
    VAR_NAVBAR = GetLink("Home", "../", "#ffffff")
    
    
    Dim thisEnt As Entity
    Dim thisFld As SchemaModel.Field
    Dim thisRel As SchemaModel.Relation
    
    'Building global navigation bar
    For Each thisEnt In iSch.SchemaEntities
        VAR_NAVBAR = VAR_NAVBAR & " :: " & GetLink(Replace(thisEnt.EntityAttributes("name"), "tbl_", ""), "../" & thisEnt.EntityAttributes("name") & "/list.asp", "#ffffff")
    Next
    
            
            
    Dim CurFieldIndex
            
    For Each thisEnt In iSch.SchemaEntities
    
            
            
            Set Fields = Nothing
            Set Tables = Nothing
            Set Rels = Nothing
            
            Set Fields = New Collection
            Set Tables = New Collection
            Set Rels = New Collection
            
            
            Fieldcount = 0
            CurFieldIndex = 0
            
            
    
    
            iProjectTablePath = iProjectRootPath & "\" & thisEnt.EntityAttributes("name")
            
            VAR_SUBTITLE = thisEnt.EntityAttributes("name")
            
            Call AddTableFolder(iUserRootPath, iProjectRootPath, iProjectTablePath, iProjectID, thisEnt.EntityAttributes("name"), iCorePath, iTemplate)
                PutStatus "Creating Code For " & thisEnt.EntityAttributes("name")
            'on error Resume Next
            FrontEndForm.pBar.Value = thisEnt.InnerIndex

            'Building Map String
            gMAPSTR = ""
            gMAPSTR = gMAPSTR & "Redim FieldMap(" & thisEnt.Fields.Count & ",7)" & vbCrLf
            
            
            Dim FieldCounter
            For FieldCounter = thisEnt.Fields.Count - 1 To 0 Step -1
                
                Set thisFld = thisEnt.Fields(FieldCounter + 1)
                Erase FIELDPROP
                

                Fieldcount = Fieldcount + 1
                
                    With thisFld
                    
                            On Error Resume Next
                            
                            'String To Data Type
                            STDT .FieldAttributes("type"), FIELDPROP(2), FIELDPROP(1)
                            
                            'Attribute
                            
                            
                            FIELDPROP(3) = .FieldAttributes("ordinalposition")   'Position
                            FIELDPROP(4) = .FieldAttributes("size") 'Size
                            FIELDPROP(5) = .FieldAttributes("defaultvalue") 'DefualtValue
                            FIELDPROP(6) = .FieldAttributes("validationrule") 'ValidationRule
                            FIELDPROP(7) = .FieldAttributes("validationtext") 'ValidationText
                            
                            If FIELDPROP(2) = 11 Then
                                FIELDPROP(8) = "False"    'Not required if autotype
                            Else
                                FIELDPROP(8) = .FieldAttributes("required") 'Required
                            End If
                            
                            FIELDPROP(9) = .FieldAttributes("allowzerolength")     'Required
                            
                            FIELDPROP(10) = .FieldAttributes("decimalplaces")    'Decimal Places
                            FIELDPROP(11) = .FieldAttributes("format")  'Format
                            
                            FIELDPROP(12) = .FieldHeaderAttributes("name")              'Name
                            FIELDPROP(13) = .FieldHeaderAttributes("primary")          'Primary Or Not
                            
                            FIELDPROP(14) = .FieldAttributes("caption")   'Caption
                            
                            On Error GoTo 0
                            
                            
                        'Setting relations
                        
                        FIELDPROP(15) = ""
                        FIELDPROP(16) = ""
                        
                        
                        For Each thisRel In iSch.SchemaRelations
                            If VBA.LCase(thisRel.RelationAttributes("foreignfield")) = VBA.LCase(.FieldHeaderAttributes("name")) And VBA.LCase(thisRel.RelationAttributes("foreigntable")) = VBA.LCase(thisEnt.EntityAttributes("name")) Then
                                FIELDPROP(15) = thisRel.RelationAttributes("field")
                                FIELDPROP(16) = thisRel.RelationAttributes("table")
                                FIELDPROP(2) = 100 'Datacombo
                            End If
                        Next
                            
                            
                   End With
               
               
               
               Err.Clear
                Call CreateFieldMap(thisEnt.EntityAttributes("name"), thisFld.FieldHeaderAttributes("name"), FIELDPROP(16), FIELDPROP(15), FIELDPROP(14), GetBoolean(thisFld.FieldHeaderAttributes("display")), FieldCounter, GetBoolean(thisFld.FieldHeaderAttributes("primary")), iSch)
                
                Call CreateAddTablePostHTML(iProjectRootPath, iProjectID, thisEnt.EntityAttributes("name"), iSch)
                Call CreateAddTableASP(iProjectRootPath, iProjectID, thisEnt.EntityAttributes("name"))
                Call CreateViewTable(iProjectRootPath, iProjectID, thisEnt.EntityAttributes("name"))
                Call CreateDeleteTable(iProjectRootPath, iProjectID, thisEnt.EntityAttributes("name"))
                Call CreateUpdateTablePostHTML(iProjectRootPath, iProjectID, thisEnt.EntityAttributes("name"), iSch)
                Call CreateUpdateTableASP(iProjectRootPath, iProjectID, thisEnt.EntityAttributes("name"))
                
                            
               
            Next
            
                Call CreateGetTableData(iProjectRootPath, iProjectID, thisEnt.EntityAttributes("name"))
                Call CreateSearch(iProjectRootPath, iProjectID, thisEnt.EntityAttributes("name"))
    
    Next
     
     
     
    WriteInclude iProjectRootPath, iProjectRootPath & "\database.mdb", iCorePath
    WriteCodeFile iProjectRootPath & "\default.asp", iProjectRootPath & "\default.asp", iTemplate, "All Files", True
    
    CreateProject = True
    FrontEndForm.pBar.Value = FrontEndForm.pBar.Max
    Exit Function
    
PError:
    CreateProject = False
    
End Function

Private Function AddProjectFolder(ByVal iDbPath, iUserRootPath, iProjectID, iProjectName, iProjectRootPath) As Boolean

    'this method builds the project folder and initial default.asp page that will
    'show a listing of the tables in the project.
    Dim fsoObject As New FileSystemObject
    Dim fsoTextStream As TextStream
    Dim fsoTextStreamTemp As TextStream
    Dim CurrentLine As String
    Dim fsoTextStreamTable As TextStream
    Dim BGColor As String
    
    
    Call fsoObject.CreateFolder(iProjectRootPath)
    Call fsoObject.CopyFile(iDbPath, iProjectRootPath & "\database.mdb")
    
    
    'Here, I have to create iProjectRootPath & "\default.asp" file
    
    iDbPath = iProjectRootPath & "\database.mdb"

    Set fsoTextStreamTable = fsoObject.CreateTextFile(iProjectRootPath & "\default.asp")
    
    'on error Resume Next
    
    Call fsoTextStreamTable.WriteLine("<table border=""0"" width=""100%"" cellpadding=""2"" bgcolor=""#000080"">")
    Call fsoTextStreamTable.WriteLine("<tr>")
    Call fsoTextStreamTable.WriteLine("<td><font face=""verdana,arial"" size=""1"" color=""#FFFFFF"">Table Directory</font></td>")
    Call fsoTextStreamTable.WriteLine("</tr>")
    Call fsoTextStreamTable.WriteLine("</table>")
    Call fsoTextStreamTable.WriteLine("<br>")
    Call fsoTextStreamTable.WriteLine("<table border=""0"" width=""100%"">")
    Call fsoTextStreamTable.WriteLine("<font face=""verdana,arial"" size=""1"" color=""#000000"">")
    Call fsoTextStreamTable.WriteLine("<!--$Definition -->")
    Call fsoTextStreamTable.WriteLine("</font>")
    Call fsoTextStreamTable.WriteLine("</table>")
    Call fsoTextStreamTable.Close
    
    Set fsoTextStreamTable = Nothing
    
    
    
End Function

Private Function AddTableFolder(iUserRootPath, iProjectRootPath, iProjectTablePath, iProjectID, iTableName, iCorePath, iTemplate) As Boolean
        
    'for each table that exists i create it's own subdirectory folder. this folder
    'contains its own custom pages for insert / update / delete and select functionality.
    'some of these pages will contain tags like this... <!--$Definition -->.
    'all of the code you see in this method is standard code for each table, no matter what
    'the design for the given table. what will happen in later functions is that we will modify
    'the asp code to customize it for the table that requests it. as i read each field in the
    'CreateProject method i will modify these pages specifically for that field...
    Dim fsoObject As New FileSystemObject
    Dim fsoTextStream As TextStream
    Dim fsoTextStreamTemp As TextStream
    Dim fsoTextStreamAddPage As TextStream
    Dim CurrentLine As String, TbPath
    
    
    Call fsoObject.CreateFolder(iProjectRootPath & "\" & iTableName)
    
    TbPath = iProjectRootPath & "\" & iTableName
    
    'on error Resume Next
    
    WriteTableInclude TbPath, iTableName
    
    Set fsoTextStream = fsoObject.OpenTextFile(iProjectRootPath & "\default.asp", ForReading)
    Set fsoTextStreamTemp = fsoObject.CreateTextFile(iProjectRootPath & "\defaultTemp.asp", True)
    
    While Not fsoTextStream.AtEndOfStream
        
        CurrentLine = fsoTextStream.ReadLine
        fsoTextStreamTemp.WriteLine CurrentLine
        
        If CurrentLine = "<!--$Definition -->" Then
        
            If TableBGColor = "#C0C0C0" Then
                TableBGColor = "#F3F3DC"
            Else
                TableBGColor = "#C0C0C0"
            End If
            Dim fFace
            fFace = "<font face=""verdana,arial"" size=""1"">"
        
            Call fsoTextStreamTemp.WriteLine("<tr>")
            Call fsoTextStreamTemp.WriteLine("<td width=""200"" bgcolor=""" & TableBGColor & """>" & fFace & iTableName & "</font></td>")
            Call fsoTextStreamTemp.WriteLine("<td width=""125"" bgcolor=""" & TableBGColor & """><center><a href=""" & iTableName & "/new.asp"">" & fFace & "Add New Record </font></a> </center></td>")
            Call fsoTextStreamTemp.WriteLine("<td width=""125"" bgcolor=""" & TableBGColor & """><center><a href=""" & iTableName & "/list.asp"">" & fFace & "View Data</font></a></center></td>")
            Call fsoTextStreamTemp.WriteLine("</tr>")
        End If
    Wend
    
    Call fsoTextStream.Close
    Call fsoTextStreamTemp.Close
    
    Call fsoObject.DeleteFile(iProjectRootPath & "\default.asp", True)
    Call fsoObject.MoveFile(iProjectRootPath & "\defaultTemp.asp", iProjectRootPath & "\default.asp")
    
    Set fsoTextStreamAddPage = Nothing
    
    
    Dim s As String, Path As String
      Path = iCorePath & "\*.*"
      s = Dir$(Path, vbNormal)
      Do While LenB(s) <> 0
            WriteCodeFile TbPath & "\" & s, iCorePath & "\" & s, iTemplate, iTableName
            s = Dir
      Loop
    
End Function

Private Function CreateAddTablePostHTML(iProjectRootPath, iProjectID, iTableName, Sch As SchemaModel.Schema) As Boolean

    'this is where i modify the new.asp page to customize a given table. this method
    'will be called once for each field in the table.
    Dim Datatype As Integer
    Dim attrib As Long
    Dim FieldName As String, FieldCaption As String
    Dim fsoTextStream As TextStream
    Dim fsoTextStreamTemp As TextStream
    Dim fsoObject As New FileSystemObject
    Dim CurrentLine As String
    Dim Required As String
    Dim StarRequired As String
    Dim PrimaryKey As Boolean
    Dim StarPrimary As String
    Dim Table, Field
    
    Dim Control, Caption
    
    attrib = FIELDPROP(1)
    Datatype = FIELDPROP(2)
    
    FieldName = ""
    FieldName = FIELDPROP(12)
    FieldCaption = ""
    FieldCaption = FIELDPROP(14)
    
    If Trim(FieldCaption) = "" Then FieldCaption = FIELDPROP(12)
    
    
    Required = FIELDPROP(8)
    PrimaryKey = FIELDPROP(13)
    Table = FIELDPROP(16)
    Field = FIELDPROP(15)
    
    
    Set fsoTextStream = fsoObject.OpenTextFile(iProjectRootPath & "\" & iTableName & "\new.asp", ForReading)
    Set fsoTextStreamTemp = fsoObject.CreateTextFile(iProjectRootPath & "\" & iTableName & "\addTemp.asp", True)
    
    'on error Resume Next
    
    While Not fsoTextStream.AtEndOfStream
        
        CurrentLine = fsoTextStream.ReadLine
        fsoTextStreamTemp.WriteLine CurrentLine
        
        If CurrentLine = "<!--$Definition -->" Then
            
            Control = ""
            Caption = ""
            
            If Required = "True" Then
                StarRequired = "<font color=""#FF0000"">*</font>"
            Else
                StarRequired = ""
            End If
            
                StarPrimary = ""
            
            Select Case Datatype
                Case 10            '     Text
                    Caption = ("" & FieldCaption & StarRequired & StarPrimary & ":<br>")
                    Control = ("<input type=""text"" name=""" & FieldName & """ size=""20"">")
                    
                Case 100            '     Relation Number
                    Caption = ("" & FieldCaption & StarRequired & StarPrimary & ":<br>")
                    Control = ("<%WriteCatCombo " & QT(FieldName) & "," & QT(Field) & "," & QT(GetNextField(Table, Field, Sch)) & "," & QT(Table) & ",-1 %>")
                Case 12            '     Memo
                    Select Case attrib
                        Case 2     '     memo
                            Caption = ("" & FieldCaption & StarRequired & StarPrimary & ":<br>")
                            Control = ("<textarea rows=""4"" name=""" & FieldName & """ cols=""40""></textarea>")
                        Case 32770 'hyperlink
                            Caption = ("" & FieldCaption & StarRequired & StarPrimary & ":<br>")
                            Control = ("<input type=""text"" name=""" & FieldName & """ size=""20"">")
                    End Select
                Case 3, 4, 2, 6, 7, 15, 20
                    Select Case attrib
                        Case 17    'Autonumber
                                   'do nothing
                        Case Else     '    Number
                            Caption = ("" & FieldCaption & StarRequired & StarPrimary & ":<br>")
                            Control = ("<input type=""text"" name=""" & FieldName & """ size=""20"">")
                    End Select
                Case 8             '  DateTime
                    Caption = ("" & FieldCaption & StarRequired & StarPrimary & ":<br>")
                    Control = ("<input type=""text"" name=""" & FieldName & """ size=""20"">")
                Case 5             '  Currency
                    Caption = ("" & FieldCaption & StarRequired & StarPrimary & ":<br>")
                    Control = ("<input type=""text"" name=""" & FieldName & """ size=""20"">")
                Case 1             '     YesNo
                    Caption = ("" & FieldCaption & StarRequired & StarPrimary & ":  ")
                    Control = ("Yes <input type=""radio"" value=""Yes"" checked name=""" & FieldName & """> No <input type=""radio"" name=""" & FieldName & """ value=""No"">")
                Case 11            'OLEOBject
                                   'is not supported
            End Select
            
            
            If Control <> "" And Caption <> "" Then
                    fsoTextStreamTemp.WriteLine "<tr>"
                    fsoTextStreamTemp.WriteLine "<td width=" & QT("30%") & "><font size=" & QT("1") & ">" & Caption & "</font></td>"
                    fsoTextStreamTemp.WriteLine "<td width=" & QT("70%") & ">" & Control & "</td>"
                    fsoTextStreamTemp.WriteLine "</tr>"
            End If
            
        End If
    Wend
    
    Call fsoTextStream.Close
    Call fsoTextStreamTemp.Close
    
    Call fsoObject.DeleteFile(iProjectRootPath & "\" & iTableName & "\new.asp", True)
    Call fsoObject.MoveFile(iProjectRootPath & "\" & iTableName & "\addTemp.asp", iProjectRootPath & "\" & iTableName & "\new.asp")
    
    
End Function

Private Function CreateAddTableASP(iProjectRootPath, iProjectID, iTableName) As Boolean

    'this method is where i customize the newcheck.asp page to handle and insert into the
    'database anything that has been posted by the new.asp page.
    Dim Datatype As Integer
    Dim attrib As Long
    Dim FieldName As String, FieldCaption As String
    Dim fsoTextStream As TextStream
    Dim fsoTextStreamTemp As TextStream
    Dim fsoObject As New FileSystemObject
    Dim CurrentLine As String
    Dim Required As String
    Dim Star As String
    Dim FieldNameVariable As String
    Dim PrimaryKey As Boolean
    
    attrib = FIELDPROP(1)
    Datatype = FIELDPROP(2)
    FieldName = FIELDPROP(12)
    Required = FIELDPROP(8)
    PrimaryKey = FIELDPROP(13)
    
    FieldCaption = ""
    FieldCaption = FIELDPROP(14)
    
    If Trim(FieldCaption) = "" Then FieldCaption = FIELDPROP(12)

    
    Set fsoTextStream = fsoObject.OpenTextFile(iProjectRootPath & "\" & iTableName & "\newcheck.asp", ForReading)
    Set fsoTextStreamTemp = fsoObject.CreateTextFile(iProjectRootPath & "\" & iTableName & "\ASPaddTemp.asp", True)
    
    While Not fsoTextStream.AtEndOfStream
        
        CurrentLine = fsoTextStream.ReadLine
        
        fsoTextStreamTemp.WriteLine CurrentLine
        
        If CurrentLine = "<!--$Definition -->" And attrib <> 17 And Datatype <> 11 Then
            
            fsoTextStreamTemp.WriteLine ("<%")
            
            FieldNameVariable = Replace(FieldName, " ", "")
            FieldNameVariable = "var" & FieldNameVariable
            
            If Required = "True" Or PrimaryKey = True Then
                fsoTextStreamTemp.WriteLine ("If Request.Form(""" & FieldName & """) = """" Then")
                fsoTextStreamTemp.WriteLine ("Error=" & QT("The Field " & FieldCaption & " is empty"))
                fsoTextStreamTemp.WriteLine ("End If")
            End If
            
            Select Case Datatype
                Case 10            '     Text
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                    fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                    fsoTextStreamTemp.WriteLine ("iFieldCount = iFieldCount + 1")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = Replace(" & FieldNameVariable & ", ""'"", ""''"")")
                    
                    fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & ""[" & FieldName & "]""")
                    fsoTextStreamTemp.WriteLine ("Else")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "",[" & FieldName & "]""")
                    fsoTextStreamTemp.WriteLine ("End If")
                    
                    fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                    fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & ""'"" & " & FieldNameVariable & " & ""'""")
                    fsoTextStreamTemp.WriteLine ("Else")
                    fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & "",'"" & " & FieldNameVariable & " & ""'""")
                    fsoTextStreamTemp.WriteLine ("End If")
                    fsoTextStreamTemp.WriteLine ("End If")
                    
                Case 12            '     Memo
                    Select Case attrib
                        Case 2     '     memo
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                            fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                            fsoTextStreamTemp.WriteLine ("iFieldCount = iFieldCount + 1")
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = Replace(" & FieldNameVariable & ", ""'"", ""''"")")
                    
                            fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & ""[" & FieldName & "]""")
                            fsoTextStreamTemp.WriteLine ("Else")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "",[" & FieldName & "]""")
                            fsoTextStreamTemp.WriteLine ("End If")
                    
                            fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                            fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & ""'"" & " & FieldNameVariable & " & ""'""")
                            fsoTextStreamTemp.WriteLine ("Else")
                            fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & "",'"" & " & FieldNameVariable & " & ""'""")
                            fsoTextStreamTemp.WriteLine ("End If")
                            fsoTextStreamTemp.WriteLine ("End If")
                        
                        Case 32770 'hyperlink
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                            fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                            fsoTextStreamTemp.WriteLine ("iFieldCount = iFieldCount + 1")
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = Replace(" & FieldNameVariable & ", ""'"", ""''"")")
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = " & FieldNameVariable & " & ""#"" & " & FieldNameVariable & " & ""#""")
                            
                            fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & ""[" & FieldName & "]""")
                            fsoTextStreamTemp.WriteLine ("Else")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "",[" & FieldName & "]""")
                            fsoTextStreamTemp.WriteLine ("End If")
                    
                            fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                            fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & ""'"" & " & FieldNameVariable & " & ""'""")
                            fsoTextStreamTemp.WriteLine ("Else")
                            fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & "",'"" & " & FieldNameVariable & " & ""'""")
                            fsoTextStreamTemp.WriteLine ("End If")
                            fsoTextStreamTemp.WriteLine ("End If")
                    End Select
                Case 3, 4, 2, 6, 7, 15, 20, 100
                    Select Case attrib
                        Case 17    'Autonumber
                                   'do nothing
                        Case Else     '    Number
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                            fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                            fsoTextStreamTemp.WriteLine ("iFieldCount = iFieldCount + 1")
                            fsoTextStreamTemp.WriteLine ("'on error resume next")
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = CDbl(" & FieldNameVariable & ")")
                            fsoTextStreamTemp.WriteLine ("If Err.Number = 13 Then")
                            fsoTextStreamTemp.WriteLine ("Error=" & QT("The data you entered in the field " & FieldCaption & " is invalid"))
                            fsoTextStreamTemp.WriteLine ("End If")
                            fsoTextStreamTemp.WriteLine ("Err.Clear")
                            fsoTextStreamTemp.WriteLine ("''on error GoTo 0")

                            fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & ""[" & FieldName & "]""")
                            fsoTextStreamTemp.WriteLine ("Else")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "",[" & FieldName & "]""")
                            fsoTextStreamTemp.WriteLine ("End If")
                    
                            fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                            fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & " & FieldNameVariable)
                            fsoTextStreamTemp.WriteLine ("Else")
                            fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & "","" & " & FieldNameVariable)
                            fsoTextStreamTemp.WriteLine ("End If")
                            fsoTextStreamTemp.WriteLine ("End If")
                    End Select
                Case 8             '  DateTime
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                    fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                    fsoTextStreamTemp.WriteLine ("iFieldCount = iFieldCount + 1")
                    fsoTextStreamTemp.WriteLine ("'on error resume next")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = CDate(" & FieldNameVariable & ")")
                    fsoTextStreamTemp.WriteLine ("If Err.Number = 13 Then")
                    fsoTextStreamTemp.WriteLine ("Error=" & QT("The data you entered in the field " & FieldCaption & " is invalid"))
                    fsoTextStreamTemp.WriteLine ("End If")
                    fsoTextStreamTemp.WriteLine ("Err.Clear")
                    fsoTextStreamTemp.WriteLine ("''on error GoTo 0")
                    
                    fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & ""[" & FieldName & "]""")
                    fsoTextStreamTemp.WriteLine ("Else")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "",[" & FieldName & "]""")
                    fsoTextStreamTemp.WriteLine ("End If")
                    
                    fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                    fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & ""'"" & " & FieldNameVariable & " & ""'""")
                    fsoTextStreamTemp.WriteLine ("Else")
                    fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & "",'"" & " & FieldNameVariable & " & ""'""")
                    fsoTextStreamTemp.WriteLine ("End If")
                    fsoTextStreamTemp.WriteLine ("End If")
                Case 5             '  Currency
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                    fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                    fsoTextStreamTemp.WriteLine ("iFieldCount = iFieldCount + 1")
                    fsoTextStreamTemp.WriteLine ("'on error resume next")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = CCur(" & FieldNameVariable & ")")
                    fsoTextStreamTemp.WriteLine ("If Err.Number = 13 Then")
                    fsoTextStreamTemp.WriteLine ("Error=" & QT("The data you entered in the field " & FieldCaption & " is invalid"))
                    fsoTextStreamTemp.WriteLine ("End If")
                    fsoTextStreamTemp.WriteLine ("Err.Clear")
                    fsoTextStreamTemp.WriteLine ("''on error GoTo 0")
                    
                    fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & ""[" & FieldName & "]""")
                    fsoTextStreamTemp.WriteLine ("Else")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "",[" & FieldName & "]""")
                    fsoTextStreamTemp.WriteLine ("End If")
                    
                    fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                    fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & " & FieldNameVariable)
                    fsoTextStreamTemp.WriteLine ("Else")
                    fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & "","" & " & FieldNameVariable)
                    fsoTextStreamTemp.WriteLine ("End If")
                    fsoTextStreamTemp.WriteLine ("End If")
                Case 1             '     YesNo
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                    fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                    fsoTextStreamTemp.WriteLine ("iFieldCount = iFieldCount + 1")
                    fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " = ""Yes"" then")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = True")
                    fsoTextStreamTemp.WriteLine ("Else")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = False")
                    fsoTextStreamTemp.WriteLine ("End If")
                    
                    fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & ""[" & FieldName & "]""")
                    fsoTextStreamTemp.WriteLine ("Else")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "",[" & FieldName & "]""")
                    fsoTextStreamTemp.WriteLine ("End If")
                    
                    fsoTextStreamTemp.WriteLine ("If iFieldCount = 1 Then")
                    fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & " & FieldNameVariable)
                    fsoTextStreamTemp.WriteLine ("Else")
                    fsoTextStreamTemp.WriteLine ("SecondHalfSQL = SecondHalfSQL & "","" & " & FieldNameVariable)
                    fsoTextStreamTemp.WriteLine ("End If")
                    fsoTextStreamTemp.WriteLine ("End If")
                Case 11            'OLEOBject
                                   'is not supported
            End Select
            
            fsoTextStreamTemp.WriteLine ("%>")
        End If
    Wend
    
    Call fsoTextStream.Close
    Call fsoTextStreamTemp.Close
    
    Call fsoObject.DeleteFile(iProjectRootPath & "\" & iTableName & "\newcheck.asp", True)
    Call fsoObject.MoveFile(iProjectRootPath & "\" & iTableName & "\ASPaddTemp.asp", iProjectRootPath & "\" & iTableName & "\newcheck.asp")
    
End Function

Private Function CreateViewTable(iProjectRootPath, iProjectID, iTableName) As Boolean

    'this method builds a table with checkboxes for each field in a table. this page
    'will be used to allow the user to select which fields he wants displayed in
    'the report of the data. in the future this page will also display query builder
    'functionality.
    Dim Datatype As Integer
    Dim attrib As Long
    Dim FieldName As String
    Dim fsoTextStream As TextStream
    Dim fsoTextStreamTemp As TextStream
    Dim fsoObject As New FileSystemObject
    Dim CurrentLine As String
    Dim Required As String
    Dim Star As String
    
    attrib = FIELDPROP(1)
    Datatype = FIELDPROP(2)
    FieldName = FIELDPROP(12)
    Required = FIELDPROP(8)
    
    Set fsoTextStream = fsoObject.OpenTextFile(iProjectRootPath & "\" & iTableName & "\view.asp", ForReading)
    Set fsoTextStreamTemp = fsoObject.CreateTextFile(iProjectRootPath & "\" & iTableName & "\viewtemp.asp", True)
    
    While Not fsoTextStream.AtEndOfStream
        
        CurrentLine = fsoTextStream.ReadLine
        fsoTextStreamTemp.WriteLine CurrentLine
        
        If CurrentLine = "<!--$Definition -->" Then
            If Datatype <> 11 Then
                fsoTextStreamTemp.WriteLine ("<p><input type=""checkbox"" name=""" & FieldName & """ value=""Yes""> " & FieldName & "</p>")
            End If
        End If
    Wend
    
    Call fsoTextStream.Close
    Call fsoTextStreamTemp.Close
    
    Call fsoObject.DeleteFile(iProjectRootPath & "\" & iTableName & "\view.asp", True)
    Call fsoObject.MoveFile(iProjectRootPath & "\" & iTableName & "\viewtemp.asp", iProjectRootPath & "\" & iTableName & "\view.asp")
    
End Function

Private Function CreateGetTableData(iProjectRootPath, iProjectID, iTableName) As Boolean

    'this method creates the pages that will dynamically build sql queries and display
    'the given data that has been posted to the list.asp page.
    
    Dim fsoTextStream As TextStream
    Dim fsoTextStreamTemp As TextStream
    Dim fsoObject As New FileSystemObject
    Dim CurrentLine As String
    Dim IncStr
    

    Dim sqWhat, sqFrom, sqWhere, sqOrder
    
    sqWhat = "SELECT "
    sqFrom = "FROM "
    sqWhere = ""
    
    Dim i
    For i = Fields.Count To 1 Step -1
    If Fields(i) <> "" Then
        sqWhat = sqWhat & Fields(i) & ","
    End If
    Next i
    
    For i = Tables.Count To 1 Step -1
    If Tables(i) <> "" Then
        sqFrom = sqFrom & Tables(i) & ","
    End If
    Next i
    
    For i = Rels.Count To 1 Step -1
    If Rels(i) <> "" Then
    If sqWhere <> "" Then
        sqWhere = sqWhere & " AND " & Rels(i)
    Else
        sqWhere = "WHERE " & Rels(i)
    End If
    End If
    
    Next i
    
    'on error Resume Next
    sqFrom = Left(sqFrom, Len(sqFrom) - 1)
    sqWhat = Left(sqWhat, Len(sqWhat) - 1)
    
    
    IncStr = "<!-- #include file =" & QT("inc_fieldmap.inc") & " -->" & vbCrLf
    
    WriteFile iProjectRootPath & "\" & iTableName & "\inc_fieldmap.inc", "<%" & vbCrLf & gMAPSTR & vbCrLf & "%>"
    
    Set fsoTextStream = fsoObject.OpenTextFile(iProjectRootPath & "\" & iTableName & "\list.asp", ForReading)
    Set fsoTextStreamTemp = fsoObject.CreateTextFile(iProjectRootPath & "\" & iTableName & "\getdatatemp.asp", True)
     
    'on error Resume Next
    
    While Not fsoTextStream.AtEndOfStream
        
        CurrentLine = fsoTextStream.ReadLine
        fsoTextStreamTemp.WriteLine CurrentLine
        
        If CurrentLine = "<!--$Definition -->" Then
                fsoTextStreamTemp.WriteLine (IncStr)
                fsoTextStreamTemp.WriteLine ("")
                fsoTextStreamTemp.WriteLine ("<%")
                fsoTextStreamTemp.WriteLine ("FieldCount=" & Fieldcount)
                fsoTextStreamTemp.WriteLine ("sqWhat =" & QT(sqWhat))
                fsoTextStreamTemp.WriteLine ("sqFrom =" & QT(sqFrom))
                fsoTextStreamTemp.WriteLine ("sqWhere =" & QT(sqWhere))
                fsoTextStreamTemp.WriteLine ("%>")
        End If
    
    Wend
    
    Call fsoTextStream.Close
    Call fsoTextStreamTemp.Close
    
    Call fsoObject.DeleteFile(iProjectRootPath & "\" & iTableName & "\list.asp", True)
    Call fsoObject.MoveFile(iProjectRootPath & "\" & iTableName & "\getdatatemp.asp", iProjectRootPath & "\" & iTableName & "\list.asp")
    
End Function



Private Function CreateSearch(iProjectRootPath, iProjectID, iTableName) As Boolean

    'this method creates the pages that will dynamically build sql queries and display
    'the given data that has been posted to the list.asp page.
    
    Dim fsoTextStream As TextStream
    Dim fsoTextStreamTemp As TextStream
    Dim fsoObject As New FileSystemObject
    Dim CurrentLine As String
    Dim IncStr
    

    Dim sqWhat, sqFrom, sqWhere, sqOrder
    
    sqWhat = "SELECT "
    sqFrom = "FROM "
    sqWhere = ""
    
    Dim i
    For i = Fields.Count To 1 Step -1
    If Fields(i) <> "" Then
        sqWhat = sqWhat & Fields(i) & ","
    End If
    Next i
    
    For i = Tables.Count To 1 Step -1
    If Tables(i) <> "" Then
        sqFrom = sqFrom & Tables(i) & ","
    End If
    Next i
    
    For i = Rels.Count To 1 Step -1
    If Rels(i) <> "" Then
    If sqWhere <> "" Then
        sqWhere = sqWhere & " AND " & Rels(i)
    Else
        sqWhere = "WHERE " & Rels(i)
    End If
    End If
    
    Next i
    
    'on error Resume Next
    sqFrom = Left(sqFrom, Len(sqFrom) - 1)
    sqWhat = Left(sqWhat, Len(sqWhat) - 1)
    
    
    Set fsoTextStream = fsoObject.OpenTextFile(iProjectRootPath & "\" & iTableName & "\search.asp", ForReading)
    Set fsoTextStreamTemp = fsoObject.CreateTextFile(iProjectRootPath & "\" & iTableName & "\searchtemp.asp", True)
     
    'on error Resume Next
    IncStr = "<!-- #include file =" & QT("inc_fieldmap.inc") & " -->" & vbCrLf
    
    
    While Not fsoTextStream.AtEndOfStream
        
        CurrentLine = fsoTextStream.ReadLine
        fsoTextStreamTemp.WriteLine CurrentLine
        
        If CurrentLine = "<!--$Definition -->" Then
                fsoTextStreamTemp.WriteLine (IncStr)
                fsoTextStreamTemp.WriteLine ("")
                fsoTextStreamTemp.WriteLine ("<%")
                fsoTextStreamTemp.WriteLine ("FieldCount=" & Fieldcount)
                fsoTextStreamTemp.WriteLine ("sqWhat =" & QT(sqWhat))
                fsoTextStreamTemp.WriteLine ("sqFrom =" & QT(sqFrom))
                fsoTextStreamTemp.WriteLine ("sqWhere =" & QT(sqWhere))
                fsoTextStreamTemp.WriteLine ("%>")
        End If
    
    Wend
    
    Call fsoTextStream.Close
    Call fsoTextStreamTemp.Close
    
    Call fsoObject.DeleteFile(iProjectRootPath & "\" & iTableName & "\search.asp", True)
    Call fsoObject.MoveFile(iProjectRootPath & "\" & iTableName & "\searchtemp.asp", iProjectRootPath & "\" & iTableName & "\search.asp")
    
End Function

Private Function CreateDeleteTable(iProjectRootPath, iProjectID, iTableName) As Boolean

    'this method creates a page that will delete a record from a table.
    
    Dim Datatype As Integer
    Dim attrib As Long
    Dim FieldName As String
    Dim fsoTextStream As TextStream
    Dim fsoTextStreamTemp As TextStream
    Dim fsoObject As New FileSystemObject
    Dim CurrentLine As String
    Dim Required As String
    Dim Star As String
    Dim FieldNameVariable As String
    Dim QueryCheck As String
    
    attrib = FIELDPROP(1)
    Datatype = FIELDPROP(2)
    FieldName = FIELDPROP(12)
    Required = FIELDPROP(8)
    
    Set fsoTextStream = fsoObject.OpenTextFile(iProjectRootPath & "\" & iTableName & "\delete.asp", ForReading)
    Set fsoTextStreamTemp = fsoObject.CreateTextFile(iProjectRootPath & "\" & iTableName & "\deletetemp.asp", True)
    
    While Not fsoTextStream.AtEndOfStream
        
        CurrentLine = fsoTextStream.ReadLine
        fsoTextStreamTemp.WriteLine CurrentLine
        
        If CurrentLine = "<!--$Definition -->" Then
            
            FieldNameVariable = Replace(FieldName, " ", "")
            FieldNameVariable = "var" & FieldNameVariable
                    
            QueryCheck = "Q" & FieldNameVariable
            
            fsoTextStreamTemp.WriteLine ("<%")
            
            fsoTextStreamTemp.WriteLine (QueryCheck & " = request.querystring(""" & FieldName & """)")
            fsoTextStreamTemp.WriteLine ("If " & QueryCheck & " = ""Yes"" then")
            fsoTextStreamTemp.WriteLine ("QueryCheckString = QueryCheckString & """ & FieldName & """ & ""=Yes&""")
            fsoTextStreamTemp.WriteLine ("end if")
            
            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.querystring(""dat" & FieldName & """)")
            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = replace(" & FieldNameVariable & ", ""'"", ""''"")")
            fsoTextStreamTemp.WriteLine ("ReDim Preserve arrField(totFieldCount)")
            fsoTextStreamTemp.WriteLine ("arrField(totFieldCount) = """ & FieldName & """")
            fsoTextStreamTemp.WriteLine ("if " & FieldNameVariable & " <> """" then")
            Select Case Datatype
                Case 10, 12        'Text, Memo, Hyperlink
                    
                    fsoTextStreamTemp.WriteLine ("if totfieldcount <> 0 then qString = qString & "" and""")
                    fsoTextStreamTemp.WriteLine ("qString = qString & "" [" & FieldName & "] = '"" & " & FieldNameVariable & " & ""'""")
                    fsoTextStreamTemp.WriteLine ("totFieldCount = totFieldCount + 1")
                Case 3, 4, 2, 6, 7, 15, 20, 5, 1, 100      'Number, Currency, YesNo
                    fsoTextStreamTemp.WriteLine ("if totfieldcount <> 0 then qString = qString & "" and""")
                    fsoTextStreamTemp.WriteLine ("qString = qString & "" [" & FieldName & "] = "" & " & FieldNameVariable & "")
                    fsoTextStreamTemp.WriteLine ("totFieldCount = totFieldCount + 1")
                Case 8             'Date
                    fsoTextStreamTemp.WriteLine ("if totfieldcount <> 0 then qString = qString & "" and""")
                    fsoTextStreamTemp.WriteLine ("qString = qString & "" [" & FieldName & "] = cdate('"" & " & FieldNameVariable & " & ""')""")
                    fsoTextStreamTemp.WriteLine ("totFieldCount = totFieldCount + 1")

                Case 11            'OLEOBject
                                   'is not supported
            End Select
            fsoTextStreamTemp.WriteLine ("end if")
            
            fsoTextStreamTemp.WriteLine ("%>")

        End If
    Wend
    
    Call fsoTextStream.Close
    Call fsoTextStreamTemp.Close
    
    Call fsoObject.DeleteFile(iProjectRootPath & "\" & iTableName & "\delete.asp", True)
    Call fsoObject.MoveFile(iProjectRootPath & "\" & iTableName & "\deletetemp.asp", iProjectRootPath & "\" & iTableName & "\delete.asp")
    
End Function

Private Function CreateUpdateTablePostHTML(iProjectRootPath, iProjectID, iTableName, Sch As Schema) As Boolean

    'this method builds the asp page that will disply text boxes / textareas / checkboxes
    'that will display current data that is in a given record. this data can be changed
    'by the end user and the new data will be posted to the updatecheck.asp page where the
    'specified record will be updated in the database.
    
    Dim Datatype As Integer
    Dim attrib As Long
    Dim FieldName As String, FieldCaption
    Dim fsoTextStream As TextStream
    Dim fsoTextStreamTemp As TextStream
    Dim fsoObject As New FileSystemObject
    Dim CurrentLine As String
    Dim Required As String
    Dim StarRequired As String
    Dim PrimaryKey As Boolean
    Dim StarPrimary As String
    Dim valueFieldName As String
    Dim QueryCheck As String
    Dim FieldNameVariable As String
    Dim Table, Field
    
    Dim Control, Caption
    
    Table = FIELDPROP(16)
    Field = FIELDPROP(15)
    
    FieldName = ""
    FieldName = FIELDPROP(12)
    FieldCaption = ""
    FieldCaption = FIELDPROP(14)
    
    If Trim(FieldCaption) = "" Then FieldCaption = FIELDPROP(12)
    
    
    attrib = FIELDPROP(1)
    Datatype = FIELDPROP(2)
    FieldName = FIELDPROP(12)
    Required = FIELDPROP(8)
    PrimaryKey = FIELDPROP(13)
    
    Set fsoTextStream = fsoObject.OpenTextFile(iProjectRootPath & "\" & iTableName & "\update.asp", ForReading)
    Set fsoTextStreamTemp = fsoObject.CreateTextFile(iProjectRootPath & "\" & iTableName & "\updateTemp.asp", True)
    
    While Not fsoTextStream.AtEndOfStream
        
        CurrentLine = fsoTextStream.ReadLine
        fsoTextStreamTemp.WriteLine CurrentLine
        
        
        If CurrentLine = "<!--$Definition -->" Then
            
            valueFieldName = Replace(FieldName, " ", "")
            valueFieldName = valueFieldName
            
            FieldNameVariable = Replace(FieldName, " ", "")
            FieldNameVariable = "var" & FieldNameVariable
            QueryCheck = "Q" & valueFieldName
            
            
            If VBA.LCase(Required) = "true" Then
                StarRequired = "<font color=""#FF0000"">*</font>"
            Else
                StarRequired = ""
            End If
            
            StarPrimary = ""
            Control = ""
            Caption = ""
            
            Select Case Datatype
                Case 10            '     Text
                    Caption = ("<p>" & FieldCaption & StarRequired & StarPrimary & ":<br>")
                    Control = ("<input type=""text"" name=""" & FieldName & """ size=""20"" value=""<%=adoRecordset(""" & valueFieldName & """)%>""></p>")
                    
                Case 100            '     Relation Number
                    Caption = ("<p>" & FieldCaption & StarRequired & StarPrimary & ":<br>")
                    Control = ("<%WriteCatCombo " & QT(FieldName) & "," & QT(Field) & "," & QT(GetNextField(Table, Field, Sch)) & "," & QT(Table) & ", adoRecordset(""" & valueFieldName & """) %></p>")
                    
                Case 12            '     Memo
                    Select Case attrib
                        Case 2     '     memo
                            Caption = ("<p>" & FieldCaption & StarRequired & StarPrimary & ":<br>")
                            Control = ("<textarea rows=""4"" name=""" & FieldName & """ cols=""40""><%=adoRecordset(""" & valueFieldName & """)%></textarea></p>")
                        Case 32770 'hyperlink
                            Caption = ("<p>" & FieldCaption & StarRequired & StarPrimary & ":<br>")
                            Control = ("<%if adoRecordset(""" & valueFieldName & """) <> """" then%>") & vbCrLf & _
                            ("<input type=""text"" name=""" & FieldName & """ size=""20"" value=""<%=mid(adoRecordset(""" & valueFieldName & """),1,(len(adoRecordset(""" & valueFieldName & """))/2)-1)%>""></p>") & vbCrLf & _
                            ("<%else%>") & vbCrLf & _
                            ("<input type=""text"" name=""" & FieldName & """ size=""20""></p>") & vbCrLf & _
                            ("<%end if%>") & vbCrLf
                    End Select
                Case 3, 4, 2, 6, 7, 15, 20
                    Select Case attrib
                        Case 17    'Autonumber
                                   'do nothing
                            
                        Case Else     '    Number
                            Caption = ("<p>" & FieldCaption & StarRequired & StarPrimary & ":<br>")
                            Control = ("<input type=""text"" name=""" & FieldName & """ size=""20"" value=""<%=adoRecordset(""" & valueFieldName & """)%>""></p>")
                            
                    End Select
                Case 8             '  DateTime
                    Caption = ("<p>" & FieldCaption & StarRequired & StarPrimary & ":<br>")
                    Control = ("<input type=""text"" name=""" & FieldName & """ size=""20"" value=""<%=adoRecordset(""" & valueFieldName & """)%>""></p>")
                Case 5             '  Currency
                    Caption = ("<p>" & FieldCaption & StarRequired & StarPrimary & ":<br>")
                    Control = ("<input type=""text"" name=""" & FieldName & """ size=""20"" value=""<%=adoRecordset(""" & valueFieldName & """)%>""></p>")
                Case 1             '     YesNo
                    
                    Caption = ("<p>" & FieldCaption & StarRequired & StarPrimary & ":<br>")
                    Control = ("<%if adoRecordset(""" & valueFieldName & """) <> ""False"" then%>") & vbCrLf & _
                    ("Yes <input type=""radio"" value=""Yes"" checked name=""" & FieldName & """> No <input type=""radio"" name=""" & FieldName & """ value=""No""></p>") & vbCrLf & _
                    ("<%else%>") & vbCrLf & _
                    ("Yes <input type=""radio"" value=""Yes"" name=""" & FieldName & """> No <input type=""radio"" name=""" & FieldName & """ value=""No"" checked></p>") & vbCrLf & _
                    ("<%end if%>")
                Case 11            'OLEOBject
                                   'is not supported
            End Select
            
            If Control <> "" And Caption <> "" Then
                    fsoTextStreamTemp.WriteLine "<tr>"
                    fsoTextStreamTemp.WriteLine "<td width=" & QT("30%") & "><font size=" & QT("1") & ">" & Caption & "</font></td>"
                    fsoTextStreamTemp.WriteLine "<td width=" & QT("70%") & ">" & Control & "</td>"
                    fsoTextStreamTemp.WriteLine "</tr>"
            End If
            
        End If
    Wend
    
    Call fsoTextStream.Close
    Call fsoTextStreamTemp.Close
    
    Call fsoObject.DeleteFile(iProjectRootPath & "\" & iTableName & "\update.asp", True)
    Call fsoObject.MoveFile(iProjectRootPath & "\" & iTableName & "\updateTemp.asp", iProjectRootPath & "\" & iTableName & "\update.asp")
    
    
End Function

Private Function CreateUpdateTableASP(iProjectRootPath, iProjectID, iTableName) As Boolean

    'this method builds the updatecheck.asp page which is used to process the data that
    'has been sent to it by the update.asp page and update a given record in the database.
    Dim Datatype As Integer
    Dim attrib As Long
    Dim FieldName As String, FieldCaption As String
    Dim fsoTextStream As TextStream
    Dim fsoTextStreamTemp As TextStream
    Dim fsoObject As New FileSystemObject
    Dim CurrentLine As String
    Dim Required As String
    Dim Star As String
    Dim FieldNameVariable As String
    Dim PrimaryKey As Boolean
    
    attrib = FIELDPROP(1)
    Datatype = FIELDPROP(2)
    FieldName = FIELDPROP(12)
    Required = FIELDPROP(8)
    PrimaryKey = FIELDPROP(13)
    
    
    FieldCaption = ""
    FieldCaption = FIELDPROP(14)
    
    If Trim(FieldCaption) = "" Then FieldCaption = FIELDPROP(12)
    
    
    Set fsoTextStream = fsoObject.OpenTextFile(iProjectRootPath & "\" & iTableName & "\updatecheck.asp", ForReading)
    Set fsoTextStreamTemp = fsoObject.CreateTextFile(iProjectRootPath & "\" & iTableName & "\ASPupdateTemp.asp", True)
    
    While Not fsoTextStream.AtEndOfStream
        
        CurrentLine = fsoTextStream.ReadLine
        
        fsoTextStreamTemp.WriteLine CurrentLine
        
        If CurrentLine = "<!--$Definition -->" And attrib <> 17 And Datatype <> 11 Then
            
            FieldNameVariable = Replace(FieldName, " ", "")
            FieldNameVariable = "var" & FieldNameVariable
            
            fsoTextStreamTemp.WriteLine ("<%")
            
            If VBA.LCase(Required) = "true" Or PrimaryKey = True Then
                fsoTextStreamTemp.WriteLine ("If Request.Form(""" & FieldName & """) = """" Then")
                fsoTextStreamTemp.WriteLine ("Error=" & QT(FieldCaption & " cannot be empty"))
                fsoTextStreamTemp.WriteLine ("End If")
            End If
            
            Select Case Datatype
                Case 10            '     Text
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = replace(" & FieldNameVariable & ", ""'"", ""''"")")
                    fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = Replace(" & FieldNameVariable & ", ""'"", ""''"")")
                    
                    fsoTextStreamTemp.WriteLine ("if FirstHalfSQL = ""update [" & iTableName & "] set "" then")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "" [" & FieldName & "] = '"" & " & FieldNameVariable & " & ""'""")
                    fsoTextStreamTemp.WriteLine ("else")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "", [" & FieldName & "] = '"" & " & FieldNameVariable & " & ""'""")
                    fsoTextStreamTemp.WriteLine ("end if")
                    
                    fsoTextStreamTemp.WriteLine ("end if")
                Case 12            '     Memo
                    Select Case attrib
                        Case 2     '     memo
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                            fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = Replace(" & FieldNameVariable & ", ""'"", ""''"")")
                            
                            fsoTextStreamTemp.WriteLine ("if FirstHalfSQL = ""update [" & iTableName & "] set "" then")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "" [" & FieldName & "] = '"" & " & FieldNameVariable & " & ""'""")
                            fsoTextStreamTemp.WriteLine ("else")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "", [" & FieldName & "] = '"" & " & FieldNameVariable & " & ""'""")
                            fsoTextStreamTemp.WriteLine ("end if")
                            
                            fsoTextStreamTemp.WriteLine ("end if")
                        Case 32770 'hyperlink
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                            fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = Replace(" & FieldNameVariable & ", ""'"", ""''"")")
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = " & FieldNameVariable & " & ""#"" & " & FieldNameVariable & " & ""#""")
                            
                            fsoTextStreamTemp.WriteLine ("if FirstHalfSQL = ""update [" & iTableName & "] set "" then")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "" [" & FieldName & "] = '"" & " & FieldNameVariable & " & ""'""")
                            fsoTextStreamTemp.WriteLine ("else")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "", [" & FieldName & "] = '"" & " & FieldNameVariable & " & ""'""")
                            fsoTextStreamTemp.WriteLine ("end if")
                            
                            fsoTextStreamTemp.WriteLine ("end if")
                    End Select
                Case 3, 4, 2, 6, 7, 15, 20, 100
                    Select Case attrib
                        Case 17    'Autonumber
                                   'do nothing
                        Case Else     '    Number
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                            fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                            fsoTextStreamTemp.WriteLine ("'on error resume next")
                            fsoTextStreamTemp.WriteLine (FieldNameVariable & " = CDbl(" & FieldNameVariable & ")")
                            fsoTextStreamTemp.WriteLine ("If Err.Number = 13 Then")
                            fsoTextStreamTemp.WriteLine ("Error=" & QT("The data you entered in the field " & FieldCaption & " is invalid"))
                            fsoTextStreamTemp.WriteLine ("End If")
                            fsoTextStreamTemp.WriteLine ("Err.Clear")
                            fsoTextStreamTemp.WriteLine ("''on error GoTo 0")
                            
                            fsoTextStreamTemp.WriteLine ("if FirstHalfSQL = ""update [" & iTableName & "] set "" then")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "" [" & FieldName & "] = "" & " & FieldNameVariable)
                            fsoTextStreamTemp.WriteLine ("else")
                            fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "", [" & FieldName & "] = "" & " & FieldNameVariable)
                            fsoTextStreamTemp.WriteLine ("end if")
                            
                            fsoTextStreamTemp.WriteLine ("end if")
                    End Select
                Case 8             '  DateTime
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                    fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                    fsoTextStreamTemp.WriteLine ("'on error resume next")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = CDate(" & FieldNameVariable & ")")
                    fsoTextStreamTemp.WriteLine ("If Err.Number = 13 Then")
                    fsoTextStreamTemp.WriteLine ("Error=" & QT("The data you entered in the field " & FieldCaption & " is invalid"))
                    fsoTextStreamTemp.WriteLine ("End If")
                    fsoTextStreamTemp.WriteLine ("Err.Clear")
                    fsoTextStreamTemp.WriteLine ("''on error GoTo 0")
                    
                    fsoTextStreamTemp.WriteLine ("if FirstHalfSQL = ""update [" & iTableName & "] set "" then")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "" [" & FieldName & "] = cdate('"" & " & FieldNameVariable & " & ""')""")
                    fsoTextStreamTemp.WriteLine ("else")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "", [" & FieldName & "] = cdate('"" & " & FieldNameVariable & " & ""')""")
                    fsoTextStreamTemp.WriteLine ("end if")
                    
                    fsoTextStreamTemp.WriteLine ("end if")
                    
                Case 5             '  Currency
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                    fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " <> """" then")
                    fsoTextStreamTemp.WriteLine ("'on error resume next")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = CCur(" & FieldNameVariable & ")")
                    fsoTextStreamTemp.WriteLine ("If Err.Number = 13 Then")
                    fsoTextStreamTemp.WriteLine ("Error=" & QT("The data you entered in the field " & FieldCaption & " is invalid"))
                    fsoTextStreamTemp.WriteLine ("End If")
                    fsoTextStreamTemp.WriteLine ("Err.Clear")
                    fsoTextStreamTemp.WriteLine ("''on error GoTo 0")
                    
                    fsoTextStreamTemp.WriteLine ("if FirstHalfSQL = ""update [" & iTableName & "] set "" then")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "" [" & FieldName & "] = "" & " & FieldNameVariable)
                    fsoTextStreamTemp.WriteLine ("else")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "", [" & FieldName & "] = "" & " & FieldNameVariable)
                    fsoTextStreamTemp.WriteLine ("end if")
                    
                    fsoTextStreamTemp.WriteLine ("end if")
                Case 1             '     YesNo
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = request.form(""" & FieldName & """)")
                    fsoTextStreamTemp.WriteLine ("If " & FieldNameVariable & " = ""Yes"" then")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = True")
                    fsoTextStreamTemp.WriteLine ("Else")
                    fsoTextStreamTemp.WriteLine (FieldNameVariable & " = False")
                    fsoTextStreamTemp.WriteLine ("End If")
                     
                    fsoTextStreamTemp.WriteLine ("if FirstHalfSQL = ""update [" & iTableName & "] set "" then")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "" [" & FieldName & "] = "" & " & FieldNameVariable)
                    fsoTextStreamTemp.WriteLine ("else")
                    fsoTextStreamTemp.WriteLine ("FirstHalfSQL = FirstHalfSQL & "", [" & FieldName & "] = "" & " & FieldNameVariable)
                    fsoTextStreamTemp.WriteLine ("end if")
                
                Case 11            'OLEOBject
                                   'is not supported
            End Select
            
            fsoTextStreamTemp.WriteLine ("%>")
        End If
    Wend
    
    Call fsoTextStream.Close
    Call fsoTextStreamTemp.Close
    
    Call fsoObject.DeleteFile(iProjectRootPath & "\" & iTableName & "\updatecheck.asp", True)
    Call fsoObject.MoveFile(iProjectRootPath & "\" & iTableName & "\ASPupdateTemp.asp", iProjectRootPath & "\" & iTableName & "\updatecheck.asp")
    
End Function

Private Function CheckGoodDatabase(iDbPath) As Boolean

    'function checks to see if database is valid by
    'making a connection to it and then disconnecting
    
    Dim objDAO As New DAO.DBEngine
    Dim objDatabase As DAO.Database
    Dim DBPassword As String
    Dim wrkJet As DAO.Workspace
    
    ''on error GoTo ErrorHandler:
    
    Set objDatabase = objDAO.OpenDatabase(iDbPath)
    
    CheckGoodDatabase = True
    
    Exit Function
    
ErrorHandler:
    CheckGoodDatabase = False
    
End Function

Public Function NameFromPath(Path As String, Optional dLimiter As String = "\")

    'on error Resume Next
    Dim nametoshow() As String
      nametoshow = Split(Path, dLimiter)
      nametoshow(0) = nametoshow(UBound(nametoshow))
      NameFromPath = nametoshow(0)

End Function


'==========================================================================================
Public Function ReadFile(FileName)
'==========================================================================================

Dim MyS As String, TotS As String
ReadFile = ""
TotS = ""
MyS = ""


'on error Resume Next

Open FileName For Input As #1
If Err Then GoTo NoOpen
    Do While Not EOF(1)
        Line Input #1, MyS
        TotS = TotS & MyS & vbCrLf
    Loop
Close #1

NoOpen:
 
ReadFile = TotS

End Function


'==========================================================================================
Public Function WriteFile(FileName, Str) As Boolean
'==========================================================================================

''on error GoTo NoOp
Open FileName For Output As #1
Print #1, Str
Close #1
WriteFile = True
Exit Function

NoOp:
WriteFile = False

End Function



'==========================================================================================
Public Function WriteCodeFile(ToFile, CoreFile, TemplateFile, EntityName, Optional IfTable As Boolean = False)
'==========================================================================================
'The function for interpreting the code file

Dim ToStr, CoreStr, TempStr, HTMLStr, ASPStr, IncStr

'on error Resume Next

ToStr = "": CoreStr = "": TempStr = "": HTMLStr = "": ASPStr = "": IncStr = ""

CoreStr = ReadFile(CoreFile)
TempStr = ReadFile(TemplateFile)

CoreStr = Replace(CoreStr, "$ENTITY", EntityName)


'Get the two parts of our ASP template file
SplitString CoreStr, "$SAFECODE", ASPStr, HTMLStr


If IfTable = False Then
    TempStr = Replace(TempStr, "$TITLE", "<%=VAR_TITLE%>")
    TempStr = Replace(TempStr, "$NAVBAR", UCase(VAR_NAVBAR))
    TempStr = Replace(TempStr, "$SUBTITLE", "<%=VAR_SUBTITLE%>")
    TempStr = Replace(TempStr, "$SUBNAVBAR", "<%=VAR_SUBNAVBAR%>")
    TempStr = Replace(TempStr, "$PANE", "<%=var_pane%>")
    TempStr = Replace(TempStr, "$COPYRIGHT", "<%=var_copyright%>")
Else
    TempStr = Replace(TempStr, "$TITLE", "Entity Name")
    TempStr = Replace(TempStr, "$NAVBAR", "")
    TempStr = Replace(TempStr, "$SUBTITLE", "All Tables")
    TempStr = Replace(TempStr, "$SUBNAVBAR", "")
    TempStr = Replace(TempStr, "$PANE", "")
    TempStr = Replace(TempStr, "$COPYRIGHT", "")
End If



If IfTable = False Then
    IncStr = "<!-- #include file =" & QT("..\inc\inc_global.inc") & " -->" & vbCrLf
    IncStr = IncStr & "<!-- #include file =" & QT("inc_local.inc") & " -->" & vbCrLf
Else
    IncStr = "<!-- #include file =" & QT("inc\inc_global.inc") & " -->" & vbCrLf
End If

IncStr = IncStr & GetCommentBlock(True) & vbCrLf


'Create the new file
If Trim(HTMLStr) <> "" Then
        ToStr = IncStr & ASPStr & vbCrLf & Replace(TempStr, "$CONTENT", HTMLStr)
    Else
        ToStr = IncStr & ASPStr & vbCrLf
End If


WriteFile ToFile, ToStr

End Function

'==========================================================================================
Function SplitString(Str, Delim, FirstPart, SecondPart)
'==========================================================================================

'Splits a string to two parts
    Dim Strs() As String, i, GiveStr
    
    GiveStr = ""
    
    If InStr(1, Str, Delim) = 0 Then
            FirstPart = ""
            SecondPart = Str
        Else
            Strs = Split(Str, Delim)
            FirstPart = Strs(0)
            For i = 1 To UBound(Strs())
              GiveStr = GiveStr & Strs(i)
            Next
            SecondPart = GiveStr
    End If
    
    
End Function


Function GetLink(Text, Link, Optional Color = "")
If Trim(Color) = "" Then
    GetLink = "<a href=" & QT(Link) & ">" & Text & "</a>"
Else
    GetLink = "<a href=" & QT(Link) & "><font color=" & QT(Color) & ">" & Text & "</font></a>"
End If

End Function


Function QT(Text)
    QT = """" & Text & """"
End Function

Function WriteInclude(Path, iDbPath, iCorePath)

'on error Resume Next
MkDir Path & "\inc"

    Dim IncStr As String
    
    IncStr = ""
    IncStr = IncStr & "<%@ Language=VBScript %>" & vbCrLf & vbCrLf
    IncStr = IncStr & "<%" & vbCrLf
    IncStr = IncStr & "   '-----------------------------------------------------" & vbCrLf
    IncStr = IncStr & "   'Entity Manager :" & vbCrLf
    IncStr = IncStr & "   '-----------------------------------------------------" & vbCrLf
    IncStr = IncStr & "   'Global Include File" & vbCrLf
    IncStr = IncStr & "   '-----------------------------------------------------" & vbCrLf & vbCrLf
    IncStr = IncStr & "         Dim VAR_TITLE, VAR_CONSTRING,VAR_NAVBAR" & vbCrLf
    
    IncStr = IncStr & "         VAR_TITLE=" & QT(VAR_TITLE) & vbCrLf
    IncStr = IncStr & "         VAR_CONSTRING =" & QT("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & iDbPath) & vbCrLf & "%>" & vbCrLf
    
    
    IncStr = IncStr & ReadFile(iCorePath & "\inc\inc_global.inc")
    
    
WriteFile Path & "\inc\" & "INC_GLOBAL.INC", IncStr

End Function

Function WriteTableInclude(Path, TableName)

    Dim IncStr
    Dim VAR_SUBNAVBAR
    
    IncStr = "<%" & vbCrLf & "Dim VAR_SUBTITLE, VAR_TABLE" & vbCrLf
    IncStr = IncStr & "VAR_SUBTITLE=" & QT(TableName) & vbCrLf
    IncStr = IncStr & "VAR_TABLE=" & QT(TableName) & vbCrLf & "%>"
    
    
    
    
    'IncStr = IncStr & "VAR_PRIMARY=" & QT(PrimaryKey) & vbCrLf & "%>"
    
    WriteFile Path & "\" & "INC_LOCAL.INC", IncStr
    
    
End Function

Function WriteVarsInclude(Path)

End Function

Function CreateFieldMap(Table, Field, TableTo, FieldTo, Caption, Visible, Index, IfPrimary, Sch As Schema)

Dim LinkDisplay, LinkValue
Dim FieldName, TableName, AsN1, AsN2

'on error Resume Next

If Caption = "" Then Caption = Field

LinkDisplay = ""
LinkValue = ""


'If Table = "Order Details" Then MsgBox "hello"


If TableTo <> "" And FieldTo <> "" Then


    LinkValue = "[" & TableTo & "].[" & FieldTo & "]"
    AsN1 = TableTo & FieldTo
    
    AddRel "[" & Table & "].[" & Field & "]", "[" & TableTo & "].[" & FieldTo & "]"
    
    Dim RetFld
    RetFld = ""
    RetFld = GetNextField(TableTo, FieldTo, Sch)
    
    If RetFld <> "" Then
        LinkDisplay = "[" & TableTo & "].[" & RetFld & "]"
        AsN2 = TableTo & RetFld
    Else
        LinkDisplay = "[" & TableTo & "].[" & FieldTo & "]"
        AsN2 = TableTo & FieldTo
    End If
    
    AddTable "[" & TableTo & "]"
    
    AsN1 = Replace(AsN1, " ", "")
    AsN2 = Replace(AsN2, " ", "")
    
    AddField LinkValue, AsN1
    AddField LinkDisplay, AsN2
    
Else
    LinkValue = "[" & Table & "].[" & Field & "]"
    AsN1 = Table & Field
    AsN1 = Replace(AsN1, " ", "")
    LinkDisplay = LinkValue
    AsN2 = AsN1
    AddField LinkValue, AsN1

End If

    AddField "[" & Table & "].[" & Field & "]", Replace(Table & Field, " ", "")
    AddTable "[" & Table & "]"


'Builds the FieldMap string
    gMAPSTR = gMAPSTR & vbCrLf
    gMAPSTR = gMAPSTR & "'=============================================================='" & vbCrLf
    gMAPSTR = gMAPSTR & "' Mapstring for " & Field & vbCrLf
    gMAPSTR = gMAPSTR & "'=============================================================='" & vbCrLf & vbCrLf
    gMAPSTR = gMAPSTR & "FieldMap (" & Index & ",0) = " & QT("[" & Table & "].[" & Field & "]") & vbCrLf
    gMAPSTR = gMAPSTR & "FieldMap (" & Index & ",1) = " & QT(TableTo) & vbCrLf
    gMAPSTR = gMAPSTR & "FieldMap (" & Index & ",2) = " & QT(AsN2) & vbCrLf
    gMAPSTR = gMAPSTR & "FieldMap (" & Index & ",3) = " & QT(LinkValue) & vbCrLf
    gMAPSTR = gMAPSTR & "FieldMap (" & Index & ",4) = " & QT(Caption) & vbCrLf
    gMAPSTR = gMAPSTR & "FieldMap (" & Index & ",5) = " & QT(CStr(Visible)) & vbCrLf
    gMAPSTR = gMAPSTR & "FieldMap (" & Index & ",6) = " & QT(CStr(IfPrimary)) & vbCrLf
    gMAPSTR = gMAPSTR & "FieldMap (" & Index & ",7) = " & QT(LinkDisplay) & vbCrLf
      
    
End Function

Sub AddRel(RFrom, RTo)
Dim i, RelName
RelName = ""
RelName = RFrom & "=" & RTo

On Error Resume Next
Rels.Add RelName, ReplaceChars(RelName)


End Sub


Sub AddTable(TableName)
Dim Txt

Txt = TableName

On Error Resume Next
Tables.Add Txt, ReplaceChars(Txt)


End Sub

Sub AddField(FieldName, AsName)
Dim Txt
Txt = FieldName & " As " & AsName

On Error Resume Next
Fields.Add Txt, ReplaceChars(Txt)


End Sub

Function ReplaceChars(ByVal Txt)
Txt = Replace(Txt, "[", "")
Txt = Replace(Txt, "]", "")
Txt = Replace(Txt, " ", "")
Txt = Replace(Txt, ".", "")
Txt = Replace(Txt, "=", "")
ReplaceChars = Txt
End Function


Function GetNextField(Table, Field, Sch As Schema)
Dim i, K
On Error GoTo NoField
Dim thisRel As SchemaModel.Relation

For Each thisRel In Sch.SchemaRelations

        If VBA.LCase(thisRel.RelationAttributes("field").AttribValue) = VBA.LCase(Field) And VBA.LCase(thisRel.RelationAttributes("table").AttribValue) = VBA.LCase(Table) Then
            GetNextField = thisRel.RelationAttributes("displayfield")
            Exit Function
        End If
Next

NoField:

If Err Then MsgBox Err.Description
GetNextField = ""

End Function



Function GetCommentBlock(Optional IfAsp As Boolean = True)

    HeaderString = "<%" & vbCrLf
    HeaderString = HeaderString & "'-------------------------------------------------------" & vbCrLf
    HeaderString = HeaderString & "'NET ENGINE CODE - VERSION 1.0" & vbCrLf
    HeaderString = HeaderString & "'-------------------------------------------------------" & vbCrLf
    HeaderString = HeaderString & "'Code Generated By Net Engine : Developed By Anoop" & vbCrLf
    HeaderString = HeaderString & "'http://www.logicmatrixonline.com/anoop" & vbCrLf
    HeaderString = HeaderString & "'-------------------------------------------------------" & vbCrLf
    HeaderString = HeaderString & "%>" & vbCrLf

    If Not IfAsp Then HeaderString = Replace(HeaderString, "<%", "<!--")
    If Not IfAsp Then HeaderString = Replace(HeaderString, "%>", "!-->")
    

End Function


Public Function STDT(Str, Datatype, attrib)
'on error Resume Next

'String To Data Type
Select Case LCase(Str)
    Case "text"
        Datatype = 10
        attrib = 2
    Case "relation"
        Datatype = 100
        attrib = 0
    Case "memo"
        Datatype = 12
        attrib = 2
    Case "hyperlink"
        Datatype = 12
        attrib = 32779
    Case "number"
        Datatype = 20
        attrib = 0
    Case "autonumber"
        Datatype = 20
        attrib = 17
    Case "datetime"
        Datatype = 8
        attrib = 0
    Case "currency"
        Datatype = 5
        attrib = 0
    Case "yesno"
        Datatype = 1
        attrib = 0
    Case "ole"
        Datatype = 11
        attrib = 0
End Select

End Function


Function GetBoolean(Expr)
    If LCase(Expr) = "true" Then
        GetBoolean = True
    Else
        GetBoolean = False
    End If
End Function
