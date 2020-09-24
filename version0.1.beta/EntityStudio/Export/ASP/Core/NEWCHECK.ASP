<%on error resume next%>

<% 

Dim Error,USuccess

Error=""
USuccess=True

FromString = request.querystring("From")

Set adoConnection = server.CreateObject("ADODB.Connection")
Set adoRecordset = server.CreateObject("ADODB.Recordset")
adoConnection.Open (VAR_CONSTRING)

iFieldCount = 0
FirstHalfSQL = "insert into [$ENTITY] ("
SecondHalfSQL = ") Values ("
EndSQL = ")"%>
<!--$Definition -->

<% adoRecordset.ActiveConnection = AdoConnection
SQLInsert = FirstHalfSQL & SecondHalfSQL & EndSQL

if SQLInsert <> "insert into [$ENTITY] () Values ()" then

If Trim(Error)<>"" then 
USuccess=False
End If

If USuccess=True Then
Call adoRecordset.Open(SQLInsert)

	if err then 
	Error=Err.Description 
	USuccess=False
	end if

End If

ELSE

	Error="Unable to update. No input found"


End If

%>


$SAFECODE

<% if trim(error)="" then %>

<table border="0" cellpadding="2" width="100%" height="105">
  <tr>
    <td width="12%" rowspan="4" height="99" valign="top"></td>
    <td width="88%" height="18"><b><font color="#008000" face="verdana,arial" size="2">(i)</font><font face="verdana,arial" size="2" color="#FF0000">
      </font></b><font size="2" face="Verdana"><b>The new item is added to
      $ENTITY : Thank You</b></font></td>
  </tr>
  <tr>
    <td width="88%" height="21"></td>
  </tr>
  <tr>
    <td width="88%" height="22"><b><font face="Verdana" size="2">What would you like to do?</font></b></td>
  </tr>
  <tr>
    <td width="88%" height="24">
      <blockquote>
        <ul>
          <li><font face="Verdana" size="2"><a href="list.asp">Continue browsing
            $ENTITY</a></font></li>
          <li><font face="Verdana" size="2"><a href="new.asp?<%=FromString%>">Add another item</a></font></li>
        </ul>
      </blockquote>
    </td>
  </tr>
</table>

<%else%>

<table border="0" cellpadding="2" width="100%">
  <tr>
    <td width="12%" rowspan="4" valign="top"></td>
    <td width="88%" height="22"><b><font face="verdana,arial" size="2" color="#FF0000">(!)
      </font><font face="verdana,arial" size="2">An error
      occurred while trying to save the information.</font></b></td>
  </tr>
  <tr>
    <td width="88%" height="22"><b><font face="verdana,arial" size="2" color="#FF0000">
      <%=Error%></font></b></td>
  </tr>
  <tr>
    <td width="88%" height="22"><b><font face="verdana,arial" size="2">What do
      you want to do?</font></b></td>
  </tr>
  <tr>
    <td width="88%">
      <blockquote>
        <ul>
          <li><font face="verdana,arial" size="2"><a href="javascript:history.back()">Go back and correct this
            problem</a></font></li>
          <li><font face="verdana,arial" size="2"><a href="<%=FromString%>">Continue with out making any
            changes</a></font></li>
        </ul>
      </blockquote>
    </td>
  </tr>
</table>


<% end if %>