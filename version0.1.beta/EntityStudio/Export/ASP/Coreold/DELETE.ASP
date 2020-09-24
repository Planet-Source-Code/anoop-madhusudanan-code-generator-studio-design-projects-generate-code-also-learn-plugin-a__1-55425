<%on error resume next%>

<%
Dim Error
Dim ItemStr,FromStr
ItemStr = Request.QueryString("Item") 
FromStr = Request.QueryString("From") %>
<%


Set adoConnection = server.CreateObject("ADODB.Connection")
Set adoRecordset = server.CreateObject("ADODB.Recordset")
adoConnection.Open (VAR_CONSTRING)
adoRecordset.open "delete * from [$ENTITY] where " & ItemStr,adoConnection

If Err then Error=err.Description 

Error=replace(error,"table","")
Error=replace(error,"record","item")

%>


$SAFECODE


<% if trim(error)="" then %>

<table border="0" cellpadding="2" width="100%" height="105">
  <tr>
    <td width="12%" rowspan="4" height="99" valign="top"></td>
    <td width="88%" height="18"><b><font color="#008000" face="verdana,arial" size="2">(i)</font><font face="verdana,arial" size="2" color="#FF0000">
      </font></b><font size="2" face="Verdana">
      <b>The Requested Item Is Deleted.</b></font></td>
  </tr>
  <tr>
    <td width="88%" height="21"></td>
  </tr>
  <tr>
    <td width="88%" height="22"><b><font face="Verdana" size="2">What would you like to do next?</font></b></td>
  </tr>
  <tr>
    <td width="88%" height="24">
      <blockquote>
        <ul>
          <li><font face="Verdana" size="2"><a href="list.asp">Continue browsing
            $ENTITY</a></font></li>
          <li><font face="Verdana" size="2"><a href="new.asp?<%=FromString%>">Add A New Item To $ENTITY</a></font></li>
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
      occurred while trying to delete the item.</font></b></td>
  </tr>
  <tr>
    <td width="88%" height="22"><b><font face="verdana,arial" size="2" color="#FF0000">
      <%=Error%></font></b></td>
  </tr>
  <tr>
    <td width="88%" height="22"><b><font face="verdana,arial" size="2">You Can..</font></b></td>
  </tr>
  <tr>
    <td width="88%">
      <blockquote>
        <ul>
          <li><font face="verdana,arial" size="2"><a href="javascript:history.back()">Go Back With Out Making Any Changes</a></font></li>
          <li><font face="verdana,arial" size="2"><a href="<%=FromString%>">Browse The $ENTITY</a></font></li>
        </ul>
      </blockquote>
    </td>
  </tr>
</table>


<% end if %>

