<%

'----------------------------------------------------------------------

'SQL Parts
Dim sqWhat,sqFrom,sqWhere,sqOrder

'SQL Command=Sum of SQL Parts
Dim sCommand

'Heading To Display
Dim sHeading

'----------------------------------------------------------------------

Dim pItem,p

%>


<%
pItem=Request.QueryString("Item")
pFrom=Request.QueryString("From")
response.redirect "update.asp?item=" & pitem 

%>