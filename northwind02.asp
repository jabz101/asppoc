<!DOCTYPE html>
<html>
 <head>
	<title>Sample ASP page</title>
	<link href="css/main.css" rel="stylesheet" type="text/css"/>
 </head>
 <body>
<%

Set objWSH =  CreateObject("WScript.Shell")
Set objUserVariables = objWSH.Environment("%DB%") 
Response.Write(objUserVariables("SQLAZURECONNSTR_DB"))

	set conn=objUserVariables
	set rs = Server.CreateObject("ADODB.recordset")
	sql="SELECT * FROM SalesLT.Customer WHERE FirstName ='Keith'"
	rs.Open sql, conn



%>


<h1><img src="img/AspClassic.png" alt=""/> Classic ASP website with Northwind DB, SQL Query that searches for specfic person</h1>
<table border="2" width="100%">
<tr>
<%for each x in rs.Fields
    response.write("<th>" & x.name & "</th>")
next%>
</tr>
<%do until rs.EOF%>
    <tr>
    <%for each x in rs.Fields%>
       <td><%Response.Write(x.value)%></td>
    <%next
    rs.MoveNext%>
    </tr>
<%loop
rs.close
conn.close
%>
</table>
</body>
</html>