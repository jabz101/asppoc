<%@ Language="VBScript" %>
 <html>
        <head> </head>
        <body>
            <h1>VBScript Powered ASP page on Azure Web Apps </h1>

<%  
Dim dtmHour
  
  dtmHour = Hour(Now())
  
  If dtmHour < 12 Then 
    strGreeting = "Good Morning to you!" 
  ElseIf ((dtmHour > 12) and (dtmHour < 16)) Then
    strGreeting = "Hello and Good Afternoon" 
    ElseIf ((dtmHour > 17) and (dtmHour < 19)) Then
    strGreeting = "Hello and Good Afternoon, just thought I would let you know that night is approaching and it will be dark soon"
  Else   
    strGreeting = "Hello and Good Evening!" 
  End If    
%>  

<%= strGreeting %>
<% Response.Write("    Test hello world message, the date & time now is: ") %>
<% Response.Write(Now) %>

<BR>Click this to test out <a href ="northwind01.asp"> my first ASP Classic SQL DB page</a> 
<BR>
<BR>Click this to test out <a href ="northwind02.asp"> my second ASP Classic SQL DB page</a> 
<%
Set objWSH =  CreateObject("WScript.Shell")
Set objUserVariables = objWSH.Environment("DB") 
Response.Write(objUserVariables("SQLAZURECONNSTR_DB"))
%>
        </body>
</html>
