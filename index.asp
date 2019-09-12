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
    strGreeting = "Hello and Good Afternoon this is a test for ben!" 
    ElseIf ((dtmHour > 17) and (dtmHour < 19)) Then
    strGreeting = "Hello and Good Afternoon, just thought I would let you know that night is approaching and it will be dark soon"
  Else   
    strGreeting = "Hello and Good Evening!" 
  End If    
%>  
<%= strGreeting %>
<% Response.Write ("    Test hello world message, the hour now is   ") %>
<%= dtmHour %>
        </body>
</html>