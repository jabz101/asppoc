<%@ Language="VBScript" %>
 <html>
        <head> </head>
        <body>
            <h1>VBScript Powered ASP page on Azure Web Apps </h1>
    <%  
  Dim dtmHour 

  dtmHour = Hour(Now()) 

  If dtmHour < 12 Then 
    strGreeting = "Good Morning!" 
  ElseIf ((dtmHour > 17) and (dtmHour < 19)) Then
    strGreeting = "Hello! and Good Afternoon, have a good night" 
  Else   
    strGreeting = "Hello! and Good Evening" 
  End If    
%>  

<%= strGreeting %> 
        </body>
</html>