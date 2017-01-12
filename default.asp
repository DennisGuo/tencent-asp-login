<%
dim dom 
dom =  "<table border=1>"
for each x in Request.ServerVariables
    dom = dom & "<tr><td>" & x & " </td><td>" & Request.ServerVariables(x) & "</td></tr>"
next
dom  = dom & "</table><br/>"

' for each y in Server
'    dom = dom &  y & "<br/>"
' next

response.write(dom)

' Î¢ÐÅµÇÂ¼
%>
¡¡¡¡¡¡¡¡