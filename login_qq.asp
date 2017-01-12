<%
dim url
dim callback
dim appid

url = "https://graph.qq.com/oauth2.0/authorize?"
appid = "xxx"
callback = "http://domain/login_qq_callback.asp"

url = url & "client_id=" & appid & "&response_type=code&scope=all&redirect_uri=" &  Server.URLEncode(callback)

' response.write(url)
response.redirect(url)

%>