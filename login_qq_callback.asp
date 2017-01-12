<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="GBK">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>QQ回调解析</title>
    <!-- 新 Bootstrap 核心 CSS 文件 -->
    <link rel="stylesheet" href="http://cdn.bootcss.com/bootstrap/3.3.0/css/bootstrap.min.css">
</head>
<body>
<div class="container">





<script language="jscript" runat="server">  
    Array.prototype.get = function(x) { return this[x]; };  
    function parseJSON(strJSON) { return eval("(" + strJSON + ")"); }
</script>  
<%


' 方法区

Private Function GetData(url)
    Dim request
    set request = Server.CreateObject("Msxml2.ServerXMLHTTP.3.0")
	request.setOption 2, 13056
    request.Open "GET", url, False
    request.send    
    GetData = request.ResponseText
    Set request = nothing
End Function 

Private Function JsonpToJson(str)
    str = Replace(str,"callback(","")
    str = Replace(str,");","")
    JsonpToJson = str
End Function 


Private Function GetToken(rs)
    ' access_token=BF9836A9EDE465FF3528E4327FF6E8C5&expires_in=7776000&refresh_token=256C7AF1C0AE010D3C5C24B39CB606F2
    GetToken = Mid(rs, 14, 32)
End Function 



' 方法区结束

dim appid
dim appsec
dim code 
dim token
dim url
dim state
dim callback

' 1. 取到CODE，通过CODE获取TOKEN，会再次回调到此地址，并且携带TOKEN参数
' 2. 取到TOKEN，通过TOKEN获取openid（TOKEN有效期3个月，建议缓存）
' 3. 使用openid获取用户信息

appid="xxx"
appsec="xxx"
state="qq_login_state" '可使用随机字符串
callback = "http://domain/login_qq_callback.asp"

code = request("code")


' 第一次取到code 注意此code会在10分钟内过期
if not isEmpty(code) then 
    url = "https://graph.qq.com/oauth2.0/token?"
    url = url & "grant_type=authorization_code&client_id="&appid&"&client_secret="&appsec&"&code="&code&"&state="&state&"&redirect_uri=" & Server.URLEncode(callback)
    html = html & "<h1>第一步：获取TOKEN</h1>"& url
    rs = GetData(url)
    ' access_token=BF9836A9EDE465FF3528E4327FF6E8C5&expires_in=7776000&refresh_token=256C7AF1C0AE010D3C5C24B39CB606F2
    token = GetToken(rs)

    ' 第二次获取token token具有3个月有效期
    if not isEmpty(token) then

        dim rs
        dim json
        dim openid
        dim ret
        dim html

        ' 获取openid
        url = "https://graph.qq.com/oauth2.0/me?access_token=" & token
        html = html & "<h1>第二步：获取用户openid</h1>"& url
        rs = GetData(url)
        html = html & "<hr/><h1>返回结果：</h1>"& rs
        ' callback( {“client_id”:”YOUR_APPID”,”openid”:”YOUR_OPENID”} ); 
        rs = JsonpToJson(rs)
        ' {“client_id”:”YOUR_APPID”,”openid”:”YOUR_OPENID”}
        html = html & "<hr/><h1>将jsonp结果替换为json结果</h1>"&rs

        set json = parseJSON(rs)
        openid = json.openid

        ' 获取用户信息
        url = "https://graph.qq.com/user/get_user_info?access_token="&token&"&oauth_consumer_key="&appid&"&openid="&openid
        html = html & "<hr/><h1>第三步：根据openid获取用户信息</h1>"& url

        rs = GetData(url)
        html = html & "<hr/><h1>返回信息（用户信息）</h1>"& rs
        ' {
        '     "ret":0,
        '     "msg":"",
        '     "nickname":"Peter",
        '     "figureurl":"http://qzapp.qlogo.cn/qzapp/111111/942FEA70050EEAFBD4DCE2C1FC775E56/30",
        '     "figureurl_1":"http://qzapp.qlogo.cn/qzapp/111111/942FEA70050EEAFBD4DCE2C1FC775E56/50",
        '     "figureurl_2":"http://qzapp.qlogo.cn/qzapp/111111/942FEA70050EEAFBD4DCE2C1FC775E56/100",
        '     "figureurl_qq_1":"http://q.qlogo.cn/qqapp/100312990/DE1931D5330620DBD07FB4A5422917B6/40",
        '     "figureurl_qq_2":"http://q.qlogo.cn/qqapp/100312990/DE1931D5330620DBD07FB4A5422917B6/100",
        '     "gender":"男",
        '     "is_yellow_vip":"1",
        '     "vip":"1",
        '     "yellow_vip_level":"7",
        '     "level":"7",
        '     "is_yellow_year_vip":"1"
        ' }
        set json = parseJSON(rs)
        ret = json.ret
        if ret > 0 then 
            html = html & "<hr/>获取用户信息失败！"
        else 
            dim nickname,figureurl,gender

            nickname = json.nickname
            figureurl = json.figureurl_2
            gender = json.gender

            html = html & "<hr/><h1>解析用户数据</h1><table border=1 class='table'><tr><th>nickname</th><td>"&nickname&"</td></tr>"
            html = html & "<tr><th>figureurl</th><td><img src="""&figureurl&""" /></td></tr>"
            html = html & "<tr><th>gender</th><td>"&gender&"</td></tr>"
            html = html & "</table>"

        end if

        html = html & "<hr/><h1>最后一步：保存用户信息，并跳转到首页</h1><br/><br/>"


        response.write(html)

    end if

end if

%>




</div>
</body>
</html>