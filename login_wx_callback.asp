<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="GBK">
    <meta http-equiv="X-UA-Compatible" content="IE=edge">
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>微信回调解析</title>
    <!-- 新 Bootstrap 核心 CSS 文件 -->
    <link rel="stylesheet" href="http://cdn.bootcss.com/bootstrap/3.3.0/css/bootstrap.min.css">
</head>
<body>
<div class="container">


<!-- 以下内容为解析用户信息 -->
<!-- 以下内容为解析用户信息 -->
<!-- 以下内容为解析用户信息 -->
<!-- 以下内容为解析用户信息 -->
<!-- 以下内容为解析用户信息 -->




 <script language="jscript" runat="server">  
    Array.prototype.get = function(x) { return this[x]; };  
    function parseJSON(strJSON) { return eval("(" + strJSON + ")"); }  
    function contain(src,str) {return src.indexOf(str) >= 0; }
</script>  

<%

' 微信授权回调，要获得用户信息需要以下步骤
' 1、通过code获取access_token，
' 2、通过access_token调用接口获取用户信息

' 说明：refresh_token拥有较长的有效期（30天），当refresh_token失效的后，需要用户重新授权，所以，请开发者在refresh_token即将过期时（如第29天时），进行定时的自动刷新并保存好它。


' 方法区

Public Function GetData(url)

    Dim request
    Dim rs
    set request = Server.CreateObject("Msxml2.ServerXMLHTTP.6.0")
    request.Open "GET", url, False
    request.send
    
    rs = request.ResponseText

    GetData = rs

End Function 

' 方法区结束


' 网站信息
dim appid
dim appsec

appid = "xxx"
appsec = "xxx"

'接收到微信回传的参数，用于获取用户数据
dim code
dim state
code = request.querystring("CODE")
state = request.querystring("STATE")

' 微信服务地址
dim tokenurl
dim userinfourl

tokenurl = "https://api.weixin.qq.com/sns/oauth2/access_token?appid="& appid &"&secret="& appsec 
userinfourl = "https://api.weixin.qq.com/sns/userinfo?"

' 输出内容，用于调试！！！后续将用户信息存入相关数据库
' 输出内容，用于调试！！！后续将用户信息存入相关数据库
' 输出内容，用于调试！！！后续将用户信息存入相关数据库
dim html

html = html & "<h1>微信登录回调，解析用户</h1>"

if isEmpty(code) Then
    html = "code is null."
else 
    dim rs 
    

    tokenurl = tokenurl & "&code="& code &"&grant_type=authorization_code"
    rs  = GetData(tokenurl)

    html = html & "<h2>第一步：根据CODE获取TOKEN</h2>token url : " & tokenurl 
    html = html & "<hr/>token url response : " & rs 


    ' 如果返回结果中有errcode代表错误
    if contain(rs,"errcode") Then
        ' {"errcode":40029,"errmsg":"invalid code"}
        html = html & "<hr/> 获取token 失败. 返回登录 <a href=""http://api.3d66.com/tplogin/login_wx.asp"">http://api.3d66.com/tplogin/login_wx.asp</a>"
    else

        ' { 
        ' "access_token":"ACCESS_TOKEN", 
        ' "expires_in":7200, 
        ' "refresh_token":"REFRESH_TOKEN",
        ' "openid":"OPENID", 
        ' "scope":"SCOPE" 
        ' }

        ' 获取到token了，可以将token缓存或存入数据库中
        ' 使用token和openid去获取用户信息咯
        
        set json = parseJSON(rs)
        dim json
        dim access_token
        dim refresh_token
        dim openid
        
        access_token = json.access_token
        refresh_token = json.refresh_token
        openid = json.openid

        html = html & "<hr/><h3>解析结果：</h3><table class='table'><tr><th>access_token</th><td>" & access_token & "</td> </tr>"
        html = html & "<tr><th>refresh_token </th><td>" & refresh_token & "</td> </tr>"
        html = html & "<tr><th>openid </th><td> " & openid & "</td> </tr></table>"

        userinfourl = userinfourl & "access_token=" & access_token & "&openid=" & openid
        html = html & "<hr/><h2>第二步：根据TOKEN和OPENID获取用户信息</h2>userinfo url : " & userinfourl 

        rs = GetData(userinfourl)
        html = html & "<hr/>userinfo url response : " & rs 

        ' 如果包含errcode，获取用户信息失败
        if contain(rs,"errcode") then 
            html = html & "<hr/> 获取用户信息失败. 返回登录 <a href=""http://api.3d66.com/tplogin/login_wx.asp"">http://api.3d66.com/tplogin/login_wx.asp</a>"
        else 

            ' {
            ' "openid": "oIhOdv4smUBmovJkp6mP8gzoqS6w",
            ' "nickname": "熙",
            ' "sex": 1,
            ' "language": "zh_CN",
            ' "city": "Chaoyang",
            ' "province": "Beijing",
            ' "country": "CN",
            ' "headimgurl": "http://wx.qlogo.cn/mmopen/iawHyYlDnQpk5sf5JOI6CyPDibs2DPnRpicjODicpcch2H6P5SVlILjrAC6FLFXPUws9llQ19DoqVKDsUUjNoSDLklaUFibU44ZCG/0",
            ' "privilege": [],
            ' "unionid": "oE0_9wQXFkMpM-4955gEhgD7-RVw"
            ' }

            set json = parseJSON(rs)

            dim nickname
            dim sex
            dim city
            dim province
            dim country
            dim headimgurl
            
            nickname = json.nickname
            sex = json.sex
            city = json.city
            province = json.province
            country = json.country
            headimgurl = json.headimgurl

            html = html & "<hr/><h3>解析结果：</h3><table class='table'><tr><th>nickname</th><td>" & nickname & "</td> </tr>"
            html = html & "<tr><th>headimgurl </th><td><img width='250' src='" & headimgurl & "'/></td> </tr>"
            html = html & "<tr><th>sex </th><td>" & sex & "</td> </tr>"
            html = html & "<tr><th>city </th><td>" & city & "</td> </tr>"
            html = html & "<tr><th>province </th><td>" & province & "</td> </tr>"
            html = html & "<tr><th>country </th><td>" & country & "</td> </tr></table>"

            html = html & "<hr/><h2>第三步：保存用户信息，并跳转到首页</h2> <hr/>"  


        end if
    end if
end if

response.write(html)

%>




<br/>
<br/>
<br/>
</div>
</body>
</html>
 