<%

'//////方法区////////////


'//////方法区结束////////



dim wxurl
dim appid
dim callback
dim state
dim url

wxurl = "https://open.weixin.qq.com/connect/qrconnect?"
' APPID
appid = "xxx"
' 微信授权回调地址
callback = "http://domain/login_wx_callback.asp" 
' 标识，可以随机
state = "wechat_login"
' 二维码登录地址
url =  wxurl & "appid=" & appid & "&redirect_uri=" & Server.URLEncode(callback) & "&response_type=code&scope=snsapi_login&state=" & state & "#wechat_redirect"

' url = "https://open.weixin.qq.com/connect/qrconnect?appid=wx907282aee1e8b998&redirect_uri=http%3A%2F%2Fapi.3d66.com%2Ftplogin%2Flogin_wx_callback.asp&response_type=code&scope=snsapi_login&state=3d6be0a4035d839573b04816624a415e#wechat_redirect"

'跳转到二维码地址
' response.write(url)
response.redirect(url)

%>