<%
'判断是否从外部提交
Function UrlPost()
Dim Url_1,Url_2,strMetHod
UrlPost = False
Url_1 = Cstr(Request.ServerVariables("HTTP_REFERER"))
Url_2 = Cstr(Request.ServerVariables("SERVER_NAME"))
strMetHod = UCase(Request.ServerVariables("REQUEST_METHOD"))
If Mid(Url_1,8,Len(Url_2))<>Url_2 Then
   If strMetHod = "GET" Then
    UrlPost = True
   Else
    UrlPost = False
   End If
Else
   UrlPost = True
End If
End Function
if urlpost=false then
   response.end
end if
	'功能：设置帐户有关信息及返回路径（基础配置页面）
	'版本：3.0
	'日期：2010-05-27
	'说明：
	'以下代码只是为了方便商户测试而提供的样例代码，商户可以根据自己网站的需要，按照技术文档编写,并非一定要使用该代码。
	'该代码仅供学习和研究支付宝接口使用，只是提供一个参考。

'提示：如何获取安全校验码和合作身份者ID
'1.访问支付宝首页(www.alipay.com)，然后用您的签约支付宝账号登陆.
'2.点击导航栏中的“商家服务”，即可查看

'安全校验码查看时，输入支付密码后，页面呈灰色的现象，怎么办？
'解决方法：
'1、检查浏览器配置，不让浏览器做弹框屏蔽设置
'2、更换浏览器或电脑，重新登录查询。
partner         = session("cjvljd_civjcid")					'合作身份者ID
key   			= session("cjvljd_civjkey")	'安全检验码
seller_email    = "397056446@qq.com"					'签约支付宝账号或卖家支付宝帐户

input_charset 	= "utf-8"								'字符编码格式 目前支持 gb2312 或 utf-8

notify_url      = "http://shallwin.qebang.net/alipay/js_asp_utf8/notify_url.asp"	'交易过程中服务器通知的页面 要用 http://格式的完整路径，不允许加?id=123这类自定义参数
return_url      = "http://shallwin.qebang.net/alipay/js_asp_utf8/return_url.asp"	'付完款后跳转的页面 要用 http://格式的完整路径，不允许加?id=123这类自定义参数
show_url        = "http://shallwin.qebang.net/"				'网站商品的展示地址，不允许加?id=123这类自定义参数

sign_type       = "MD5"									'加密方式 不需修改
antiphishing    = "0"									'防钓鱼功能开关，'0'表示该功能关闭，'1'表示该功能开启。默认为关闭
'一旦开启，就无法关闭，根据商家自身网站情况请慎重选择是否开启。
'申请开通方法：联系我们的客户经理或拨打商户服务电话0571-88158090，帮忙申请开
'若要使用防钓鱼功能，建议使用POST方式请求数据
mainname		= "深圳市企帮网络技术有限公司"							'收款方名称，如：公司名称、网站名称、收款人姓名等
%>