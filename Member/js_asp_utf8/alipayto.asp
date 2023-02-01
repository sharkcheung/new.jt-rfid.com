<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8"/>
<%
	'功能：设置商品有关信息（确认订单支付宝在线购买入口页）
	'详细：该页面是接口入口页面，生成支付时的URL
	'版本：3.0
	'日期：2010-06-22
	'说明：
	'以下代码只是为了方便商户测试而提供的样例代码，商户可以根据自己网站的需要，按照技术文档编写,并非一定要使用该代码。
	'该代码仅供学习和研究支付宝接口使用，只是提供一个参考。
	
'''''''''''''''''注意'''''''''''''''''''''''''
'该页面测试时出现“调试错误”请参考：http://club.alipay.com/read-htm-tid-8681712.html
'要传递的参数要么不允许为空，要么就不要出现在数组与隐藏控件或URL链接里。
''''''''''''''''''''''''''''''''''''''''''''''
%>

<!--#include file="alipay_config.asp"-->
<!--#include file="class/alipay_service.asp"-->

<%
'''以下参数是需要通过下单时的订单数据传入进来获得'''
'必填参数
sTime=now()
out_trade_no=request.Form("orderid")'请与贵网站订单系统中的唯一订单号匹配
subject      = request.Form("aliorder")		'订单名称，显示在支付宝收银台里的“商品名称”里，显示在支付宝的交易管理的“商品名称”的列表里。
body         = request.Form("alibody")		'订单描述、订单详细、订单备注，显示在支付宝收银台里的“商品描述”里
total_fee    = request.Form("alimoney")		'订单总金额，显示在支付宝收银台里的“应付总额”里

'扩展功能参数——网银提前
pay_mode	 = request.Form("pay_bank")
if pay_mode = "directPay" then
	paymethod    = "directPay"	'默认支付方式，四个值可选：bankPay(网银); cartoon(卡通); directPay(余额); CASH(网点支付)
	defaultbank	 = ""
else
	paymethod    = "bankPay"	'默认支付方式，四个值可选：bankPay(网银); cartoon(卡通); directPay(余额); CASH(网点支付)
	defaultbank  = pay_mode		'默认网银代号，代号列表见http://club.alipay.com/read.php?tid=8681379
end if

'扩展功能参数——防钓鱼
encrypt_key  = ""				'防钓鱼时间戳，初始值
exter_invoke_ip = ""			'客户端的IP地址，初始值
if(antiphishing = "1") then
	encrypt_key = query_timestamp(partner)
	exter_invoke_ip = ""		'获取客户端的IP地址，建议：编写获取客户端IP地址的程序
end if

'扩展功能参数——其他
extra_common_param = ""			'自定义参数，可存放任何内容（除=、&等特殊字符外），不会显示在页面上
buyer_email		   = ""			'默认买家支付宝账号

'扩展功能参数——分润(若要使用，请按照注释要求的格式赋值)
royalty_type		= ""		'提成类型，该值为固定值：10，不需要修改
royalty_parameters	= ""
'提成信息集，与需要结合商户网站自身情况动态获取每笔交易的各分润收款账号、各分润金额、各分润说明。最多只能设置10条
'提成信息集格式为：收款方Email_1^金额1^备注1|收款方Email_2^金额2^备注2
'如：
'royalty_type = "10"
'royalty_parameters	= "111@126.com^0.01^分润备注一|222@126.com^0.01^分润备注二"

'扩展功能参数——自定义超时(若要使用，请按照注释要求的格式赋值)
'该功能默认不开通，需联系客户经理咨询
it_b_pay			= ""	  '超时时间，不填默认是15天。八个值可选：1h(1小时),2h(2小时),3h(3小时),1d(1天),3d(3天),7d(7天),15d(15天),1c(当天)

''''''''''''''''''''''''''''''''''''''''''''''''''''
'构造要请求的参数数组，无需改动
para = Array("service=create_direct_pay_by_user","payment_type=1","partner="&partner,"seller_email="&seller_email,"return_url="&return_url,"notify_url="&notify_url,"_input_charset="&input_charset,"show_url="&show_url,"out_trade_no="&out_trade_no,"subject="&subject,"body="&body,"total_fee="&total_fee,"paymethod="&paymethod,"defaultbank="&defaultbank,"anti_phishing_key="&encrypt_key,"exter_invoke_ip="&exter_invoke_ip,"extra_common_param="&extra_common_param,"buyer_email="&buyer_email,"royalty_type="&royalty_type,"royalty_parameters="&royalty_parameters,"it_b_pay="&it_b_pay)

'构造请求函数
alipay_service(para)


'若改成GET方式传递
url = create_url()
sHtmlText = "<a href="& url &"><img border='0' src='images/alipay.gif' /></a>"
response.Redirect url

'POST方式传递，得到加密结果字符串，请取消下面一行的注释
'sHtmlText = build_postform()
%>