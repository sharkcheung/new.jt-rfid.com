<!--#include file = "../admin_conn.asp" -->
<!-- #include file="../config1.asp" -->
<%
	'版本：2.0
	'日期：2009-07-30
	'
	'说明：
	'以下代码只是方便商户测试，提供的样例代码，商户可以根据自己网站的需要，按照技术文档编写,并非一定要使用该代码。
	'该代码仅供学习和研究支付宝接口使用，只是提供一个参考。

	
	seller_email	= alipay_uid	 '请填写签约支付宝账号，
	partner			= alipay_id	 '填写签约支付宝账号对应的partnerID，
	key			    = alipay_key	 '填写签约账号对应的安全校验码

    notify_url		= ""&yuming&"/alipay/Alipay_Notify.asp"	        '交易过程中服务器通知的页面 要用 http://格式的完整路径，例如http://www.alipay.com/alipay/Alipay_Notify.asp  注意文件位置请填写正确。
	return_url		= ""&yuming&"/alipay/return_Alipay_Notify.asp"	'付完款后跳转的页面 要用 http://格式的完整路径, 例如http://www.alipay.com/alipay/return_Alipay_Notify.asp  注意文件位置请填写正确。
	'如果使用了Alipay_Notify.asp或者return_Alipay_Notify.asp，请在这两个文件中添加相应的合作身份者ID和安全校验码
	logistics_fee	   = "0.00"			'物流配送费用
	logistics_payment  = "SELLER_PAY"	'物流配送费用付款方式：SELLER_PAY(卖家支付)、BUYER_PAY(买家支付)、BUYER_PAY_AFTER_RECEIVE(货到付款)
	logistics_type	   = "EXPRESS"		'物流配送方式：POST(平邮)、EMS(EMS)、EXPRESS(其他快递)

	 	 
'登陆 www.alipay.com 后, 点商家服务,可以看到支付宝安全校验码和合作id,导航栏的下面 
%>