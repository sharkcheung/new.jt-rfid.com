<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%> 
<!--#include file="./classes/MediPayResponseHandler.asp"-->
<%
'---------------------------------------------------------
'财付通中介担保支付应答（处理回调）示例，商户按照此文档进行开发即可
'---------------------------------------------------------

'平台商密钥
Dim key
key = "123456"

'创建支付应答对象
Dim resHandler
Set resHandler = new MediPayResponseHandler
resHandler.setKey(key)

'判断签名
If resHandler.isTenpaySign() = True Then
	
	Dim cft_tid
	Dim total_fee
	Dim retcode
	Dim status

	'财付通交易单号
	cft_tid = resHandler.getParameter("cft_tid")
	
	'金额金额,以分为单位
	total_fee = resHandler.getParameter("total_fee")

	'返回码
	retcode = resHandler.getParameter("retcode")

	'状态
	status = resHandler.getParameter("status")
	
	'------------------------------
	'处理业务开始
	'------------------------------ 

	'注意交易单不要重复处理
	'注意判断返回金额
	
	'返回码判断
	If "0" = retcode Then
		Select Case status
			Case "1":	'交易创建

			Case "2":	'收获地址填写完毕

			Case "3":	'买家付款成功，注意判断订单是否重复的逻辑

			Case "4":	'卖家发货成功

			Case "5":	'买家收货确认，交易成功

			Case "6":	'交易关闭，未完成超时关闭

			Case "7":	'修改交易价格成功

			Case "8":	'买家发起退款

			Case "9":	'退款成功

			Case "10":	'退款关闭

			Case else:	'error
				'nothing to do
		End Select
	Else
		'错误的返回码
		
	End If
	
	'------------------------------
	'处理业务完毕
	'------------------------------

	resHandler.doShow()
	

Else

	'签名失败
	Response.Write("签名签证失败")

	'Dim debugInfo
	'debugInfo = resHandler.getDebugInfo()
	'Response.Write("<br/>debugInfo:" & debugInfo & "<br/>")

End If


%>