<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
	'功能：支付宝主动通知调用的页面（通知页）
	'版本：3.0
	'日期：2010-05-28
	'说明：
	'以下代码只是为了方便商户测试而提供的样例代码，商户可以根据自己网站的需要，按照技术文档编写,并非一定要使用该代码。
	'该代码仅供学习和研究支付宝接口使用，只是提供一个参考。
	
''''''''''''页面功能说明'''''''''''''''''''
'创建该页面文件时，请留心该页面文件中无任何HTML代码及空格。
'该页面不能在本机电脑测试，请到服务器上做测试。请确保外部可以访问该页面。
'该页面调试工具请使用写文本函数log_result，该函数已被默认开启，见alipay_notify.asp中的函数notify_verify
'TRADE_FINISHED(表示交易已经成功结束，通用即时到帐反馈的交易状态成功标志);
'TRADE_SUCCESS(表示交易已经成功结束，高级即时到帐反馈的交易状态成功标志);
'该通知页面主要功能是：对于返回页面（return_url.asp）做补单处理。如果没有收到该页面返回的 success 信息，支付宝会在24小时内按一定的时间策略重发通知
'''''''''''''''''''''''''''''''''''''''''''
%>

<!--#include file="alipay_config.asp"-->
<!--#include file="class/alipay_notify.asp"-->

<%
'计算得出通知验证结果
verify_result = notify_verify()

if verify_result then	'验证成功
    '获取支付宝的通知返回参数
    order_no		= request.Form("out_trade_no")	'获取订单号
    total_fee		= request.Form("total_fee")		'获取总金额
    sOld_trade_status = 1							'获取商户数据库中查询得到该笔交易当前的交易状态

    '假设：
	'sOld_trade_status="0"	表示订单未处理；
	'sOld_trade_status="1"	表示交易成功（TRADE_FINISHED/TRADE_SUCCESS）
	
	if request.Form("trade_status") = "TRADE_FINISHED" or request.Form("trade_status") = "TRADE_SUCCESS" then
		'放入订单交易完成后的数据库更新程序代码，请务必保证response.Write出来的信息只有success
		'为了保证不被重复调用，或重复执行数据库更新程序，请判断该笔交易状态是否是订单未处理状态
		if sOld_trade_status < 1 then
			'根据订单号更新订单，把订单处理成交易成功
		end if
		
		response.Write "success"	'请不要修改或删除
		
		'调试用，写文本函数记录程序运行情况是否正常
        'log_result("这里写入想要调试的代码变量值，或其他运行的结果记录")
	else
		response.Write "success"	'其他状态判断。普通即时到帐中，其他状态不用判断，直接打印success。
		'调试用，写文本函数记录程序运行情况是否正常
		'log_result ("这里写入想要调试的代码变量值，或其他运行的结果记录")
	end if
else '验证失败
    response.Write "fail"
	'调试用，写文本函数记录程序运行情况是否正常
	'log_result ("这里写入想要调试的代码变量值，或其他运行的结果记录")
end if
%>