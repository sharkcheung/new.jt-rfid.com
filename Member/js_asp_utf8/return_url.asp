<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
	'功能：付完款后跳转的页面（返回页）
	'版本：3.0
	'日期：2010-05-28
	'说明：
	'以下代码只是为了方便商户测试而提供的样例代码，商户可以根据自己网站的需要，按照技术文档编写,并非一定要使用该代码。
	'该代码仅供学习和研究支付宝接口使用，只是提供一个参考。
	
''''''''页面功能说明''''''''''''''''
'该页面可在本机电脑测试
'该页面称作“返回页”，是由支付宝服务器同步调用，可当作是支付完成后的提示信息页，如“您的某某某订单，多少金额已支付成功”。
'可放入HTML等美化页面的代码和订单交易完成后的数据库更新程序代码
'该页面可以使用ASP开发工具调试，也可以使用写文本函数log_result进行调试，该函数已被默认关闭，见alipay_notify.asp中的函数return_verify
'TRADE_FINISHED(表示交易已经成功结束，为普通即时到帐的交易状态成功标识);
'TRADE_SUCCESS(表示交易已经成功结束，为高级即时到帐的交易状态成功标识);
''''''''''''''''''''''''''''''''''''
%>

<!--#include file="alipay_config.asp"-->
<!--#include file="class/alipay_notify.asp"-->

<%
'计算得出通知验证结果
verify_result = return_verify()

if verify_result then	'验证成功
    '获取支付宝的通知返回参数
    order_no		= request.QueryString("out_trade_no")	'获取订单号
    total_fee		= request.QueryString("total_fee")		'获取总金额
    sOld_trade_status = "0"									'获取商户数据库中查询得到该笔交易当前的交易状态

    '假设：
	'sOld_trade_status="0"	表示订单未处理；
	'sOld_trade_status="1"	表示交易成功（TRADE_FINISHED/TRADE_SUCCESS）
	
	if request.QueryString("trade_status") = "TRADE_FINISHED" or request.QueryString("trade_status") = "TRADE_SUCCESS" then
		'为了保证不被重复调用，或重复执行数据库更新程序，请判断该笔交易状态是否是订单未处理状态
		if sOld_trade_status < 1 then
			'根据订单号更新订单，把订单处理成交易成功
		end if
	else
		response.Write "trade_status="&request.QueryString("trade_status")
	end if
else '验证失败
    '如要调试，请看alipay_notify.asp页面的return_verify函数，比对sign和mysign的值是否相等，或者检查responseTxt有没有返回true
    response.Write "fail"
end if
%>
<html>
<head>
<title>支付宝即时支付</title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<style type="text/css">
            .font_content{
                font-family:"宋体";
                font-size:14px;
                color:#FF6600;
            }
            .font_title{
                font-family:"宋体";
                font-size:16px;
                color:#FF0000;
                font-weight:bold;
            }
            table{
                border: 1px solid #CCCCCC;
            }
        </style>
</head>
<body>
<table align="center" width="350" cellpadding="5" cellspacing="0">
  <tr>
    <td align="center" class="font_title" colspan="2">通知返回</td>
  </tr>
  <tr>
    <td class="font_content" align="right">支付宝交易号：</td>
    <td class="font_content" align="left"><%=request.QueryString("trade_no")%></td>
  </tr>
  <tr>
    <td class="font_content" align="right">订单号：</td>
    <td class="font_content" align="left"><%=request.QueryString("out_trade_no")%></td>
  </tr>
  <tr>
    <td class="font_content" align="right">付款总金额：</td>
    <td class="font_content" align="left"><%=request.QueryString("total_fee")%></td>
  </tr>
  <tr>
    <td class="font_content" align="right">商品标题：</td>
    <td class="font_content" align="left"><%=request.QueryString("subject")%></td>
  </tr>
  <tr>
    <td class="font_content" align="right">商品描述：</td>
    <td class="font_content" align="left"><%=request.QueryString("body")%></td>
  </tr>
  <tr>
    <td class="font_content" align="right">买家账号：</td>
    <td class="font_content" align="left"><%=request.QueryString("buyer_email")%></td>
  </tr>
  <tr>
    <td class="font_content" align="right">交易状态：</td>
    <td class="font_content" align="left"><%=request.QueryString("trade_status")%></td>
  </tr>
</table>
</body>
</html>
