<!--#include file = "../admin/admin_conn.asp" -->
<!-- #include file="../config.asp" -->
<%
'*
'* @Description: 快钱人民币支付网关接口范例
'* @Copyright (c) 上海快钱信息服务有限公司
'* @version 2.0
'*

'*
'在本文件中，商家应从数据库中，查询到订单的状态信息以及订单的处理结果。给出支付人响应的提示。

'本范例采用最简单的模式，直接从receive页面获取支付状态提示给用户。
'*

orderId=trim(request("orderId"))
orderAmount=trim(request("orderAmount"))
msg=trim(request("msg"))
select case msg
   case "ok"
      msg="支付成功!"
	  conn.execute("update u_order set pro_paystatu=1 where order_id='"&orderId&"'")
	  conn.close
	  set conn=nothing
   case "false"
      msg="支付失败!"
   case "error"
      msg="支付错误!"
   case else
end select
%>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en" >
<html>
	<head>
		<title>快钱支付结果</title>
		<meta http-equiv="content-type" content="text/html; charset=gb2312" >
        <link href="../css.css" rel="stylesheet" type="text/css" />
	</head>
	
<BODY>
	
	<div align="center" style="margin:0 auto; margin-top:200px;">
		<table width="259" border="0" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC" >
	  <tr bgcolor="#FFFFFF">
				<td width="80">支付方式:</td>
				<td >快钱[99bill]</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td >订单编号:</td>
				<td ><%=orderId %></td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td>订单金额:</td>
				<td><%=(orderAmount)/100%> 元</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td>支付结果:</td>
				<td><%=msg %></td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td></td>
				<td><%if msg<>"ok" then%>
				<input type="button" name="button" id="button" value="重新支付" onClick="javascript:window.location.href='../prolist_2.asp?pro_id=<%=orderId%>';">
				<%else%>
			    <input type="button" name="button" id="button" value="返回用户中心" onClick="javascript:window.location.href='../member/member_center.asp';">
				<%end if%></td>
			</tr>
	  </table>
</div>

</BODY>
</HTML>