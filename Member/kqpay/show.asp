<!--#include file = "../admin/admin_conn.asp" -->
<!-- #include file="../config.asp" -->
<%
'*
'* @Description: ��Ǯ�����֧�����ؽӿڷ���
'* @Copyright (c) �Ϻ���Ǯ��Ϣ�������޹�˾
'* @version 2.0
'*

'*
'�ڱ��ļ��У��̼�Ӧ�����ݿ��У���ѯ��������״̬��Ϣ�Լ������Ĵ�����������֧������Ӧ����ʾ��

'������������򵥵�ģʽ��ֱ�Ӵ�receiveҳ���ȡ֧��״̬��ʾ���û���
'*

orderId=trim(request("orderId"))
orderAmount=trim(request("orderAmount"))
msg=trim(request("msg"))
select case msg
   case "ok"
      msg="֧���ɹ�!"
	  conn.execute("update u_order set pro_paystatu=1 where order_id='"&orderId&"'")
	  conn.close
	  set conn=nothing
   case "false"
      msg="֧��ʧ��!"
   case "error"
      msg="֧������!"
   case else
end select
%>
<!doctype html public "-//w3c//dtd html 4.0 transitional//en" >
<html>
	<head>
		<title>��Ǯ֧�����</title>
		<meta http-equiv="content-type" content="text/html; charset=gb2312" >
        <link href="../css.css" rel="stylesheet" type="text/css" />
	</head>
	
<BODY>
	
	<div align="center" style="margin:0 auto; margin-top:200px;">
		<table width="259" border="0" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC" >
	  <tr bgcolor="#FFFFFF">
				<td width="80">֧����ʽ:</td>
				<td >��Ǯ[99bill]</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td >�������:</td>
				<td ><%=orderId %></td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td>�������:</td>
				<td><%=(orderAmount)/100%> Ԫ</td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td>֧�����:</td>
				<td><%=msg %></td>
			</tr>
			<tr bgcolor="#FFFFFF">
				<td></td>
				<td><%if msg<>"ok" then%>
				<input type="button" name="button" id="button" value="����֧��" onClick="javascript:window.location.href='../prolist_2.asp?pro_id=<%=orderId%>';">
				<%else%>
			    <input type="button" name="button" id="button" value="�����û�����" onClick="javascript:window.location.href='../member/member_center.asp';">
				<%end if%></td>
			</tr>
	  </table>
</div>

</BODY>
</HTML>