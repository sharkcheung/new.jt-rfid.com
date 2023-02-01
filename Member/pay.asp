<!--#include file = inc2.asp -->
<!--#include file="func.asp"-->
<%
'if session("u_id")="" then
'response.write "<script language=javascript>alert('登录超时,请重新登录!');window.location.href='../';<'/script>"
'response.End
'end if
   PayKey=session("pro_paytype")
   orderAmount=session("pro_pro_xj")
   orderid=session("pro_id")
   productid=session("pid")
   payerName=session("u_id")
   pro_pro_num=session("pro_pro_num")
   pro_price=session("pro_pro_price")
   pro_mobi=session("pro_mobi")
   pro_contact=session("pro_contact")
   pro_fee=session("pro_fee")
   if PayKey="" then
         response.write "请选择支付平台<br><br><a href=javascript:history.back()>返回</a>"
         response.end
         
   end if
   if PayKey<>"" then         '修改
   
         set rs=server.CreateObject("ADODB.RecordSet") 
         sql="select * from Iheeo_Pay where PayKey="&PayKey
         set rs=connn.execute(sql)
         if not rs.eof then
                   Name=trim(rs("PayName"))
                   typeid=trim(rs("PayKey"))
                   cid=trim(rs("PayShopID"))
                   mykey=trim(rs("PayShopKey"))
         else
                   response.write "参数有误"
                   response.end
         end if
         rs.close
         set rs=nothing
   end if
   
   if PayKey=1 then
      nenturl="kqpay/send.asp"
   end if
   if PayKey=2 then
      nenturl="alipay/js_asp_utf8/index.asp"
   end if   
   if PayKey=3 then
      nenturl="tenpay/"
   end if
   if PayKey=4 then
      nenturl="Iheeo_Pay/yeepay/send.asp"
   end if 
   if PayKey=5 then
      nenturl="Iheeo_Pay/alipay/send.asp"
   end if       
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
<title>在线支付</title>
</head>

<body oncontextmenu="return false">
<form name="form_pcmmks_sd" action="<%=nenturl%>" method="post">
<input type="hidden" name="orderid" value="<%=orderid%>">
<input type="hidden" name="orderAmount" value="<%=orderAmount%>">
<input type="hidden" name="productid" value="<%=productid%>">
<input type="hidden" name="payerName" value="<%=payerName%>">
<input type="hidden" name="pro_price" value="<%=pro_price%>">
<input type="hidden" name="pro_pro_num" value="<%=pro_pro_num%>">
<input type="hidden" name="pro_mobi" value="<%=pro_mobi%>">
<input type="hidden" name="pro_contact" value="<%=pro_contact%>">
<input type="hidden" name="cid" value="<%=cid%>">
<input type="hidden" name="mykey" value="<%=mykey%>">
<input type="hidden" name="pro_fee" value="<%=pro_fee%>">

</form>
<script>
document.form_pcmmks_sd.submit();
</script>

</body>

</html>