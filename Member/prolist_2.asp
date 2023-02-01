<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!--#include file = inc2.asp -->
<!--#include file="func.asp"-->
<%
if session("u_id")="" then
response.write "<script language=javascript>window.location.href='../';</script>"
response.End
end if
pro_id=request("pro_id")
if pro_id="" or isnull(pro_id) then
   response.Write "<script language=javascript>alert('参数提交错误!');window.location.href='./';</script>"
   response.end
end if
session("order_id")=pro_id%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=uft-8" />
<title><%=hometit%>-<%=company%></title>
<meta name="keywords" content="<%=keywords%>" />
<meta http-equiv="x-ua-compatible" content="ie=7" />

<link href="style.css" rel="stylesheet" type="text/css" />
<script language="JavaScript">
<!--
//功能：去掉字符串前后空格
//返回值：去掉空格后的字符串
function fnRemoveBrank(strSource)
{
 return strSource.replace(/^\s*/,'').replace(/\s*$/,'');
}
function String.prototype.lenB()
{
return this.replace(/[^\x00-\xff]/g,"**").length;
}
function Juge(theForm)
{
  if (fnRemoveBrank(theForm.txt_Count.value) != "")
  {
	  var objv = fnRemoveBrank(theForm.txt_Count.value);
	  var pattern = /^[0-9]+$/;
	  flag = pattern.test(objv);
	  if(!flag)
	  {
		alert("请正确输入商品数量!");
		theForm.txt_Count.focus();
		return (false);
	   }
   }
  if (fnRemoveBrank(theForm.txt_Invoice.value) == "")
  {
    alert("请正确输入发票抬头!");
    theForm.txt_Invoice.focus();
    return (false);
  }
   if (fnRemoveBrank(theForm.txt_Consignee.value) == "")
  {
    alert("请正确输入收货联系人!");
    theForm.txt_Consignee.focus();
    return (false);
  }
   if (fnRemoveBrank(theForm.hukouprovince.value) == "")
  {
    alert("请选择所在省份!");
    theForm.hukouprovince.focus();
    return (false);
  }
//     if (fnRemoveBrank(theForm.school.value) == "")
//  {
//    alert("请填写毕业学校!");
//    theForm.school.focus();
//    return (false);
//  }
  if (fnRemoveBrank(theForm.txt_Telphone.value) == "" && fnRemoveBrank(theForm.txt_Mobile.value) == "")
  {
    alert("请输入联系电话或手机号码，必须填写一项!");
    theForm.txt_Telphone.focus();
    return (false);
  }
  if (fnRemoveBrank(theForm.txt_Telphone.value) != "")
  {
	  var objv = fnRemoveBrank(theForm.txt_Telphone.value);
	  var pattern = /^[0-9\s+.-]+$/;
	  flag = pattern.test(objv);
	  if(!flag)
	  {
		alert("联系电话：格式不正确!请重新输入。");
		theForm.txt_Telphone.focus();
		return (false);
	   }
  }
  if (fnRemoveBrank(theForm.post.value) != "")
  {
	  var objv = fnRemoveBrank(theForm.post.value);
	  var pattern = /^[0-9]+$/;
	  flag = pattern.test(objv);
	  if(!flag)
	  {
		alert("邮政编码：要求为数字!请重新输入。");
		theForm.post.focus();
		return (false);
	   }
   }
  for(i=0;i <document.theForm.mainRadio.length;i++){
  if (!document.theForm.mainRadio[i].checked)  
  {
  alert("选择支付方式");
  theForm.mainRadio[0].focus();
  return (false);
  }
  }
 }
function setCount(text){ 
   var count = text.value;    
   document.modicompany.result.value = count*4000; 
   document.modicompany.result1.value = count*4000;
  } 
-->
    </script>
</head>
<body oncontextmenu="return false">
<%'response.Write session("a")
'response.Write tenpay_id
'response.end
uname=M_memberID(session("u_id"))
paysql="select * from u_order where u_id="&uname&" and order_id='"&pro_id&"'"
set payrs=connn.execute(paysql)
if payrs.eof then
   response.Write "<script language=javascript>alert('您暂无商品订单');history.back();</script>"
   response.end
end if
num=payrs("pro_num")
p_id=payrs("p_id")
if p_id=0 then

a=split(payrs("pro_ids"),"|")
for i=0 to ubound(a)-1
   b=split(a(i),",")
   for j=0 to ubound(b)-1
      ab=ab&p_price(b(j),"Fk_Product_Title")&" x "&b(1)
	  if i<>ubound(a)-1 then ab=ab&"+"
      abc=abc&p_price(b(j),"Fk_Product_Title")
	  if i<>ubound(a)-1 then abc=abc&"+"
   next
next
bookname=ab
price=payrs("pro_price")
else
price=payrs("pro_price")
bookname=p_price(p_id,"Fk_Product_Title")
abc=p_price(p_id,"Fk_Product_Title")
end if
pro_fee=payrs("pro_fee_type")
pro_contact=payrs("pro_contact")
pro_mobi=payrs("pro_tel")
session("pid")=abc
session("pro_id")=payrs("order_id")
session("pro_pro_id")=payrs("order_id")
all_price=price+M_fee(2,pro_fee)
session("pro_pro_xj")=all_price*100
session("pro_pro_num")=num
session("pro_contact")=pro_contact
session("pro_mobi")=pro_mobi
session("pro_pro_price")=payrs("pro_price")
session("pro_paytype")=payrs("pro_paytype")
session("pro_fee")=M_fee(2,pro_fee)
%>
<div id="main">
<div class="m1" style="margin:0 auto; margin-top:50px;margin-bottom:50px;">
<div class="m1-title" style="background:#FFF;">订单支付：</div>
<div class="m1-body" style="padding:0px;">
<div class="midright">
        <table border="0" cellspacing="1" cellpadding="0" width="96%" bgcolor="#bfbfbf" align="center">
          <tbody>
            <tr>
              <td width="70%" height="30" bgcolor="#ffffff">&nbsp;订单号：<span id="lbl_orderNumber"><b style="color:red"><%=payrs("order_id")%></b></span></td>
              <td bgcolor="#ffffff" colspan="2">&nbsp;订购日期：<span id="lbl_orderingDate2"><%=payrs("order_time")%></span></td>
            </tr>
            <tr>
              <td height="30" bgcolor="#FFFFFF">&nbsp; <strong>商品名称</strong></td>
              <td colspan="2" align="center" bgcolor="#FFFFFF"><strong>小计</strong><strong></td>
            </tr>
            <tr>
              <td height="30" bgcolor="#FFFFFF">&nbsp; <%=bookname%></td>
              <td colspan="2" width="30%" align="center" bgcolor="#FFFFFF"><%=price%>.00</td>
            </tr>
          </tbody>
        </table>
        <table width="96%" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#bfbfbf">
          <tr>
            <td bgcolor="#ffffff" height="25" colspan="4" align="right">&nbsp;&nbsp;&nbsp;&nbsp; 
              运费：<span id="lbl_favoriablePrice"><%=M_fee(2,pro_fee)%></span>元&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
              &nbsp;&nbsp;&nbsp;</td>
          </tr>
          <tr>
            <td bgcolor="#ffffff" height="25" colspan="4" align="right">合计：<span id="lbl_totalPrice"><%=all_price%>.00</span> 元&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </td>
          </tr>
          <tr>
            <td bgcolor="#ffffff" height="25" colspan="2">&nbsp;联系人手机：<span id="lbl_mobile"><%=payrs("pro_mobi")%></span></td>
            <td bgcolor="#ffffff" height="25" colspan="2">&nbsp;电话：<span id="lbl_phone"><%=payrs("pro_tel")%></span></td>
          </tr>
          <tr>
            <td height="25" colspan="4" valign="middle" bgcolor="#ffffff">&nbsp;支付方式：<span id="lbl_patTypeInfo">
              <%select case payrs("pro_paytype")
															case 1
															   img="<img src=""images/payType/kq.gif"" />"
															   paylink="<a href=""javascript:void(0);"" onclick=""window.location.href='pay.asp';return false;"" id=""hl_Submit"" target=""_self""><IMG style=""CURSOR: hand"" border=""0"" src=""images/fkkk.gif"" width=""288"" height=""39"">"
															case 2
															   img="<img src=""images/payType/zfb.gif"" />"
															   paylink="<a href=""javascript:void(0);"" onclick=""window.location.href='pay.asp';return false;"" id=""hl_Submit"" target=""_self""><IMG style=""CURSOR: hand"" border=""0"" src=""images/fkkk.gif"" width=""288"" height=""39"">"
															case 3
															   img="<img src=""images/payType/tenpay.gif"" />"
															   paylink="<a href=""javascript:void(0);"" onclick=""window.location.href='pay.asp';return false;"" id=""hl_Submit"" target=""_self""><IMG style=""CURSOR: hand"" border=""0"" src=""images/fkkk.gif"" width=""288"" height=""39"">"
															case 99999
															   img="<img src=""images/payType/yhfk.gif"" />"
															   paylink="您选择的是银行汇款,请汇款后及时通知我们! &nbsp; <a href=""javascript:void(0);"" onclick=""window.location.href='pay_type.asp';return false;"" target=_self' style='color:#0066FF;' title='点击查看银行汇款账号'>查看银行汇款账号</a>"
															case 100000
															   img="货到付款"
															   paylink="您选择的是货到付款,请及时查收付款!"
															end select
															%>
            <%=img%></span></td>
          </tr>
        </table>
</TD>
								<tr>
									<td><br>
										<table border="0" cellSpacing="0" cellPadding="0" width="96%" align="center">
											<tr>
											  <td width="138"></td>
												<td width="449"><%=paylink%></td>
										  </tr>
									  </table>
									</td>
								</tr>
								</TBODY></TABLE> 
 	                          </div>
		<div style="clear:both; height:5px;"></div>
</div>
<div class="m1-foot"></div>
</div>
</div>
<%payrs.close
set payrs=nothing
connn.close
set connn=nothing%>
</body>
</html>