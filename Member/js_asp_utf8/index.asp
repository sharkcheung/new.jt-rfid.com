<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
	'功能：快速付款入口模板页
	'详细：该页面是针对不涉及到购物车流程、充值流程等业务流程，只需要实现买家能够快速付款给卖家的付款功能。
	'版本：3.1
	'日期：2010-07-26
	'说明：
	'以下代码只是为了方便商户测试而提供的样例代码，商户可以根据自己网站的需要，按照技术文档编写,并非一定要使用该代码。
	'该代码仅供学习和研究支付宝接口使用，只是提供一个参考。
	cid=trim(request("cid"))
	mykey=trim(request("mykey"))
	product_name=trim(request("productid"))
	trueName=trim(request("pro_contact"))
	contact_phone=trim(request("pro_mobi"))
	product_price=request("pro_price")
	payName=trim(request("payerName"))
	orderid=trim(request("orderid"))
	pro_fee=request("pro_fee")
	session("cjvljd_civjcid")=cid
	session("cjvljd_civjkey")=mykey
	product_price=product_price-0+pro_fee
	details="会员账号: "&payName&""&chr(10)&"收货人姓名: "&trueName&""&chr(10)&"联系电话: "&contact_phone&""&chr(10)&"订单号: "&orderid&""
%>
<!--#include file="alipay_Config.asp"-->
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML XMLNS:CC><HEAD><TITLE>支付宝 - 网上支付 安全快速！</TITLE>
<META http-equiv=Content-Type content="text/html; charset=utf-8">
<META content=网上购物/网上支付/安全支付/安全购物/购物，安全/支付,安全/支付宝/安全,支付/安全，购物/支付, 
name=description 在线 付款,收款 网上,贸易 网上贸易.>
<META content=网上购物/网上支付/安全支付/安全购物/购物，安全/支付,安全/支付宝/安全,支付/安全，购物/支付, name=keywords 
在线 付款,收款 网上,贸易 网上贸易.><LINK href="images/layout.css" 
type=text/css rel=stylesheet>

<SCRIPT language=JavaScript>
<!-- 
  //校验输入框  -->
function CheckForm()
{
	if (document.alipayment.aliorder.value.length == 0) {
		alert("请输入商品名称.");
		document.alipayment.aliorder.focus();
		return false;
	}
	if (document.alipayment.alimoney.value.length == 0) {
		alert("请输入付款金额.");
		document.alipayment.alimoney.focus();
		return false;
	}
	if (document.alipayment.buyer_mail.value.length == 0) {
		alert("请输入付款方信息.");
		document.alipayment.alimoney.focus();
		return false;
	}

}  

<!-- 
  //控制文字显示 -->
function glowit(which){
if (document.all.glowtext[which].filters[0].strength==2)
document.all.glowtext[which].filters[0].strength=1
else
document.all.glowtext[which].filters[0].strength=2
}
function glowit2(which){
if (document.all.glowtext.filters[0].strength==2)
document.all.glowtext.filters[0].strength=1
else
document.all.glowtext.filters[0].strength=2
}
function startglowing(){
if (document.all.glowtext&&glowtext.length){
for (i=0;i<glowtext.length;i++)
eval('setInterval("glowit('+i+')",150)')
}
else if (glowtext)
setInterval("glowit2(0)",150)
}
if (document.all)
window.onload=startglowing


</SCRIPT>
<title>支付宝即时到帐付款快速通道</title>
</HEAD>
<style>
<!--
#glowtext{
filter:glow(color=red,strength=2);
width:100%;
}
.STYLE1,.style2{line-height:120%; word-spacing:2px; letter-spacing:1px;}
.style2 span{color:#A80000;font-weight:bold;}
-->
</style>
<BODY text=#000000 bgColor=#ffffff leftMargin=0 topMargin=4  oncontextmenu="return false">
<CENTER>
<FORM name=alipayment action="" method=post target="_blank">
<table>
 <tr>
   <td>
     <TABLE cellSpacing=0 cellPadding=0 width=740 border=0>
        <TR>
          <TD class=form-left>收款方： </TD>
          <TD class=form-star>* </TD>
          <TD class=form-right><%=mainname%>&nbsp;</TD>
        </TR>
        <TR>
          <TD colspan="3" align="center"><HR width=600 SIZE=2 color="#999999"></TD>
        </TR>
        <TR>
          <TD class=form-left>产品名称： </TD>
          <TD class=form-star>* </TD>
          <TD class=form-right><INPUT size=30 name=aliorder maxlength="200" value="<%=product_name%>" readonly></TD>
        </TR>
        <TR>
          <TD class=form-left>付款金额： </TD>
          <TD class=form-star>*</TD>
          <TD class=form-right><INPUT maxLength=10 size=30 name=alimoney value="<%=product_price%>" readonly/></TD>
        </TR>
        <TR>
          <TD class=form-left>备注：</TD>
          <TD class=form-star></TD>
          <TD class=form-right><TEXTAREA name=alibody rows=4 cols=40 wrap="physical" readonly><%=details%></TEXTAREA><input type="hidden" name="orderid" value="<%=orderid%>"/></TD>
        </TR>
        <TR>
          <TD class=form-left>支付方式：</TD>
          <TD class=form-star></TD>
          <TD class=form-right>
               <table>
                 <tr>
                   <td><input type="radio" name="pay_bank" value="directPay" checked><img src="images/alipay_1.gif" border="0"/></td>
                 </tr>
                 <tr>
                   <td><input type="radio" name="pay_bank" value="ICBCB2C"/><img src="images/ICBC_OUT.gif" border="0"/></td>
                   <td><input type="radio" name="pay_bank" value="CMB"/><img src="images/CMB_OUT.gif" border="0"/></td>
                   <td><input type="radio" name="pay_bank" value="CCB"/><img src="images/CCB_OUT.gif" border="0"/></td>
                   <td><input type="radio" name="pay_bank" value="BOCB2C"><img src="images/BOC_OUT.gif" border="0"/></td>
                 </tr>
                 <tr>
                   <td><input type="radio" name="pay_bank" value="ABC"/><img src="images/ABC_OUT.gif" border="0"/></td>
                   <td><input type="radio" name="pay_bank" value="COMM"/><img src="images/COMM_OUT.gif" border="0"/></td>
                   <td><input type="radio" name="pay_bank" value="SPDB"/><img src="images/SPDB_OUT.gif" border="0"/></td>
                   <td><input type="radio" name="pay_bank" value="GDB"><img src="images/GDB_OUT.gif" border="0"/></td>
                 </tr>
                 <tr>
                   <td><input type="radio" name="pay_bank" value="CITIC"/><img src="images/CITIC_OUT.gif" border="0"/></td>
                   <td><input type="radio" name="pay_bank" value="CEBBANK"/><img src="images/CEB_OUT.gif" border="0"/></td>
                   <td><input type="radio" name="pay_bank" value="CIB"/><img src="images/CIB_OUT.gif" border="0"/></td>
                   <td><input type="radio" name="pay_bank" value="SDB"><img src="images/SDB_OUT.gif" border="0"/></td>
                 </tr>
                 <tr>
                   <td><input type="radio" name="pay_bank" value="CMBC"/><img src="images/CMBC_OUT.gif" border="0"/></td>
                   <td><input type="radio" name="pay_bank" value="HZCBB2C"/><img src="images/HZCB_OUT.gif" border="0"/></td>
                   <td><input type="radio" name="pay_bank" value="SHBANK"/><img src="images/SHBANK_OUT.gif" border="0"/></td>
                   <td><input type="radio" name="pay_bank" value="NBBANK "><img src="images/NBBANK_OUT.gif" border="0"/></td>
                 </tr>
                 <tr>
                   <td><input type="radio" name="pay_bank" value="SPABANK"/><img src="images/SPABANK_OUT.gif" border="0"/></td>
                   <td><input type="radio" name="pay_bank" value="BJRCB"/><img src="images/BJRCB_OUT.gif" border="0"/></td>
                   <td><input type="radio" name="pay_bank" value="ICBCBTB"/><img src="images/ENV_ICBC_OUT.gif" border="0"/></td>
                   <td><input type="radio" name="pay_bank" value="CCBB2B"/><img src="images/ENV_CCB_OUT.gif" border="0"/></td>
                 </tr>
                 <tr>
                   <td><input type="radio" name="pay_bank" value="SPDBB2B"/><img src="images/ENV_SPDB_OUT.gif" border="0"/></td>
                   <td><input type="radio" name="pay_bank" value="ABCBTB"/><img src="images/ENV_ABC_OUT.gif" border="0"/></td>
				   <td></td>
				   <td></td>
                 </tr>
               </table>
          </TD>
        </TR>
         <TR>
          <TD class=form-left></TD>
          <TD class=form-star></TD>
          <TD class=form-right><a href="javascript:void(0);" onClick="alipayment.action='alipayto.asp';alipayment.submit();"><img src="images/button_sure.gif" name=nextstep/></a></TD>
        </TR>
</TABLE>
   </td>
   <td vAlign=top width=205 style="font-size:12px;font-family:'宋体'">
   <span id="glowtext">小贴士：</span>
   <fieldset>
      <P class=STYLE1>本通道为<a href="<%=show_url%>" target="_blank"><strong><%=mainname%></strong></a>客户专用，采用支付宝付款。请在支付前与本网站达成一致。</P>
      <P class="style2">请务必与<a href="<%=show_url%>" target="_blank"><strong><%=mainname%></strong></a>确认好订单和货款后，再付款。<br>
<span>务必看清卖家账号是否与你要支付的账号一致;如果不一致,请及时与卖家取得联系.</span></P>
      <P class="style2 style3">&nbsp;</P>
      </fieldset>
   </td>
 </tr>
</table>

</FORM>
</CENTER>
</BODY></HTML>
