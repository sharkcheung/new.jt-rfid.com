<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--#include file = inc2.asp -->
<!--#include file="func.asp"-->
<%
if session("u_id")="" then
   response.write "<script language=javascript>alert('登录超时,请重新登录!');window.top.location.href='/';</script>"
   response.End
end if
u_id=M_memberID(session("u_id"))%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>提交订单</title>
<meta http-equiv="x-ua-compatible" content="ie=7" />
<link href="style.css" rel="stylesheet" type="text/css" />
</head>
<body oncontextmenu="return false">
<div id="main" style="height:450px;">
<div class="member_center">
<TABLE cellSpacing=0 cellPadding=0 width=100% align=center border=0>
  <TBODY>
    <TR>
      <TD class=b vAlign=top align=left>
<%dim bookid,action,i
action=request("action")
set rs=server.CreateObject("adodb.recordset")
rs.open "select count(id) from u_shopcart where u_id="&u_id&"",connn,1,1
if rs(0)=0 then
response.write "<script language=javascript>alert('对不起，您购物车没有商品，请在购物后，再去“结算中心”！');window.parent.location.href='../';</script>"
response.End
end if
rs.close
set rs=nothing
'-----------------------
select case action
case ""
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from u_shopcart where u_id="&u_id&"",connn,1,1
%>
        <table width="100%" align="center" border="0" cellspacing="0" cellpadding="0" class="table-zuoyou" bordercolor="#CCCCCC">
          <tr>
            <td class="table-shangxia" background="images/class_bg.jpg" height=30>　<img src="images/ring02.gif" width="23" height="15" align="absmiddle"> 下订单 <font color="#FF6633"><b>(在最后结算前您还可以修改购物车内容)</b></font></td>
          </tr>
        </table>
        <table width="100%" align="center" border="0" cellspacing="0" cellpadding="0" class="table-zuoyou" bordercolor="#CCCCCC" style="border:#CCCCCC solid 1px;">
          <tr>
            <td bgColor=#ffffff colSpan=4 height=1></TD>
          </tr>
          <form name='form1' method='post' action="">
            <tr bgcolor="#f1f1f1" align="center">
              <td width=35% height="25">商品名称 </td>
              <td width=19%>单价</td>
              <td width=19%>数量</td>
              <td width=14%>总价</td>
            </tr>
            <tr>
              <td bgColor=#cccccc colSpan=4 height=1></TD>
            </tr>
            <tr>
              <td bgColor=#f1f1f1 colSpan=4 height=3></TD>
            </tr>
            <%shuliang=rs.recordcount
jianshu=0
zongji=0
do while not rs.eof%>
            <tr bgcolor="#ffffff">
              <td height="25" width="35%" class="table-xia"><%=rs("ProductTitle")%></a>
                  <input name=bookid type=hidden value="<%=rs("product_id")%>">
                  <input name=actionid type=hidden value="<%=rs("id")%>">
              </td>
              <td align="center" class="table-xia"><%=rs("danjia")%>
        元</td>
              <td align="center" class="table-xia"><%=rs("u_num")%></td>
              <td align="center" class="table-xia"><%=rs("u_num")*rs("danjia")%> 元</td>
            </tr>
            <%
jianshu=jianshu+rs("u_num")
total_price=total_price+rs("u_num")*rs("danjia")
rs.movenext
loop
rs.close
set rs=nothing%>
            <tr bgcolor=#ffffff align=center>
              <td height=30 colspan=4> 您的购物车里有商品：<%=jianshu%> 件　总数量：<%=jianshu%> 件　共计：<font color=red><%=total_price%></font> 元</td>
            </tr>
            <tr bgcolor=#ffffff align=center>
              <td height=40 colspan=4><input class="go-wenbenkuang" type="button" name="Submit" value="修改购物车" onClick="this.form.action='shop_cart.asp';this.form.submit()">
                  <input class="go-wenbenkuang" type="button"  onClick="this.form.action='?action=shop1';this.form.submit()" name="Submit3" value="确认订单 下一步">
              </td>
            </tr>
          </form>
        </table>
        <%
'-----------------------
case "shop1"
set rs=connn.execute("select * from u_members where id="&u_id&"")
userid=rs("m_uid")
%>
        <table width="100%" align="center" border="0" cellspacing="0" cellpadding="0" class="table-zuoyou" bordercolor="#CCCCCC">
          <tr>
            <td class="table-shangxia" background="images/class_bg.jpg" height=30>　<img src="images/ring02.gif" width="23" height="15" align="absmiddle"> 填写收货信息</td>
          </tr>
        </table>
        <table width="100%" align="center" border="0" cellspacing="0" cellpadding="0" class="table-zuoyou" bordercolor="#CCCCCC" style="border:#CCCCCC solid 1px;">
          <tr>
            <td bgColor=#ffffff colSpan=2 height=1></TD>
          </tr>
          <tr bgcolor="#ffffff">
            <td bgColor="#f1f1f1" colspan="2" height="25" align="center">请正确填写以下收货信息</td>
          </tr>
          <tr>
            <td bgColor=#cccccc colSpan=2 height=1></TD>
          </tr>
          <tr>
            <td bgColor=#f1f1f1 colSpan=2 height=3></TD>
          </tr>
          <form name="shouhuoxx" method="post" action="" onSubmit="ssxx">
            <tr bgcolor="#ffffff">
              <td width="30%" height="25" align="right" class="table-xia">收货人真实姓名：</td>
              <td width="70%" height="25" style="PADDING-LEFT: 20px" class="table-xia"><input name=userid type=hidden value="<%=userid%>">
                  <input name="userzhenshiname" class="wenbenkuang" type="text" id="userzhenshiname" size="16" value=<%=trim(rs("m_uname"))%>>
        性别：
        <select class="wenbenkuang" name="shousex" id="shousex">
          <option value=0 <%if rs("m_usex")=0 then%>selected<%end if%>>男</option>
          <option value=1 <%if rs("m_usex")=1 then%>selected<%end if%>>女</option>
        </select>
              </td>
            </tr>
            <tr bgcolor="#ffffff">
              <td width="30%" height="25" align="right" class="table-xia">详细地址：</td>
              <td width="70%" height="25" style="PADDING-LEFT: 20px" class="table-xia">
			  <input class="wenbenkuang" name="shouhuodizhi" type="text" id="shouhuodizhi" size="60" value=<%=trim(rs("m_uaddress"))%>>
              </td>
            </tr>
            <tr bgcolor="#ffffff">
              <td width="30%" height="25" align="right" class="table-xia">邮政编码：</td>
              <td width="70%" height="25" style="PADDING-LEFT: 20px" class="table-xia"><input class="wenbenkuang" name="youbian" type="text" id="youbian" size="16" value="<%=rs("m_uzip")%>" ONKEYPRESS="event.returnValue=IsDigit();"></td>
            </tr>
            <tr bgcolor="#ffffff">
              <td width="30%" height="25" align="right" class="table-xia">手机号码：</td>
              <td width="70%" height="25" style="PADDING-LEFT: 20px" class="table-xia">
			  <input class="wenbenkuang" name="usertel" type="text" id="usertel" size="30" value='<%=rs("m_umobile")%>' style="height: 20px">
              </td>
            </tr>
            <tr bgcolor="#ffffff">
              <td width="30%" height="25" align="right" class="table-xia">电子邮件：</td>
              <td width="70%" height="25" style="PADDING-LEFT: 20px" class="table-xia"><input class="wenbenkuang" name="useremail" type="text" id="useremail" size="30" value=<%=trim(rs("m_uemail"))%>>
              </td>
            </tr>
            <tr bgcolor="#ffffff">
              <td width="30%" height="25" align="right" class="table-xia">送货方式：</td>
              <td width="70%" height="25" style="PADDING-LEFT: 20px" class="table-xia"><%set rs6=server.CreateObject("adodb.recordset")
rs6.open "select * FROM Iheeo_Delivery order by SongList",connn,1,1
%><select name="songhuofangshi" size=3 style="WIDTH: 180px">
        <%do while not rs6.eof%>
        <option value="<%=rs6("SongKey")%>" ><%=trim(rs6("SongName"))%></option><%rs6.movenext
loop
rs6.close
set rs6=nothing%>
</select></td>
            </tr>
            <tr bgcolor="#ffffff">
              <td width="30%" height="25" align="right" class="table-xia">支付方式：</td>
              <td width="70%" height="25" style="PADDING-LEFT: 20px" class="table-xia"><%set rs5=server.CreateObject("adodb.recordset")
rs5.open "select * FROM Iheeo_Pay order by PayList",connn,1,1
%><select name="zhifufangshi" size=3 style="WIDTH: 180px">
        <%do while not rs5.eof%>
        <option value="<%=rs5("PayKey")%>"><%=trim(rs5("PayName"))%></option><%rs5.movenext
loop
rs5.close
set rs5=nothing%>
<!--<option value="99999" >银行付款</option>
<option value="100000" >货到付款</option>--></select></td>
            </tr>
            <tr bgcolor="#ffffff">
              <td height="40" colspan="2" align=center><input class="go-wenbenkuang" type="button" name="Submit2" value="上一步" onClick="javascript:history.go(-1)">
                  <input class="go-wenbenkuang" type="button" name="Submit4" value="确认收货信息 下一步" onclick='return ssxx();'>
              </td>
            </tr>
          </form>
        </table>
        <SCRIPT LANGUAGE="JavaScript">
<!--
function IsDigit()
{
  return ((event.keyCode >= 48) && (event.keyCode <= 57));
}
function checkspace(checkstr) {
  var str = '';
  for(i = 0; i < checkstr.length; i++) {
    str = str + ' ';
  }
  return (str == checkstr);
}

function ssxx()
{
   if(checkspace(document.shouhuoxx.userzhenshiname.value)) {
	document.shouhuoxx.userzhenshiname.focus();
    alert("对不起，请填写收货人姓名！");
	return false;
  }
  if(checkspace(document.shouhuoxx.shouhuodizhi.value)) {
	document.shouhuoxx.shouhuodizhi.focus();
    alert("对不起，请填写收货人详细收货地址！");
	return false;
  }
  if(checkspace(document.shouhuoxx.youbian.value)) {
	document.shouhuoxx.youbian.focus();
    alert("对不起，请填写邮编！");
	return false;
  }
  if(document.shouhuoxx.youbian.value.length!=6) {
	document.shouhuoxx.youbian.focus();
    alert("对不起，请正确填写邮编！");
	return false;
  } 
    if(checkspace(document.shouhuoxx.usertel.value)) {
	document.shouhuoxx.usertel.focus();
    alert("对不起，请留下您的电话！");
	return false;
  }
    if(checkspace(document.shouhuoxx.songhuofangshi.value)) {
	document.shouhuoxx.songhuofangshi.focus();
    alert("对不起，请选择送货方式！");
	return false;
  }
    if(checkspace(document.shouhuoxx.zhifufangshi.value)) {
	document.shouhuoxx.zhifufangshi.focus();
    alert("对不起，请选择支付方式！");
	return false;
  }
  if(document.shouhuoxx.useremail.value.length!=0)
  {
    if (document.shouhuoxx.useremail.value.charAt(0)=="." ||        
         document.shouhuoxx.useremail.value.charAt(0)=="@"||       
         document.shouhuoxx.useremail.value.indexOf('@', 0) == -1 || 
         document.shouhuoxx.useremail.value.indexOf('.', 0) == -1 || 
         document.shouhuoxx.useremail.value.lastIndexOf("@")==document.shouhuoxx.useremail.value.length-1 || 
         document.shouhuoxx.useremail.value.lastIndexOf(".")==document.shouhuoxx.useremail.value.length-1)
     {
      alert("Email地址格式不正确！");
      document.shouhuoxx.useremail.focus();
      return false;
      }
   }
 else
  {
   alert("Email不能为空！");
   document.shouhuoxx.useremail.focus();
   return false;
   }
    document.shouhuoxx.action='?action=shop2';
	document.shouhuoxx.submit();
}
//-->
        </script>
        <%
rs.close
set rs=nothing
'-----------------------
case "shop2"
shijian=now()
dingdan="D-"&year(shijian)&month(shijian)&day(shijian)&hour(shijian)&minute(shijian)&second(shijian)
%>
        <table width="100%" align="center" border="0" cellspacing="0" cellpadding="0" class="table-zuoyou" bordercolor="#CCCCCC">
          <tr>
            <td class="table-shangxia" background="images/class_bg.jpg" height=30>　<img src="images/ring02.gif" width="23" height="15" align="absmiddle"> 提交订单</td>
          </tr>
        </table>
        <table width="100%" align="center" border="0" cellspacing="0" cellpadding="0" class="table-zuoyou" bordercolor="#CCCCCC">
          <tr>
            <td bgColor=#ffffff height=1></TD>
          </tr>
          <tr bgcolor="#ffffff">
            <td bgColor="#f1f1f1" height="25" align="center">请确认您填写的订单以便收货|订单号：<span style="color:#E14900;font-weight:bolder;font-size:14px;"><%= dingdan%></span> </td>
          </tr>
          <tr>
            <td bgColor=#cccccc height=1></TD>
          </tr>
          <tr>
            <td bgColor=#f1f1f1 height=3></TD>
          </tr>
          <tr>
            <td><table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
                <tr bgcolor="#ffffff">
                  <td align="center" valign="top" height=50 colspan=2><%set rs=server.CreateObject("adodb.recordset")
rs.open "select * from u_shopcart where u_id="&u_id&"",connn,1,1 
%>
                      <div align="center">
                      <table width=99% border=0 cellspacing=1 bgcolor=#cccccc>
                        <tr align=center bgcolor=#f1f1f1>
                          <td width=39%>商品名称</td>
                          <td width=15%>价格</td>
                          <td width=15%>数 量</td>
                          <td width=15%>总 价</td>
                        </tr>
                        <%shuliang=rs.recordcount
jianshu=0
zongji=0
fudongjia=0
bjxbookname=""

for i=1 to rs.recordcount
bjxbookname=bjxbookname&rs("ProductTitle")
if i<>rs.recordcount then bjxbookname=bjxbookname&"+"
prodid=prodid&rs("product_id")&","&rs("u_num")&"|"
pronum=pronum&rs("u_num")&","
%>
                        <tr align="center" bgcolor=#ffffff>
                          <td height="22" align="left">　<%=rs("ProductTitle")%>
                              <input name=bookid2 type=hidden value="<%=rs("product_id")%>">
                              <input name=actionid2 type=hidden value="<%=rs("id")%>">
                          </td>
                          <td><%=rs("danjia")%>
                  元</td>
                          <td><%=rs("u_num")%></td>
                          <td><%=rs("u_num")*rs("danjia")%> 元</td>
                        </tr>
<%
jianshu=jianshu+rs("u_num")
total_price=total_price+rs("u_num")*rs("danjia")
'算每件商品的浮动价
rs.movenext
next
rs.close
set rs=nothing
%>
                        <tr bgcolor=#ffffff>
                          <td height=22 colspan=4 align=center>商品总计：<font style="color:#E14900;font-size:18px; font-weight:bolder;"><%=total_price%></font> 元</td>
                        </tr>
                      </table>
                  </div>
                  </td>
                </tr>
                <tr bgcolor="#ffffff">
                  <td width="50%" align="center" valign="top">
					<table width="99%" border="0" cellspacing="1" bgcolor="#CCCCCC">
                      <tr>
                        <td colspan="2" height="24" bgcolor="#f1f1f1" align="center">您的订单—收货信息</td>
                      </tr>
                      <tr bgcolor="ffffff">
                        <td width="30%" align="center">姓名：</td>
                        <td width="70%" style="PADDING-LEFT: 20px;text-align:left;"><%=request("userzhenshiname")%></td>
                      </tr>
                      <tr bgcolor="ffffff">
                        <td align="center">邮编：</td>
                        <td style="PADDING-LEFT: 20px;text-align:left;"><%=request("youbian")%></td>
                      </tr>
                      <tr bgcolor="ffffff">
                        <td align="center">地址：</td>
                        <td style="PADDING-LEFT: 20px;text-align:left;"><%=request("shouhuodizhi")%></td>
                      </tr>
                      <tr bgcolor="ffffff">
                        <td align="center">电话：</td>
                        <td style="PADDING-LEFT: 20px;text-align:left;"><%=request("usertel")%></td>
                      </tr>
                      <tr bgcolor="ffffff">
                        <td align="center">邮箱：</td>
                        <td style="PADDING-LEFT: 20px;text-align:left;"><%=request("useremail")%></td>
                      </tr>
                      <tr bgcolor="ffffff">
                        <td align="center">送货：</td>
                        <td style="PADDING-LEFT: 20px;text-align:left;"><%
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from Iheeo_Delivery where SongKey="&request("songhuofangshi"),connn,1,1
if rs.eof and rs.bof then
response.write "方式已经被删除"
else
response.write rs("SongName")
end if
rs.close
set rs=nothing%>
                        </td>
                      </tr>
                      <tr bgcolor="ffffff">
                        <td align="center">支付：</td>
                        <td style="PADDING-LEFT: 20px;text-align:left;"><%
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from Iheeo_Pay where PayKey="&request("zhifufangshi"),connn,1,1
if rs.eof and rs.bof then
response.write "方式不存在已删除"
else
response.write rs("PayName")

end if
rs.close
set rs=nothing%>
                        </td>
                      </tr>
                  </table></td>
                  <td width="50%" align="center" valign="top"><table width="99%" border="0" cellpadding="2" cellspacing="1" bgcolor="#CCCCCC">
                      <tr>
                        <td height="24" bgcolor="#f1f1f1" align="center">总费用计算</td>
                      </tr>
                      <%
'计算费用
'先取出参数
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from Iheeo_Delivery where SongKey="&request("songhuofangshi"),connn,1,1
SongFei=rs("SongFei")
	
	feiyong=SongFei
		if request("zhifufangshi")=71 then
          '得到预存款
          set rs2=server.CreateObject("adodb.recordset")
          rs2.Open "select yucun from bjx_User where username='"&username&"'",connn,1,1
          yucunkuan=rs2("yucun")
          rs2.close
          set rs2=nothing
	  if yucunkuan<feiyong+zongji then
	  	response.write "<script language=javascript>alert('您的预存款不足，请更换支付方式！');history.go(-1);</script>"
	end if
	end if%>
                      <tr>
                        <td style="PADDING-right: 20px; font-size:14px;" align="right" bgcolor="ffffff">商品总价：<font style="color:#E14900;font-size:18px; font-weight:bolder;"><%=FormatNumber(total_price,2)%></font> 元<br>您的送货费用计：<font style="color:#E14900;font-size:18px; font-weight:bolder;"><%=FormatNumber(feiyong,2)%></font> 元</td>
                      </tr>
                      <tr>
                        <td style="PADDING-right: 20px;" align="right" bgcolor="ffffff"><font style="color:#E14900;font-size:18px; font-weight:bolder;">您的订单总金额： <%=FormatNumber(total_price+feiyong,2)%> 元</font></td>
                      </tr>
                      <tr>
                        <td style="PADDING-right: 20px;" align="right" bgcolor="ffffff">
                        <form name="shouhuoxx2" method="post" action=" ">
                  <tr bgcolor="#ffffff" align="center">
                    <td colspan="2"><table width="95%" border="0" cellspacing="0" cellpadding="2">
                        <tr>
                          <td width="20%"><!--<input type="checkbox" name="fapiao" value="1">
                    是否要发票？-->
                      <input name="dingdan" type="hidden" value=<%=trim(dingdan)%>>
                      <input name="userzhenshiname" type="hidden" value=<%=trim(request("userzhenshiname"))%>>
                      <input name="shousex" type="hidden" value=<%=trim(request("shousex"))%>>
                      <input name="useremail" type="hidden" value=<%=trim(request("useremail"))%>>
                      <input name="shouhuodizhi" type="hidden" value=<%=trim(request("shouhuodizhi"))%>>
                      <input name="youbian" type="hidden" value=<%=trim(request("youbian"))%>>
                      <input name="usertel" type="hidden" value=<%=trim(request("usertel"))%>>
                      <input name="songhuofangshi" type="hidden" value=<%=trim(request("songhuofangshi"))%>>
                      <input name="zhifufangshi" type="hidden" value=<%=trim(request("zhifufangshi"))%>>
                      <input name="feiyong" type="hidden" value=<%=feiyong%>>
                      <input name="zongji" type="hidden" value=<%=total_price%>>
                      <input name="bookid2" type="hidden" value=<%=prodid%>>
                      <input name=userid type=hidden value="<%=request("userid")%>" >
					  <input name="pronum" type="hidden" value=<%=pronum%>>
					  <input name="bjxbookname" type="hidden" value="<%=bjxbookname%>">
					  <input name="jianshu" type="hidden" value=<%=jianshu%>>
					  <input name="SongKey" type="hidden" value=<%=SongKey%>>
                          </td>
                          <td width="60%"><input class="go-wenbenkuang" type="button" name="Submit22" value="上一步" onClick="javascript:history.go(-1)"><!--<input class="wenbenkuang" type="text" name="liuyan" size="35" maxlength="30">
                    您对此订单的特殊说明(30字内) --></td>
                          <td width="20%" align="center" height="60">
                              <input class="go-wenbenkuang" type="button" onClick="this.form.action='?action=ok';this.form.submit()"  name="Submit42" value="完成订单并支付">
                          </td>
                        </tr>
                    </table></td>
                  </tr>
                </form>
                        </td>
                      </tr>
</table></td>
                </tr>
                
            </table></td>
          </tr>
        </table>
        <span class="m1-foot">
        <%
case "ok"
'-----------------------
function HTMLEncode2(fString)
	fString = Replace(fString, CHR(13), "")
	fString = Replace(fString, CHR(10) & CHR(10), "</P><P>")
	fString = Replace(fString, CHR(10), "<BR>")
	HTMLEncode2 = fString
end function
'修改用户的送货信息
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from u_members where id="&u_id&"",connn,1,3
if not rs.eof then
rs("m_uname")=trim(request("userzhenshiname"))
rs("m_usex")=trim(request("shousex"))
rs("m_uemail")=trim(request("useremail"))
rs("m_uaddress")=trim(request("shouhuodizhi"))
rs("m_uzip")=trim(request("youbian"))
rs("m_umobile")=trim(request("usertel"))
rs.update
end if
rs.close
set rs=nothing

if session("xiadan")<>minute(now) then
'再判断库存
'未写

dim shijian,dingdan,zongji,feiyong
dingdan=trim(request("dingdan"))
zongji=trim(request("zongji"))
feiyong=request("feiyong")
shijian=now()
jianshu=Request("jianshu")
pronum=Request("pronum")
p_num=split(pronum,",")
bookid2=trim(request("bookid2"))
bookid=split(bookid2,",")
'for tt=0 to ubound(bookid)
'connn.execute("update BJX_goods set kucun=kucun-"&p_num(tt)&" where bookid="&bookid(tt))
'next
session("pro_pro_xj")=(zongji-0+feiyong)*100
session("pro_pro_num")=jianshu
session("pro_paytype")=int(request("zhifufangshi"))
session("pid")=request("bjxbookname")
session("pro_id")=dingdan
session("pro_fee")=feiyong
session("pro_pro_price")=zongji
session("pro_contact")=trim(request("userzhenshiname"))
session("pro_mobi")=trim(request("usertel"))
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from u_order where u_id="&u_id&"",connn,1,3
if rs.eof then
'得到价格，减库存
rs.addnew
if not isarray(bookid) then
rs("p_id")=bookid2
else
rs("p_id")=0
rs("pro_ids")=bookid2
end if
rs("order_time")=shijian
rs("pro_price")=zongji
if request("zhifufangshi")=100000  then    '送货上门或预存款支付，直接改为订单完成（已收到款）
rs("pro_paystatu")=4
else
rs("pro_paystatu")=0
end if
rs("order_id")=dingdan
rs("pro_post")=int(request("youbian"))
rs("pro_contact")=trim(request("userzhenshiname"))
rs("pro_add")=trim(request("shouhuodizhi"))
rs("pro_paytype")=int(request("zhifufangshi"))
rs("pro_num")=jianshu
rs("pro_sex")=int(request("shousex"))
rs("pro_message")=HTMLEncode2(trim(request("liuyan")))
rs("pro_contact")=trim(request("userzhenshiname"))
rs("pro_tel")=trim(request("usertel"))
rs("u_id")=u_id
'新增
if request("fapiao")<>1 then 
fapiao=0
else
fapiao=1
end if
rs("pro_tax")=fapiao
rs("pro_fee_type")=request("songhuofangshi")
rs.update
else
   if rs("order_id")<>dingdan then
      rs.addnew
if not isarray(bookid) then
rs("p_id")=bookid2
else
rs("p_id")=0
rs("pro_ids")=bookid2
end if
rs("order_time")=shijian
rs("pro_price")=zongji
if request("zhifufangshi")=100000  then    '送货上门或预存款支付，直接改为订单完成（已收到款）
rs("pro_paystatu")=4
else
rs("pro_paystatu")=0
end if
rs("order_id")=dingdan
rs("pro_post")=int(request("youbian"))
rs("pro_contact")=trim(request("userzhenshiname"))
rs("pro_add")=trim(request("shouhuodizhi"))
rs("pro_paytype")=int(request("zhifufangshi"))
rs("pro_num")=jianshu
rs("pro_sex")=int(request("shousex"))
rs("pro_message")=HTMLEncode2(trim(request("liuyan")))
rs("pro_contact")=trim(request("userzhenshiname"))
rs("pro_tel")=trim(request("usertel"))
rs("u_id")=u_id
'新增
if request("fapiao")<>1 then 
fapiao=0
else
fapiao=1
end if
rs("pro_tax")=fapiao
rs("pro_fee_type")=request("songhuofangshi")
rs.update
   else
      response.Write "<script language=javascript>alert('您不能重复提交!!');window.parent.location.href='../';</script>"
	  response.end
   end if
'rs.close
'set rs=nothing
'connn.execute "delete from BJX_action where username='"&request.cookies("bookshop")("username")&"' and bookid in ("&bookid&") and zhuangtai=6"
end if
rs.close
set rs=nothing
session("xiadan")=minute(now)
else
      response.Write "<script language=javascript>alert('您不能重复提交!');window.parent.location.href='../';</script>"
	  response.end
end if
%>
        </span>
<%
connn.execute("delete from u_shopcart where u_id="&M_memberID(session("u_id"))&"")
'清空购物车
%>
        <table width="100%" align="center" border="0" cellspacing="0" cellpadding="0" class="table-zuoyou" bordercolor="#CCCCCC">
          <tr>
            <td class="table-shangxia" background="images/class_bg.jpg" height=30>　<img src="images/ring02.gif" width="23" height="15" align="absmiddle"> 订单提交成功</td>
          </tr>
        </table>
        <table width="100%" align="center" border="0" cellspacing="0" cellpadding="0" class="table-zuoyou" bordercolor="#CCCCCC">
          <tr>
            <td bgColor=#ffffff height=1></TD>
          </tr>
          <tr bgcolor="#ffffff">
            <td bgColor="#f1f1f1" height="25" align="center">您的订单已经成功提交，我们会在第一时间进行处理，请记清您的订单号以备查询。</td>
          </tr>
          <tr>
            <td bgColor=#cccccc height=1></TD>
          </tr>
          <tr>
            <td bgColor=#f1f1f1 height=3></TD>
          </tr>
          <tr>
            <td height="25" bgcolor="ffffff" style="PADDING-LEFT: 100px">订单号：<font color=red><%=dingdan%></font></td>
          </tr>
		  <%if u_id<>"" then%>
          <tr>
            <td height="25" bgcolor="ffffff" style="PADDING-LEFT: 100px">订单查询：您可通过“用户中心”&gt;&gt;“我的订单”查询您的订单状态。</td>
          </tr>
<!--          <tr>
            <td height="60" bgcolor="ffffff" style="PADDING-LEFT: 100px">购物积分：请在收货后通过“<a href="javascript:;" onClick="javascript:window.open('user.asp','','')">用户中心</a>”&gt;&gt;“<a href="javascript:;" onClick="javascript:window.open('member_center.asp?action=dindan','','')">我的订单</a>”及时更改您的订单状态为“完成”<br>
			因为每笔订单的积分只有在订单完成后才能累计到您的购物积分中。 </td>
          </tr>-->
		  <%else%>
		  <tr>
            <td height="25" bgcolor="ffffff" style="PADDING-LEFT: 100px">订单查询：因为您不是我们的会员，所以您不能查询您的订单状态。</td>
          </tr>
<!--		  <tr>
            <td height="25" bgcolor="ffffff" style="PADDING-LEFT: 100px">购物积分：因为您不是我们的会员，所以您不能获得积分奖励。 </td>
          </tr>-->
          <%
		  end if
		  if request("zhifufangshi")=99999 then %>
          <tr>
            <td height="25" bgcolor="ffffff" style="PADDING-LEFT: 100px">您是通过银行汇款方式支付的，请您在一周内依照您选择的支付方式进行汇款！汇款时请注明您的<font color="#FF0000">订单号</font>！</td>
          </tr>
          <%else
if request("zhifufangshi")=100000 then %>
          <tr>
            <td height="25" bgcolor="ffffff" style="PADDING-LEFT: 100px">您是选择的“货到付款”，我们会尽快给您送货的！</td>
          </tr>
          <%else%>
          <tr>
            <td height="25" bgcolor="ffffff" style="PADDING-LEFT: 100px">请您在一周内依照您选择的支付方式进行汇款，汇款时请注明您的<font color="#FF0000">订单号</font>！</td>
          </tr><%
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from Iheeo_Pay where PayKey="&request("zhifufangshi"),connn,1,1
PayKey=rs("PayKey")
rs.close
set rs=nothing
select case request("zhifufangshi")
   case 1
      payimg="images/paytype/kq.gif"
   case 2
      payimg="images/paytype/zfb.gif"
   case 3
      payimg="images/paytype/tenpay.gif"
end select
%>
          <tr>
            <td height="25" bgcolor="ffffff" align="center"><form name="onlinepay" id="onlinepay" action="" method="post" target="_self" >
<input type="hidden" name="orderid" value="<%=dingdan%>">
<input type="hidden" name="totalmoney" value="<%=money%>">
<a href="javascript:void(0);" title="点出支付" onClick="onlinepay.action='pay.asp';onlinepay.submit();return false;"><img src="<%=payimg%>" border="0" /></a><br>
<a href="javascript:void(0);" title="点出支付" onClick="onlinepay.action='pay.asp';onlinepay.submit();return false;"><img src="images/fkkk.gif" border="0" title="点出支付"></a></FORM></td>
          </tr>
          <%end if%>
          <%end if%><tr>
            <td height="25" bgcolor="ffffff" style='PADDING-LEFT: 100px' align="right"><font color="#999999">　订单提交完成 创建时间：<%=shijian%>&nbsp;</font> </td>
          </tr>
        </table>
        <%
		response.Cookies("bjx")("dingdanusername")=""
		end select%></TD>
    </TR>
  </TBODY>
</TABLE></div>
<div style=" clear:both;"></div>

</div>
</body>
</html>
<script language=javascript>
<!--
function regInput(obj, reg, inputStr)
{
	var docSel	= document.selection.createRange()
	if (docSel.parentElement().tagName != "INPUT")	return false
	oSel = docSel.duplicate()
	oSel.text = ""
	var srcRange	= obj.createTextRange()
	oSel.setEndPoint("StartToStart", srcRange)
	var str = oSel.text + inputStr + srcRange.text.substr(oSel.text.length)
	return reg.test(str)
}
function checkspace(checkstr) {
  var str = '';
  for(i = 0; i < checkstr.length; i++) {
    str = str + ' ';
  }
  return (str == checkstr);
}
   //-->
</script>