<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<!--#include file = inc2.asp -->
<!--#include file="func.asp"-->
<%
bookid=Decrypt(trim(request("pid")))
num=trim(request("buy_num"))
u_name=session("u_id")
if u_name="" then
   response.Write "<script language=javascript>alert('请先登录再购买!');history.back();</script>"
   response.end
end if
if bookid="" then
   response.Redirect "./"
end if
set bookrs=connn.execute("select * from BJX_goods where bookid="&bookid&"")
if bookrs.eof then
   bookrs.close
   set bookrs=nothing
   response.Write "<script language=javascript>alert('无此商品信息!');window.location.href='./';</script>"
   response.end
else
   zhuang=bookrs("zhuang")
   bookname=bookrs("bookname")
   bookid=bookrs("bookid")
   shichangjia=bookrs("shichangjia")
   bookchuban=bookrs("bookchuban")
   huiyuanjia=bookrs("huiyuanjia")
   dazhe=bookrs("dazhe")
   kucun=bookrs("kucun")
   bookcontent=bookrs("bookcontent")
   bookinfo=bookrs("bookinfo")
   
   bookother1=bookrs("bookother1")
   bookother2=bookrs("bookother2")
   bookother3=bookrs("bookother3")
   bookother4=bookrs("bookother4")
   pingpai=bookrs("pingpai")
end if
bookrs.close
set bookrs=nothing
if zhuang="" then 
   zhuang"images/emptybook.gif"
end if


set bookrs=connn.execute("select * from u_members where m_uid='"&u_name&"'")
if bookrs.eof then
   bookrs.close
   set bookrs=nothing
   response.Write "<script language=javascript>window.location.href='./';</script>"
   response.end
else
resume_hukouprovinceid=bookrs("szSheng")
resume_hukoucapitalid=bookrs("szShi")
resume_hukoucityid=bookrs("szXian")
m_uaddress=bookrs("m_uaddress")
m_uname=bookrs("m_uname")
m_utel=bookrs("m_utel")
m_umobile=bookrs("m_umobile")
m_uzip=bookrs("m_uzip")
end if
bookrs.close
set bookrs=nothing
%>
<head>
<meta content="text/html; charset=gb2312" http-equiv="Content-Type" />
<meta content="<%=keywords%>" name="keywords" />
<meta content="<%=description%>" name="description" />
<title><%=title%>-<%=lmname%>-<%=company%></title>
<link href="style.css" rel="stylesheet" type="text/css" />
<script language = "JavaScript" src="js/GetProvince.js"></script>
<script src="js/shop_cart.js"></script>
</head>

<body onload="initt(<%=num%>,<%=huiyuanjia%>);">
<div id="main">
    <div class="left_buy">
		<h3><span>商品信息</span></h3>
	<div class="listnews2">
		<div class="product_detail">
                <div class="pic">
			      <a href="product.asp?bookid=<%=server.URLEncode(Encrypt(bookid))%>"><img src="<%=zhuang%>" border="0" /></a>
			   </div>
               <div class="title">
			   <h4>商品名称：<a href="product.asp?bookid=<%=server.URLEncode(Encrypt(bookid))%>"><%=bookname%> <%=bookad%></a></h4>
			   <h2>市场价：&yen;<s><%=formatnumber(shichangjia,2,true)%>元</s>/<%=bookchuban%></h2>
			   <h2>会员价：<b><font color="#93393A" style="font-size:11pt">&yen;<%=formatnumber(huiyuanjia,2,true)%>元</font></b></h2>
               </div>
			  
			</div>
	</div>
	</div>
	<div class="right_buy">
		<div class="listnews">
			<div class="product_detail">
               <div class="title">
			      <form name="gobuy" action="buy_save.asp" method="post" onsubmit="return chk_order();">
				    <table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td colspan="8" class="capt">确认收货地址
                        <input type="hidden" name="pnjd_did" value="<%=bookid%>"/></td>
                      </tr>
                      <tr>
					    <td height="10" bgcolor="#FFFFFF">　</td>
                        <td width="142" height="10" bgcolor="#FFFFFF">　</td>
                        <td width="38" height="10" bgcolor="#FFFFFF">　</td>
                        <td width="147" height="10" bgcolor="#FFFFFF">　</td>
                        <td width="37" height="10" bgcolor="#FFFFFF">　</td>
                        <td width="130" height="10" bgcolor="#FFFFFF">　</td>
                        <td height="10" bgcolor="#FFFFFF">　</td>
                        <td height="10" bgcolor="#FFFFFF">　</td>
                      </tr>
                      <tr>
                        <td width="93" align="right" bgcolor="#D9F1F9">省：</td>
                        <td colspan="5" bgcolor="#D9F1F9">
						<select name="hukouprovince" size="1" id="select5" onChange="changeProvince(document.gobuy.hukouprovince.options[document.gobuy.hukouprovince.selectedIndex].value)">
		<%if resume_hukouprovinceid<>"" then%>
		<option value="<%=resume_hukouprovinceid%>"><%=Hireworkadds(resume_hukouprovinceid)%></option>
		<%else%>
		<option value="">选择省</option>
		<%end if%>
		</select>市：
						<select name="hukoucapital" onchange="changeCity(document.gobuy.hukoucapital.options[document.gobuy.hukoucapital.selectedIndex].value)">
						  <%if resume_hukoucapitalid<>"" then%>
						  <option value="<%=resume_hukoucapitalid%>"><%=Hireworkadds(resume_hukoucapitalid)%></option>
						  <%else%>
						  <option value="">选择市</option>
						  <%end if%>
	                      </select>县：
		                  <select name="hukoucity">
		                    <%if resume_hukoucityid<>"" then%>
		                    <option value="<%=resume_hukoucityid%>"><%=Hireworkadds(resume_hukoucityid)%></option>
		                    <%else%>
		                    <option value="">选择区</option>
		                    <%end if%>
                          </select></td>
                        <td width="87" bgcolor="#D9F1F9">邮政编码：</td>
                        <td width="58" bgcolor="#D9F1F9"><input name="zip" type="text" value="<%= m_uzip %>" onkeyup="this.value=this.value.replace(/\D/g,'');" size="6" maxlength="6" /></td>
                      </tr>
                      <tr>
                        <td align="right" valign="top" bgcolor="#D9F1F9">街道地址：</td>
                        <td colspan="7" bgcolor="#D9F1F9"><textarea name="detail_address" cols="50" rows="3"><%= m_uaddress %></textarea></td>
                      </tr>
                      <tr>
                        <td align="right" bgcolor="#D9F1F9">收货人姓名：</td>
                        <td bgcolor="#D9F1F9"><input name="shouhuo_name" type="text" value="<%= m_uname %>" size="20" maxlength="20" /></td>
                        <td bgcolor="#D9F1F9">　</td>
                        <td bgcolor="#D9F1F9">　</td>
                        <td bgcolor="#D9F1F9">　</td>
                        <td bgcolor="#D9F1F9">　</td>
                        <td bgcolor="#D9F1F9">　</td>
                        <td bgcolor="#D9F1F9">　</td>
                      </tr>
                      <tr>
                        <td align="right" bgcolor="#D9F1F9">手机：</td>
                        <td bgcolor="#D9F1F9"><input name="shouhuo_mobile" type="text" value="<%= m_umobile %>" size="20" maxlength="20" /></td>
                        <td bgcolor="#D9F1F9">　</td>
                        <td bgcolor="#D9F1F9">　</td>
                        <td bgcolor="#D9F1F9">　</td>
                        <td bgcolor="#D9F1F9">　</td>
                        <td bgcolor="#D9F1F9">　</td>
                        <td bgcolor="#D9F1F9">　</td>
                      </tr>
                      <tr>
					    <td align="right" bgcolor="#D9F1F9">电话：</td>
                        <td bgcolor="#D9F1F9"><input name="shouhuo_tel" type="text" value="<%= m_utel %>" size="20" maxlength="20" /></td>
                        <td bgcolor="#D9F1F9">　</td>
                        <td bgcolor="#D9F1F9">　</td>
                        <td bgcolor="#D9F1F9">　</td>
                        <td bgcolor="#D9F1F9">　</td>
                        <td bgcolor="#D9F1F9">　</td>
                        <td bgcolor="#D9F1F9">　</td>
                      </tr>
                    </table>
					<table cellspacing="0" cellpadding="0" width="100%" align="center" border="0">
      <tr>
        <td class="capt">选择支付方式：</td>
      </tr>
      <tr>
        <td><table cellspacing="0" cellpadding="2" width="530" align="center" border="0">
          <tr>
            <td><input name="mainRadio" type="radio" id="OnlineBank" style="BORDER-RIGHT: 0px; BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px" title="Father" onclick="document.getElementById('div_OnlineBank').style.display='block'" value="0"/>
                    <strong><font color="#222222" size="3"> 网上支付平台</font></strong><font color="#999999">(支持支付宝、快钱等)：</font></td>
          </tr>
          <tr>
            <td><div id="div_OnlineBank" style="DISPLAY:none ;padding-left:20px;">
              <table cellspacing="0" cellpadding="0" width="96%" align="right" border="0">
                <tr>
                  <td valign="top"><%call pay_type()%></td>
                </tr>
              </table>
            </div></td>
          </tr>
          <tr>
            <td><table height="1" cellspacing="0" cellpadding="0" width="95%" align="right" border="0">
              <tr>
                <td bgcolor="#b1b1b1"></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td><input id="Bank" style="BORDER-RIGHT: 0px; BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px" onclick="document.getElementById('div_OnlineBank').style.display='none'" type="radio" value="99999" name="mainRadio"/>
                    <strong><font color="#222222" size="3"> 银行汇款</font></strong></td>
          </tr>
          <tr>
            <td><table height="1" cellspacing="0" cellpadding="0" width="95%" align="right" border="0">
              <tr>
                <td bgcolor="#b1b1b1"></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td><input id="Post" style="BORDER-RIGHT: 0px; BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px" onclick="document.getElementById('div_OnlineBank').style.display='none'" type="radio" value="100000" name="mainRadio"/>
                    <strong><font color="#222222" size="3"> 货到付款</font></strong></td>
          </tr>
          <tr>
            <td><table height="1" cellspacing="0" cellpadding="0" width="95%" align="right" border="0">
              <tr>
                <td bgcolor="#b1b1b1"></td>
              </tr>
            </table></td>
          </tr>
        </table></td>
      </tr>
    </table>
					<script language="javascript">
					   
function changeProvince(selvalue)
{
document.gobuy.hukoucapital.length=0; 
document.gobuy.hukoucity.length=0;
var selvalue=selvalue;	  
var j,d,mm;
d=0;
for(j=0;j<provincearray.length;j++) 
	{
		if(provincearray[j][1]==selvalue) 
		{
			if (d==0)
			{
			mm=provincearray[j][2];
			}
		var newOption2=new Option(provincearray[j][0],provincearray[j][2]);
		document.all.hukoucapital.add(newOption2);
		d=d+1;	
		}		
		if(provincearray[j][1]==mm) 
		{		
			var newOption3=new Option(provincearray[j][0],provincearray[j][2]);
			document.all.hukoucity.add(newOption3);
		}			
	}
}
function changeCity(selvalue)  
{ 
	document.gobuy.hukoucity.length=0;  
	var selvalue=selvalue;
	var j;
	for(j=0;j<provincearray.length;j++) 
	{
		if(provincearray[j][1]==selvalue) 
		{
			var newOption4=new Option(provincearray[j][0],provincearray[j][2]);
			document.all.hukoucity.add(newOption4);
		}
	}
}
function selectprovince() 
{ 
	var j;
	for(j=0;j<provincearray.length;j++) 
	{
		if(provincearray[j][1]==0) 
		{
			var newOption4=new Option(provincearray[j][0],provincearray[j][2]);
			document.all.hukouprovince.add(newOption4);
		}
	}
}
selectprovince();
					</script>
					<table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td colspan="2" class="capt">确认购买信息</td>
                      </tr>
                      <tr>
					    <td height="27" bgcolor="#FFFFFF">　</td>
                        <td width="638" height="27" bgcolor="#FFFFFF">　</td>
                      </tr>
                      <tr>
                        <td width="30%" align="right" bgcolor="#FFFFFF">购买数量：</td>
                        <td bgcolor="#FFFFFF"><input name="buy_num" type="text" value="<%=num%>" size="5" maxlength="5" onblur="checknum1(this.value,<%=kucun%>,<%=huiyuanjia%>);" onkeyup="this.value=this.value.replace(/\D/g,'');pall(this.value,<%=huiyuanjia%>);"/> 库存(<%=kucun%>件)<span id="numerror"></span></td>
                      </tr>
					  
            <tr bgcolor="#ffffff">
              <td width="30%" height="30" align="right" class="table-xia">支付方式：</td>
              <td width="70%" height="30" class="table-xia"><%set rs5=connn.execute("select * FROM Iheeo_Delivery order by SongList")%>
              <%do while not rs5.eof%>
			<input name="songhuofangshi" type="radio" id="songhuofangshi" style="BORDER-RIGHT: 0px; BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px" value="<%=rs5("SongKey")%>" onclick="change_p(<%=rs5("SongFei")%>,<%=huiyuanjia%>)"/><strong><%=trim(rs5("SongName"))%> </strong><br />
			<%rs5.movenext
loop
rs5.close
set rs5=nothing%><span id="lbl_fee"></span></td>
            </tr>
                      <tr>
                        <td align="right" valign="top" bgcolor="#FFFFFF">给店主留言：</td>
                        <td bgcolor="#FFFFFF"><textarea name="tomessage" cols="80" rows="5" class="msgtosaler" id="J_msgtosaler" tabindex="10" title="选填，可以告诉卖家您对商品的特殊要求，如：颜色、尺码等" onfocus="this.value='';this.className='tips';">选填，可以告诉卖家您对商品的特殊要求，如：颜色、尺码等</textarea></td>
                      </tr>
                    </table>
					<table width="100%" border="0" cellpadding="0" cellspacing="0">
                      <tr>
                        <td colspan="2" class="capt">确认提交订单</td>
                      </tr>
                      <tr>
					    <td height="27" bgcolor="#FFFFFF">　</td>
                        <td height="27" bgcolor="#FFFFFF">　</td>
                      </tr>
                      <tr>
                        <td width="142" align="right" valign="top" bgcolor="#FFFFFF">实付款(含运费)：</td>
                        <td bgcolor="#FFFFFF"><span id="price_all"></span> 元</td>
                      </tr>
                      <tr>
                        <td align="right" valign="top" bgcolor="#FFFFFF">　</td>
                        <td bgcolor="#FFFFFF"><input alt="确认无误，购买" id="submit_b" name="submit_b" type="submit" style=" background-color:#5C99CF; color:#fff; padding:2px 0px; font-weight:bolder; font-size:14px; cursor:hand;" value="确认无误，购买"/><input name="huiyuanjia" type="hidden" value="<%=huiyuanjia%>" /></td>
                      </tr>
                    </table>
				  </form>
               </div>
			</div>
		</div>
	</div>
</div>
</body>
</html>
