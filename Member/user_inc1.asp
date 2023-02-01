<head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type">
<style type="text/css">
   .td_l{color:#132B4B;background:#F1F5FC;}
   .td_r{background:#F1F5FC;}
</style>
</head>

<%
sub dindan()
if session("u_id")="" then
response.write "<script language=javascript>window.location.href='../';</script>"
response.End
end if
member_id=M_memberID(session("u_id"))
uname=session("u_id")
%> 

<div align="right">
<table width="100%" border="0" cellpadding="4" cellspacing="0">
<tr> 
<td width="12%" style="font-size:14px;color:#FF6600;font-weight:bolder;">我的订单</td>
<td width="88%" align="right">
  <select name="zhuangtai" onChange="var jmpURL=this.options[this.selectedIndex].value ; if(jmpURL!='') {window.location=jmpURL;} else {this.selectedIndex=0 ;}" >
<option value="?action=dindan&zhuangtai=0" selected>==请选择查讯状态==</option>
<option value="?action=dindan&zhuangtai=0" >全部订单状态</option>
<option value="?action=dindan&zhuangtai=1" >未作任何处理</option>
<option value="?action=dindan&zhuangtai=2" >会员已经划出款</option>
<option value="?action=dindan&zhuangtai=3" >服务商已经收到款</option>
<option value="?action=dindan&zhuangtai=4" >服务商已经发货</option>
<option value="?action=dindan&zhuangtai=5" >会员已经收到货</option>
</select></td>
</tr>
</table>
</div>
<div align="right">
<table width="100%" border="0" align="center" cellpadding="3" cellspacing="0">
            <tr align="center"> 
              <td width="24%" style="border-top:solid #ccc 1px;border-bottom:solid #ccc 1px;"><font color="#555555">订单号</font></td>
              <td width="27%" style="border-top:solid #ccc 1px;border-bottom:solid #ccc 1px;"><font color="#555555">商品详情</font></td>
              <td width="10%" style="border-top:solid #ccc 1px;border-bottom:solid #ccc 1px;"><font color="#555555">订单金额</font></td>
              <td width="13%" style="border-top:solid #ccc 1px;border-bottom:solid #ccc 1px;"><font color="#555555">付款方式</font></td>
              <td width="14%" style="border-top:solid #ccc 1px;border-bottom:solid #ccc 1px;"><font color="#555555">订单日期</font></td>
              <td width="12%" style="border-top:solid #ccc 1px;border-bottom:solid #ccc 1px;"><font color="#555555">订单状态</font></td>
            </tr>
            <%
			page=int(request("page"))
			if request("zhuangtai")="" then
			   zhuangtai=0
			else
			zhuangtai=cint(request("zhuangtai"))
			end if
			
			  linkurl="member_center.asp?action=dindan&zhuangtai="&request("zhuangtai")&""
			  maxperpage=7
			  list_num=count_num("u_order",member_id)
			  if list_num=0 then
			     pagenum=1
			  else
				    if list_num mod maxperpage=0 then
				       pagenum=list_num\maxperpage
					else
					   pagenum=list_num\maxperpage+1
					end if
			  end if
			  if page="" or isnull(page) and int(page)<1 then
			     response.Redirect "member_center.asp?action=dindan&zhuangtai="&request("zhuangtai")&"&page=1"
				 response.end
			  end if
			  if page>pagenum then
			     response.Redirect "member_center.asp?action=dindan&zhuangtai="&request("zhuangtai")&"&page="&pagenum&""
				 response.end
			  end if
			  order_sql="select top "&maxperpage&" * from [u_order] where u_id="&member_id&""
			  if isnumeric(zhuangtai) and zhuangtai<>0 then
			     pagewhere=" and pro_paytype="&zhuangtai&""
			     order_sql=order_sql&" and pro_paytype="&zhuangtai&""
			  else
				 pagewhere=""
			  end if
			  if page>1 then
			     order_sql=order_sql&" and id not in(select top "&(page-1)*maxperpage&" id from [u_order] where u_id="&member_id&" "&pagewhere&" order by id desc)"
			  end if
			     order_sql=order_sql&" order by id desc"
  set rs=connn.execute(order_sql)
  if not rs.eof then
  iitt=0
  do while not rs.eof
bookname1=""
ab=""
  iitt=iitt+1
  if iitt mod 5=0 or iitt=1 then
     bgcolor="#F1F5FC"
     bgcolor1="#E7EEFA"
  else
     bgcolor="#FAFAFA"
     bgcolor1="#F1F5FC"
  end if
p_id=rs("p_id")
if p_id=0 then

a=split(rs("pro_ids"),"|")
for i=0 to ubound(a)-1
   b=split(a(i),",")
   for j=0 to ubound(b)-1
      ab=ab&p_price(b(j),"Fk_Product_Title")&" x "&b(1)
	  if i<>ubound(a)-1 then ab=ab&"+"
   next
next
bookname1=ab
price_h=rs("pro_price")
all_price=rs("pro_price")
titb=bookname1
p_link="#"
else
price_h=rs("pro_price")
all_price=num_p*price_h
bookname1=p_price(p_id,"Fk_Product_Title")

end if
num_p=rs("pro_num")
if len(bookname1)>140 then
   bookname=left(bookname1,140)&"..."
else
   bookname=bookname1
end if
								feiyong=rs("pro_fee_type")
								set frs=connn.execute("select SongFei,SongName from Iheeo_Delivery where SongKey="&feiyong&"")
								feiyong=frs("SongFei")
								SongName=frs("SongName")
								frs.close
								set frs=nothing
pay_statu=rs("pro_paystatu")
pro_paytype=rs("pro_paytype")%>
            <tr onmouseover="this.style.backgroundColor='#FFF6D1';" onmouseout="this.style.backgroundColor='<%=bgcolor1%>'" style="background:<%=bgcolor%>;">
              <td style="text-align:left;"><span id="lbl_id"><a href="javascript:showmenu(lay<%=iitt%>)" title="点击查看订单详情"><%=rs("order_id")%></a></span></td>
              <td style="text-align:left;"><%=bookname%></a></td>
              <td style="text-align:left;">￥<span id="lbl_Price"><%=price_h-0+feiyong%>.00</span>(含运费)</td>
              <td align="center">
			  <%select case pro_paytype
			     case 99999
				    img="银行付款"
				 case 100000
				    img="货到付款"
				 case else
                    set rs2=connn.execute("select * from Iheeo_Pay where PayKey="&pro_paytype)
		            if rs2.eof and rs2.bof then
		               img= "方式已被删除"
		            else
                       img=rs2("PayName")
                    end if
		            rs2.Close
                    set rs2=nothing
			   end select%><%=img%></td>
              <td align="center"><%=rs("order_time")%></td>
              <td align="center"><%
			  select case rs("pro_paytype")
			     case 100000
				    select case pay_statu
					   case 0
					      response.Write "未作任何处理"
					   case 2
					      response.Write "订单已确认"
					   case 4
					      response.Write "订单已完成"
					   case 5
					      response.Write "完成评价"
					   case else
					      response.Write "未作任何处理"
					end select
				 case else
				    select case pay_statu
	                  case 0
	                     response.write "未付款<br><a href=""javascript:void(0);"" onclick=""window.location.href='prolist_2.asp?pro_id="&rs("order_id")&"';return false;"" title=点出支付 target=_self><img src=images/ding_1.gif border=0></a>"
	                  case 1
	                     response.write "已付款<br><img src=images/ding_2.gif border=0>"
	                  case 2
	                     response.write "已收款<br><img src=images/ding_3.gif border=0>"
	                  case 3
	                     response.write "已发货<br><a href=""javascript:void(0);"" onclick=""window.location.href='dingdan.asp?dan="&rs("order_id")&"&oid="&rs("id")&"';return false;"" target=_self><img src=images/ding_4.gif border=0></a>"
	                  case 4
	                     response.write "已收货<br><img src=images/ding_5.gif border=0>"
	                  case 5
	                     response.write "评价完成<br><img src=images/ding_6.gif border=0>"
	                  end select
				end select
					  %></td>
            </tr>  
  <tr id='lay<%=iitt%>' style="display:none;"> 
    <td colspan="6" style="padding:0px;border-top:dashed 1px #CCC;border-bottom:dashed 1px #CCC;"><table width="100%" border="0" cellspacing="0"> 
      <tr> 
        <td width="14%" class="td_l">收&nbsp;货&nbsp; 人：</td> 
        <td width="86%" align="left" class="td_r"><%=rs("pro_contact")%></td> 
      </tr> 
      <tr> 
        <td class="td_l">收货地址：</td> 
        <td align="left" class="td_r"><%=rs("pro_add")%></td> 
      </tr> 
      <tr> 
        <td class="td_l">送货方式：</td> 
        <td align="left" class="td_r"><%=SongName%>(运费:<%=feiyong%> 元)</td> 
      </tr> 
      <tr> 
        <td class="td_l">邮 &nbsp; &nbsp; &nbsp;编：</td> 
        <td align="left" class="td_r"><%=rs("pro_post")%></td> 
      </tr> 
      <tr> 
        <td class="td_l">联系电话：</td> 
        <td align="left" class="td_r"><%=rs("pro_tel")%></td> 
      </tr> 
    </table></td> 
  </tr> 
            <%rs.movenext
			if rs.eof then exit do
			loop
			else
			response.Write "暂无订单！"
			end if
  rs.close
  set rs=nothing
  %>
  <script language=JavaScript> 
function showmenu(strID){ 
    var i; 
    for(i=1;i<=<%=iitt%>;i++){ 
        var lay; 
        lay = eval('lay' + i); 
        if (lay.style.display=="block" && lay!=eval(strID)){ 
            lay.style.display = "none"; 
        } 
    } 
    if (strID.style.display=="none"){ 
        strID.style.display = "block"; 
    }else{ 
        strID.style.display = "none"; 
    } 
} 
</script>
  </table>
  <%call showpagelist(pagenum,maxperpage,linkurl,page)%>
</div>
<br>
<%
end sub

sub myinfo()
if session("u_id")<>"" then
%>
<div align="right">
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#ffffff">
	<tr align="center"> 
	<td colspan="3"  style="border-bottom:#ccc solid 1px;" class="mem_info">&nbsp;&nbsp;会员中心</td>
	</tr>
  <%
	set bjx=server.CreateObject("adodb.recordset")
	bjx.open "select * from u_members where m_uid='"&session("u_id")&"' ",connn,1,1
	ky_jifen=bjx("jifen")
	ky_yucun=bjx("yucun")
	if ky_jifen="" then ky_jifen=0
	if ky_yucun="" then ky_yucun=0
	%>
  <tr bgcolor="#ffffff"> 
    <td width="100%" colspan="3" align="left" bgcolor="#ffffff" style="line-height:150%"><b><font color="#FF6600"><%=session("u_id")%></font></b> 欢迎您来到<strong><font color="#ff0000"><%=company%></font></strong>会员管理中心!
	<br/>
	您上次登陆的时间是：<%=bjx("m_last_logintime")%>&nbsp; &nbsp; 登录次数：<%=bjx("m_login_count")%> 次，订单(<%=get_count(member_id,0)%>) 个。</td>
  </tr>
  <tr bgcolor="#ffffff" align="left" style="display:none;"> 
    <td colspan="3" bgcolor="#ffffff">余额：<%=ky_yucun%> 元<br>
	  <!-----红包：使用红包>><br>----->
	积分：<%=ky_jifen%> 点
	</td>
    </tr>
  
    <%
	set bjx1=server.CreateObject("adodb.recordset")
	bjx1.open "select sum(zonger) as sum_jine from BJX_action where username='"&session("u_id")&"' and zhuangtai<=5",connn,1,1
	bjx1.close
	set bjx1=nothing
	if request.Cookies("bjx")("reglx")="2" then
	end if
	set bjx1=server.CreateObject("adodb.recordset")
	bjx1.open "select count(*) as rec_count from BJX_action where username='"&session("u_id")&"' and zhuangtai=6",connn,1,1
	bjx1.close
	set bjx1=nothing
	bjx.close
	set bjx=nothing%>
</table>
</div>
<%else%>
<table width="100%" height="60"  border="0" align="center" cellpadding="0" cellspacing="0" bgcolor="#cccccc">
  <tr>
    <td bgcolor="#FFFFFF"><div align="center">请登陆后操作</div></td>
  </tr>
</table>
<%end if%>
<br>
<%
end sub
sub jifen()
if session("u_id")="" then
response.Redirect "../"
response.End
end if
%>
<div align="right">
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#cccccc">
	<tr  align="center" class="mem_td mem_info">&nbsp;&nbsp;积分相关信息</td>
  </tr>
  <tr bgcolor="#FFFFFF" align="center"> 
    <td> 
      <table width="95%" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
          <td>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr> 
    <td align="center"><strong><font color=#FFFFFF>我的积分情况</font></strong>
    </td>
  </tr>
  <tr bgcolor="#FFFFFF" align="center"> 
    <td> 
        <table width="100%" height="24" border="0" cellspacing="0" cellpadding="0" align="center">
        <tr>
        <%
	set bjx=server.CreateObject("adodb.recordset")
	bjx.open "select jifen from u_members where m_uid='"&session("u_id")&"' ",connn,1,1
	ky_jifen=bjx("jifen")
	bjx.close
	set bjx=nothing%>
	<td width="50%">我的可用积分：<font color=#FF0000><%=ky_jifen%></font></td>
	</tr>
	</table>
    </td>
  </tr>
	
	  <tr> 
    <td align="center"><strong><font color=#FFFFFF>积分与预存款换算</font></strong></td>
  </tr>
  <tr bgcolor="#FFFFFF" align="center"> 
    <td> 
        <table width="100%" height="24" border="0" cellspacing="0" cellpadding="0" align="center">
            <tr>
	          <td height="30"><div align="center"><font color="#FF0000">换算：10积分 = 1元</font></div></td>
	          <%
	set bjx=server.CreateObject("adodb.recordset")
	bjx.open "select sum(zonger) as sum_jine from BJX_action where username='"&session("u_id")&"' and zhuangtai<=5",connn,1,1
	ky_jifen=bjx("sum_jine")
	bjx.close
	set bjx=nothing%>
          </tr>
            <tr>
	          <td height="30" align="center"><form name="form2" method="post" action="huansuan.asp">将<input name="jifen" type="text" id="jifen" size="10">
积分换算成预存款
<input type="submit" name="Submit" value="换算">
<input name="act" type="hidden" id="act" value="jifen">
</form></td>
          </tr>
	</table>
    </td>
  </tr>
	
  <tr align="center"> 
    <td align=center><strong><font color=#FFFFFF>奖品清单</font></strong></td>
  </tr>
  <tr bgcolor="#FFFFFF" align="center"> 
    <td> 
      <table width="95%" border="0">
        <tr align="center"> 
          <td height="24">奖品名称</td>
          <td height="24">需要积分</td>
          <td height="24">操作</td>
        </tr>
        <%
	set bjx=server.CreateObject("adodb.recordset")
	bjx.open "select * from BJX_jiangpin where xianshi=1",connn,1,1
	while not bjx.eof%>
        <tr>
          <td height="24"><div align="center"><a href="iheeo_jp.asp?id=<%=bjx("bookid")%>" ><%=bjx("bookname")%></a></div></td>
          <td align="center" height="24"><%=bjx("jifen")%></td>
          <td align="center" height="24"><a href="jifen.asp?id=<%=bjx("bookid")%>&action=add">选择此项</a></td>
        </tr>
        <%
	bjx.movenext
	wend
	bjx.close
	set bjx=nothing%>
      </table>
    </td>
  </tr>
  <%
	set bjx=server.CreateObject("adodb.recordset")
	bjx.open "select * from BJX_action_jp where username='"&session("u_id")&"' and zhuangtai=7",connn,1,1
	if bjx.recordcount>0 then%>
  <tr align="center"> 
    <td><strong><font color=#FFFFFF>您已选择的奖品清单</font></strong></td>
  </tr>
  <tr bgcolor="#FFFFFF" align="center"> 
    <td> 
      <table width="95%" border="0">
        <tr align="center"> 
          <td height="25">奖品名称</td>
          <td height="25">使用积分</td>
          <td height="25">操作</td>
        </tr>
        <%
	while not bjx.eof%>
        <tr>
          <td height="23">
            <div align="center">
              <%
	set bjx1=server.CreateObject("adodb.recordset")
	bjx1.open "select * from BJX_jiangpin where bookid="&bjx("bookid"),connn,1,1
	if bjx1.recordcount=1 then
	response.write "<a href='iheeo_jp?id="&bjx("bookid")&"' >"&bjx1("bookname")&"</a>"
	end if
	bjx1.close
	set bjx1=nothing%>
            </div></td>
          <td align="center" height="23"><%=bjx("jifen")%></td>
          <td align="center" height="23"><a href="jifen.asp?actionid=<%=bjx("actionid")%>&action=del">删除此项</a></td>
        </tr>
        <%
	bjx.movenext
	wend%>
      </table>
    </td>
  </tr>
  <%end if
	bjx.close
	set bjx=nothing%>
</table></div>
<br>
<%
end sub
sub shoucang()
if session("u_id")="" then
response.Redirect "../"
response.End
end if
%>
<script language="JavaScript">
<!--
var newWindow = null
function windowOpener(loadpos)
{	
  newWindow = window.open(loadpos,'newwindow','width=450,height=350,toolbar=no, status=no, menubar=no, resizable=yes, scrollbars=yes');
	newWindow.focus();
}

//-->
</script>
<div align="right">
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#cccccc">
<tr><td align=right><strong><font color=#FFFFFF>您最多只能收藏十种商品</font></strong></td></tr></table>
</div>
<%
set rs=server.CreateObject("adodb.recordset")
rs.open "select bjx_action.actionid,bjx_action.bookid,BJX_goods.bookname,BJX_goods.shichangjia,BJX_goods.huiyuanjia,BJX_goods.vipjia,BJX_goods.dazhe from BJX_goods inner join  bjx_action on BJX_goods.bookid=bjx_action.bookid where bjx_action.username='"&session("u_id")&"' and bjx_action.zhuangtai=6",connn,1,1 
%>
<div align="right">
<table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#cccccc">
  <form action="sctogw.asp" target="newwindow" method=post name=form1 onsubmit="windowOpener('')">
    <tr bgcolor="#FFFFFF" align="center"> 
      <td width=8%>选择</td>
      <td width=42%>商品名称</td>
      <td width=14%>市场价</td>
      <td width=14%>会员价</td>
      <td width=14%>VIP 价</td>
      <td width=8%>删除</td>
    </tr>
    <%do while not rs.eof%>
    <tr bgcolor="#ffffff" align=center> 
      <td><input name=bookid type=checkbox checked value="<%=rs("bookid")%>" ></td>
      <td align=left><a href=product.asp?Iheeoid=<%=rs("bookid")%> ><%=rs("bookname")%></a></td>
      <td><s><%=rs("shichangjia")%></s>元</td>
      <td><%=rs("huiyuanjia")%>元</td>
      <td><%=rs("vipjia")%>元</td>
      <td><a href=shoucang.asp?action=del&actionid=<%=rs("actionid")%>&ll=1><img src=../images/trash.gif width=15 height=17 border=0></a>
      </td>
    </tr>
    <%
rs.movenext
loop
rs.close
set rs=nothing
%>
	<tr bgcolor="#ffffff" align=center>
	<td height=25 colspan=6> 
	<input class="go-wenbenkuang" onFocus="this.blur()" type=submit name="submit" value=" 加入购物车 ">
	</td>
	</tr>
  </form>
</table></div>
<br>
<%
end sub

sub savepass()
if session("u_id")="" then
response.Redirect "../"
response.End
end if
%>
<script language=JavaScript>
function checkspace(checkstr) {
  var str = '';
  for(i = 0; i < checkstr.length; i++) {
    str = str + ' ';
  }
  return (str == checkstr);
}
function passcheck()
{
    if(document.userpass.userpassword.value.length < 6 || document.userpass.userpassword.value.length >20) {
	document.userpass.userpassword.focus();
    alert("密码长度不能不能这空，在6位到20位之间，请重新输入！");
	return false;
  }
   if(document.userpass.userpassword.value!==document.userpass.userpassword2.value) {
	document.userpass.userpassword.focus();
    alert("对不起，两次密码输入不一样！");
	return false;
  }
}
</script>
<%
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from u_members where m_uid='"&session("u_id")&"' ",connn,1,1
%>
<div align="right">
<table width="100%" border="0" cellpadding="5" cellspacing="0">
  <form name="userpass" method="post" action="saveuserinfo.asp?action=savepass">
    <tr> 
      <td colspan=2 align="left" class="mem_td mem_info">&nbsp;&nbsp;我的密码管理</td>
    </tr>
    <tr> 
      <td width="30%" bgcolor=#ffffff align="right">用 户 名：</td>
      <td width="70%" align="left" bgcolor=#ffffff><font color=#FF6600><%=session("u_id")%></font></td>
	</tr>
    <tr> 
      <td bgcolor=#ffffff align="right">新 密 码：</td>
      <td align="left" bgcolor=#ffffff><input name="userpassword" class="wenbenkuang"; type="password" value="" size="18">
	  <font color="#FF0000">**</font> 不修改请为空</td>
    </tr>
    <tr> 
      <td bgcolor=#ffffff align="right">密码确认：</td>
      <td align="left" bgcolor=#ffffff><input name="userpassword2" class="wenbenkuang" type="password" value="" size="18">
	  <font color="#FF0000">**</font></td>
    </tr>
    <tr align="center"> 
      <td height=25 bgcolor=#ffffff colspan="2">
	  <input class="go-wenbenkuang" onclick="return passcheck();" type="submit" name="submit" value=" 提交保存 ">
	  <input class="go-wenbenkuang" onclick="ClearReset()" type=reset name="Clear" value=" 重新填写 ">
      </td>
    </tr>
  </form>
</table></div>
<br>
<%rs.close
set rs=nothing
end sub

sub userziliao()
if session("u_id")="" then
response.Redirect "../"
response.End
end if
%>
<script language=JavaScript>
	function chsel(){
		with (document.userinfo){
			if(szSheng.value) {
				szShi.options.length=0;
				for(var i=0;i<selects[szSheng.value].length;i++){
					szShi.add(selects[szSheng.value][i]);
				}
			}
		}
	}
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
function checkuserinfo()
{
 if(document.userinfo.useremail.value.length!=0)
  {
    if (document.userinfo.useremail.value.charAt(0)=="." ||        
         document.userinfo.useremail.value.charAt(0)=="@"||       
         document.userinfo.useremail.value.indexOf('@', 0) == -1 || 
         document.userinfo.useremail.value.indexOf('.', 0) == -1 || 
         document.userinfo.useremail.value.lastIndexOf("@")==document.userinfo.useremail.value.length-1 || 
         document.userinfo.useremail.value.lastIndexOf(".")==document.userinfo.useremail.value.length-1)
     {
      alert("Email地址格式不正确！");
      document.userinfo.useremail.focus();
      return false;
      }
   }
 else
  {
   alert("Email不能为空！");
   document.userinfo.useremail.focus();
   return false;
   }
   if(checkspace(document.userinfo.userzhenshiname.value)) {
	document.userinfo.userzhenshiname.focus();
    alert("对不起，请填写您的真实姓名！");
	return false;
  }
  /*
   if(checkspace(document.userinfo.sfz.value)) {
	document.userinfo.sfz.focus();
    alert("对不起，请填写您的身份证号码！");
	return false;
  }
  if((document.userinfo.sfz.value.length!=15)&&(document.userinfo.sfz.value.length!=18)) {
	document.userinfo.sfz.focus();
    alert("对不起，请正确填写身份证号码！");
	return false;
  } */
  if(checkspace(document.userinfo.shouhuodizhi.value)) {
	document.userinfo.shouhuodizhi.focus();
    alert("对不起，请填写您的详细地址！");
	return false;
  }
  if(checkspace(document.userinfo.youbian.value)) {
	document.userinfo.youbian.focus();
    alert("对不起，请填写邮编！");
	return false;
  }
  if(document.userinfo.youbian.value.length!=6) {
	document.userinfo.youbian.focus();
    alert("对不起，请正确填写邮编！");
	return false;
  } 
    if(checkspace(document.userinfo.usertel.value) && checkspace(document.userinfo.usermobile.value)) {
	document.userinfo.usertel.focus();
    alert("对不起，请留下您的手机或联系电话！");
	return false;
  }
  
}
</script>
<%
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from u_members where m_uid='"&session("u_id")&"' ",connn,1,1
resume_hukouprovinceid=rs("szSheng")
resume_hukoucapitalid=rs("szShi")
resume_hukoucityid=rs("szXian")
m_question=rs("m_question")
select case m_question
   case 0
      m_question="我身份证最后6位数"
   case 1
      m_question="我父亲的名字"
   case 2
      m_question="我母亲的名字"
   case 3
      m_question="我就读的小学校名"
   case 4
      m_question="我最喜欢的颜色"
end select
%>
<div align="right">
<table width="100%" border="0" cellpadding="5" cellspacing="0">
  <form name="userinfo" method="post" action="saveuserinfo.asp?action=userziliao">
    <tr> 
      <td colspan=4 align="left" class="mem_td mem_info">&nbsp;&nbsp;我的会员信息</strong>
		</td>
    </tr>
    <tr bgcolor="#ffffff"> 
      <td align="right" class="mem_td">用 户 名：</td>
      <td align="left" class="mem_td"><font color=#FF6600><b>
	  <%=session("u_id")%></b></font></td>
      <td align="right" class="mem_td">真实姓名：</td>
      <td align="left" class="mem_td"> 
        <input name=userzhenshiname class="wenbenkuang" type=text value="<%=trim(rs("m_uname"))%>" size="20">
		<font color="#FF0000">**</font></td>
    </tr>
    <tr bgcolor="#ffffff"> 
      <td align="right" class="mem_td">安全问题：</td>
      <td align="left" class="mem_td"><%=m_question%></td>
      <td align="right" class="mem_td">安全密码：</td>
      <td align="left" class="mem_td"> 
        <input name=m_answer class="wenbenkuang" type=text value="<%=trim(rs("m_answer"))%>" size="20">
		<font color="#FF0000">**</font></td>
    </tr>
    <tr bgcolor="#ffffff"> 
      <td align="right" class="mem_td">电子邮件：</td>
      <td align="left" class="mem_td"> 
        <input name=useremail class="wenbenkuang" type=text value="<%=trim(rs("m_uemail"))%>">
		<font color="#FF0000">**</font></td>
      <td align="right" class="mem_td"> 
        性别：</td>
      <td align="left" class="mem_td"> 
        <input type="radio" style="border:none;" name="shousex" value="0" <%if rs("m_usex")=0 then%>checked<%end if%> checked> 
		男<input type="radio" style="border:none;" name="shousex" value="1" <%if rs("m_usex")=1 then%>checked<%end if%>>
        女　　年龄：<input name=nianling type=text class="wenbenkuang" value="<%if rs("m_uage")<>"" and rs("m_uage")<>0 then response.Write rs("m_uage")%>" size="4" maxlength="2" ONKEYPRESS="event.returnValue=IsDigit();"></td>
    </tr>
    <tr bgcolor="#ffffff"> 
      <td align="right" class="mem_td">电话号码</td>
      <td align="left" class="mem_td"> 
        <input name="usertel" class="wenbenkuang" type=text value="<%=trim(rs("m_utel"))%>"></td>
      <td align="right" class="mem_td"> 
        手机号码：</td>
      <td align="left" class="mem_td"> 
        <input name="usermobile" type=text class="wenbenkuang" value="<%=rs("m_umobile")%>" size="20">
		<font color="#FF0000">**</font></td>
    </tr>
    <tr bgcolor="#ffffff" style="display:none"> 
      <td align="right" class="mem_td">身份证号码：</td>
      <td align="left" class="mem_td" colspan="3"> 
        <input name=sfz type=text class="wenbenkuang" value="<%=trim(rs("m_sfz"))%>" size="30" maxlength="18">
		<font color="#FF0000">**</font></td>
    </tr>
    <tr bgcolor="#ffffff"> 
      <td align="right" class="mem_td">所在城市：</td>
      <td align="left" class="mem_td" colspan="3"> 
        <select name="hukouprovince" size="1" id="select5" onChange="changeProvince(document.userinfo.hukouprovince.options[document.userinfo.hukouprovince.selectedIndex].value)">
		<%if resume_hukouprovinceid<>"" then%>
		<option value="<%=resume_hukouprovinceid%>"><%=Hireworkadds(resume_hukouprovinceid)%></option>
		<%else%>
		<option value="">选择省</option>
		<%end if%>
		</select>
						<select name="hukoucapital" onchange="changeCity(document.userinfo.hukoucapital.options[document.userinfo.hukoucapital.selectedIndex].value)">
						  <%if resume_hukoucapitalid<>"" then%>
						  <option value="<%=resume_hukoucapitalid%>"><%=Hireworkadds(resume_hukoucapitalid)%></option>
						  <%else%>
						  <option value="">选择市</option>
						  <%end if%>
	                      </select>
		                  <select name="hukoucity">
		                    <%if resume_hukoucityid<>"" then%>
		                    <option value="<%=resume_hukoucityid%>"><%=Hireworkadds(resume_hukoucityid)%></option>
		                    <%else%>
		                    <option value="">选择区</option>
		                    <%end if%>
                          </select>
		<font color="#FF0000">**</font></td>
    </tr>
    <tr bgcolor="#ffffff"> 
      <td align="right" class="mem_td">详细地址：</td>
      <td align="left" class="mem_td" colspan="3"> 
        <input name=shouhuodizhi type=text class="wenbenkuang" value="<%=trim(rs("m_uaddress"))%>" size="60">
		<font color="#FF0000">**</font></td>
    </tr>
    <tr bgcolor="#ffffff"> 
      <td align="right" class="mem_td">邮编： </td>
      <td align="left" class="mem_td"> 
        <input name=youbian type=text class="wenbenkuang" value="<%=trim(rs("m_uzip"))%>" ONKEYPRESS="event.returnValue=IsDigit();" size="20">
		<font color="#FF0000">**</font></td>
      <td align="right" class="mem_td"> 
        QQ：</td>
      <td align="left" class="mem_td"> 
        <input name=QQ type=text class="wenbenkuang" value="<%=trim(rs("m_uQQ"))%>" size="20" maxlength="12"></td>
    </tr>
    <tr bgcolor="#ffffff"> 
      <td align="right">自我介绍：</td>
      <td align="left" colspan="3"> 
        <textarea name="content" cols="30" rows="5" class="wenbenkuang"><%=trim(rs("content"))%></textarea>
      </td>
    </tr>
    <tr align="center"> 
      <td height=25 bgcolor=#ffffff colspan="4">
	  <input class="go-wenbenkuang" onclick="return checkuserinfo();" type="submit" name="submit" value=" 提交保存 ">
	  <input class="go-wenbenkuang" onclick="ClearReset()" type=reset name="Clear" value=" 重新填写 ">
      </td>
    </tr>
  </form>
</table><script language = "JavaScript" charset="gb2312" src="js/GetProvince.js"></script>
					<script language="javascript">
					   
function changeProvince(selvalue)
{
document.userinfo.hukoucapital.length=0; 
document.userinfo.hukoucity.length=0;
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
	document.userinfo.hukoucity.length=0;  
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
					</script></div>
<br>
<%rs.close
set rs=nothing
end sub
%>