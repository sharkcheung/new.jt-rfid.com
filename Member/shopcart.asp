<!--#include file = inc2.asp -->
<!--#include file="func.asp"-->
<%
u_id=session("u_id")
member_id=M_memberID(u_id)
if u_id="" then
%>
<a href="" onclick="openaddcat('/member/',300,250,'欢迎登陆！');return false;">请先登录</a>，<a href=""  onclick="openaddcat('/member/huiyuan-reg.asp',690,450,'欢迎注册会员！');return false;">还未注册会员？</a>
<%
else%>
<%
	set bjx=server.CreateObject("adodb.recordset")
	bjx.open "select * from u_members where m_uid='"&session("u_id")&"' ",connn,1,1
	ky_jifen=bjx("jifen")
	ky_yucun=bjx("yucun")
	if ky_jifen="" then ky_jifen=0
	if ky_yucun="" then ky_yucun=0
%>
欢迎：<a  onclick="openaddcat('/member/member_center.asp',890,450,'会员中心');return false;" href=""><%=session("u_id")%></a>，
<%
	bjx.close
	set bjx=nothing
	%>
<a onclick="openaddcat('/member/shop_cart.asp',890,450,'购物车');return false;" href="">购物车</a>(<%=get_count(member_id,1)%>)，<a onclick="openaddcat('/member/member_center.asp?action=dindan',890,450,'我的订单');return false;"  href="">订单(<%=get_count(member_id,0)%>)</a>，
<a href="/member/get_exit.asp">退出</a> 
<%end if%>
  