<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!--#include file = inc2.asp -->
<!--#include file="func.asp"-->
<head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title><%=hometit%>-<%=company%></title>
<meta content="<%=keywords%>" name="keywords" />
<meta content="商赢快车shangwin是企帮首创的网络营销软件，是一款网站营销软件，实现了网络营销工具与系统的完美整合，商赢快车软件将企业策略分析、营销网站建设、沟通工具、网站SEO优化，网络营销推广、网络营销学习、人才进行无缝整合，让企业网络营销像聊QQ一样简单，轻松管理网站和营销推广，全面提升企业网络营销整体竞争力，是网站营销软件的最佳品牌！" name="description" />
<script src="js/chk.js" charset="gb2312" type="text/javascript"></script>
<link href="style.css" rel="stylesheet" type="text/css" />

</head>
<%u_id=session("u_id")%>
<div class="Side">
<div class="side-title">用户登录</div>
<div class="side-body">
    <div class="main_login">
      <div class="main_dlk">
         <%if u_id="" then%>
         <form name="loginfrm" method="post" action="chklogin.asp" onsubmit="return chkform();">
         <ul>
            <li>会员账号:&nbsp; <input name="login_name" type="text" size="16" maxlength="16" onfocus="this.style.background='#F4F4FF'" onblur="this.style.background='#ffffff'"/>
            </li>
            <li>会员密码:&nbsp; <input name="login_pwd" type="password" size="16" maxlength="16" onfocus="this.style.background='#F4F4FF'" onblur="this.style.background='#ffffff'"/>
            </li>
            <li>验&nbsp; 证&nbsp; 码:&nbsp;&nbsp;<input type="text" size="4" maxlength="4" name="login_code" onfocus="this.style.background='#F4F4FF'" onblur="this.style.background='#ffffff'"/> <img src="getcode.asp" onclick="this.src = this.src+'?'+Math.random();" alt="点击刷新验证码" style="cursor:pointer"> </li>
            <li class="login_bt"><input name="bt1" type="submit" value="登 录"/>&nbsp;<span><!--<a title="点击注册" href="huiyuan-reg.asp">注册</a></span>&nbsp;<span><a href="get_pwd.asp" title="点击找回密码">忘记密码?</a>--></span>
            </li>
         </ul>
         </form>
         <%else%>
         <ul class="logged">
            <li>尊敬的会员 <span class="u"><%=u_id%></span> 您好！&nbsp;欢迎登录
            </li>
            <li><span><a href="member_center.asp" title="点击进入会员中心">进入会员中心</a></span>
            </li>
            <li><span><a href="member_center.asp?action=dindan" title="点击查看订单">订单管理</a>(<%=get_count(M_memberID(u_id),0)%>)</span>
            </li>
            <li><span><a href="shop_cart.asp" title="点击查看购物车">购物车</a>(<%=get_count(M_memberID(u_id),1)%>)</span>&nbsp; <span><a href="get_exit.asp" title="点击退出">退出</a></span>
            </li>
         </ul>
         <%end if%>
      </div>
   </div>
</div>
</div>
