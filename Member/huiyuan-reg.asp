<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">

<!--#include file = inc2.asp -->
<!--#include file="func.asp"-->

<head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type" />
<title>注册</title>
<meta content="商赢快车shangwin是企帮首创的网络营销软件，是一款网站营销软件，实现了网络营销工具与系统的完美整合，商赢快车软件将企业策略分析、营销网站建设、沟通工具、网站SEO优化，网络营销推广、网络营销学习、人才进行无缝整合，让企业网络营销像聊QQ一样简单，轻松管理网站和营销推广，全面提升企业网络营销整体竞争力，是网站营销软件的最佳品牌！" name="description" />
<script language="javascript" src="js/chk.js" charset="gb2312" type="text/javascript"></script>
<link href="style.css" rel="stylesheet" type="text/css" />
<title>会员注册</title>
</head>

<body oncontextmenu="return false">

<div id="main">
	<div class="m1">
		<div class="m1-body">
			<div class="member_reg">
				<form action="huiyuan-reg-save.asp" method="post" name="regform">
					<h4 class="span"><b>会员注册：</b>为了商品准时快速送达及方便联系发货，请建议你填写完整以下信息。<span class="msg">带*为必填项</span></h4>
					<ul>
						<li>账　　号：<input type="text" name="u_name" id="u_name" size="20" maxlength="16" onfocus="this.style.background='#FFFCDF';showinfos_name();" onblur="this.style.background='#FFFEEE';isName();" />
						<span class="STYLE1" id="name_re"></span>
						<span class="STYLE1" id="name_re_m">*</span></li>
						<li>姓　　名：<input name="u_name_zs" type="text" id="u_name_zs" onfocus="this.style.background='#FFFCDF';showinfosname_zs();" onblur="this.style.background='#FFFEEE';name_zs();" size="20" maxlength="10" />
						<span class="STYLE1" id="name_zs_re">*</span>&nbsp;<span class="STYLE1" id="name_zs_re_m"></span>
						</li><li>性　　别：<input type="radio" name="u_sex" id="u_sex" value="0" checked />男 
						<input type="radio" id="u_sex" name="u_sex" value="0" />女 
						<span class="STYLE1" id="sex_re">*</span>
						<span class="STYLE1" id="sex_re_m"></span></li>
						<li>密　　码：<input name="u_pass" id="u_pass" type="password" onfocus="this.style.background='#FFFCDF';showinfos_pass();" onblur="this.style.background='#FFFEEE';password();" size="20" maxlength="16" onkeyup="showStrongPic();" />&nbsp;<span class="STYLE1" id="pass_re">*</span>
						<span id="lowPic" style="display:none">
						<img src="images/bad.gif" /> 弱</span>
						<span id="midPic" style="display:none">
						<img src="images/comm.gif" /> 中</span>
						<span id="highPic" style="display:none">
						<img src="images/good.gif" /> 强</span> </li>
						<li>确认密码：<input name="u_pass_re" type="password" id="u_pass_re" onfocus="this.style.background='#FFFCDF';showinfos_pass_re();" onblur="this.style.background='#FFFEEE';pass_re();" size="20" maxlength="16" />&nbsp;<span class="STYLE1" id="pass_re_re">*</span>
						<span class="STYLE1" id="pass_re_re_m"></span>
						</li>
						<li>安全问题：<select name="u_ask" id="u_ask">
						<option value="0">我身份证最后6位数</option>
						<option value="1">我父亲的名字</option>
						<option value="2">我母亲的名字</option>
						<option value="3">我就读的小学校名</option>
						<option value="4">我最喜欢的颜色</option>
						</select> <span class="STYLE1">* 选一个熟悉的问题</span></li>
						<li>安全答案：<input name="u_answer" type="text" id="u_answer" onfocus="this.style.background='#FFFCDF';showinfos_answer();" onblur="this.style.background='#FFFEEE';answer();" size="20" maxlength="20" />
						<span class="STYLE1" id="answer_re">*</span>
						<span class="STYLE1" id="answer_re_m"></span>   
						</li>
						<li>电子邮箱：<input type="text" name="u_mail" id="u_mail" size="20" maxlength="20" onfocus="this.style.background='#FFFCDF';showinfos_email();" onblur="this.style.background='#FFFEEE';isEmail();" />
						<span class="STYLE1" id="mail_re">*</span>
						<span class="STYLE1" id="mail_re_m"> </span></li>
						
						<li>电　　话：<input type="text" name="member_tel" id="member_tel" size="20" maxlength="20" onfocus="this.style.background='#FFFCDF'" onblur="this.style.background='#FFFEEE';tel();"/>
                        <span class="STYLE1" id="tel_re"></span>
				    <span class="STYLE1" id="tel_re_m"></span></li>
						<li>手　　机：<input type="text" name="member_mobile" id="member_mobile" size="20" maxlength="20" onfocus="this.style.background='#FFFCDF'" onblur="this.style.background='#FFFEEE';mobile()" />
                        <span class="STYLE1" id="mobile_re"></span>
				    <span class="STYLE1" id="mobile_re_m"></span></li>
						<li>地　　址：<input type="text" name="member_address" id="member_address" size="60" maxlength="16" onfocus="this.style.background='#FFFCDF'" onblur="this.style.background='#FFFEEE'" /></li>
						<li  style="display:none">年　　龄：<select name="member_age" id="member_age">
						<option value="0" checked>请选择</option>
						<%for i=12 to 70
               response.Write "<option value="&i&">"&i&"</option>"
			   next
			   %></select></li>
						<li>邮　　编：<input type="text" name="member_zip" id="member_zip" size="20" maxlength="6" onkeyup="value=value.replace(/[^\d]/g,'')" onfocus="this.style.background='#FFFCDF'" onblur="this.style.background='#FFFEEE';zip();" />
                        <span class="STYLE1" id="zip_re"></span>
				    <span class="STYLE1" id="zip_re_m"></span></li>
						<li>Q Q号码： <input name="u_qq" id="u_qq" type="text" onfocus="this.style.background='#FFFCDF'" onblur="this.style.background='#FFFEEE';qq();" onbeforepaste="clipboardData.setData('text',clipboardData.getData('text').replace(/[^\d]/g,''))" onkeyup="value=value.replace(/[^\d]/g,'')" size="20" maxlength="20" />
						<span class="STYLE1" id="qq_re"></span>&nbsp;<span class="STYLE1" id="qq_re_m"></span>
						</li>
						<li style="display:none">&nbsp; 网&nbsp; 站：<input type="text" name="member_web" id="member_web" size="30" onfocus="this.style.background='#FFFCDF'" onblur="this.style.background='#FFFEEE'" />
						</li>
						<li>验&nbsp;证&nbsp;码： <input type="text" size="4" maxlength="4" name="CheckCode" id="CheckCode" onfocus="this.style.background='#FFFCDF'" onblur="this.style.background='#FFFEEE';isCheckCode();" />
						<img src="getcode.asp" onclick="this.src = this.src+'?'+Math.random();" alt="点击刷新验证码" style="cursor:pointer">
						<span class="STYLE1" id="CheckCode_re">*</span>
						<span class="STYLE1" id="CheckCode_re_m"></span></li>
						<li class="login_bt">&nbsp;
						<input type="button" onclick="tijiao()" name="Submit" value="确认注册" /></li>
					</ul>
				</form>
			</div>
		</div>
	</div>
</div>

</body>

</html>
