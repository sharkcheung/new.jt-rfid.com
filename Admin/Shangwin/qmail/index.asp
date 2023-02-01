<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"  oncontextmenu="return false">
<script>
preloadImg=new Image();
preloadImg.src="images/banner_bg.jpg";
</script>
<%
Response.Expires = 1 '不过期的
Response.Expiresabsolute = Now() + 100 
url=Request.ServerVariables("HTTP_REFERER")
domain=durl(url)
%>
<head>
<meta content="text/html; charset=utf-8" http-equiv="Content-Type">
<style type="text/css">
body{
	background:#EAF0F7;
}
#bizmail_main{
	background:url('images/banner_bg.jpg') no-repeat; width:980px; height:450px; margin:0 auto;
	margin-top:20px;
}
.bizmail_loginpanel{font-size:12px;width:320px;height:auto;border:1px solid #cccccc; margin-left:600px; margin-top:100px;
}
.bizmail_LoginBox{padding:10px 15px;background:#ffffff;}
.bizmail_loginpanel h3{padding-bottom:5px;margin:0 0 5px 0;border-bottom:1px solid #cccccc;font-size:14px;}
.bizmail_loginpanel form{margin:0;padding:0;}
.bizmail_loginpanel .bizmail_column{height:28px;}
.bizmail_loginpanel .bizmail_column label{display:block;float:left;width:30px;height:24px;line-height:24px;font-size:12px;}
.bizmail_loginpanel .bizmail_column .bizmail_inputArea{float:left;width:240px; height:24px;}
.bizmail_loginpanel input.text{font-size:14px;width:100px;height:20px;margin:0 2px;border:1px solid #C3C3C3;border-color:#7C7C7C #C3C3C3 #C3C3C3 #9A9A9A;}
.bizmail_loginpanel select{width:110px;height:20px;margin:0 2px;}
.bizmail_loginpanel .bizmail_SubmitArea{margin-left:30px;clear:both;}
.bizmail_loginpanel .bizmail_SubmitArea a{font-size:12px;margin-left:5px;}
.bizmail_loginpanel input{
	font-size:14px; height:24px;
}
</style>

</head>
<body oncontextmenu="return false">
<div id="bizmail_main">
<div id="divLoginpanelVer" class="bizmail_loginpanel">
	<div class="bizmail_LoginBox">
		<h3>登录企业邮箱</h3>
		<form action="https://exmail.qq.com/cgi-bin/login" method="post" onsubmit="if(0 == this.uin.value.length){ this.uin.focus(); return false;};if(0 == this.pwd.value.length){ this.pwd.focus(); return false;};this.submit();this.pwd.value='';return false;" target="_self">
			<input name="firstlogin" type="hidden" value="false">
			<input name="errtemplate" type="hidden" value="dm_loginpage">
			<input name="aliastype" type="hidden" value="other">
			<input name="dmtype" type="hidden" value="bizmail">
			<input name="p" type="hidden">
			<div class="bizmail_column">
				<label>帐号:</label>
				<div class="bizmail_inputArea">
					<input class="text" name="uin" value="<%=Request.Cookies("emailuser")%>">@<select name="domain">
					<option value="<%=domain%>"><%=domain%></option>
<% if domain="qebang.cn" or domain="localhost" then %>
<option value="qebang.com">qebang.com</option>
<%end if%>
					</select></div>
			</div>
			<div class="bizmail_column">
				<label>密码:</label>
				<div class="bizmail_inputArea">
	 				<input class="text" name="pwd" type="password" value="<%=Request.Cookies("emailpwd")%>"></div>
			</div>
			<div class="bizmail_SubmitArea">
				<input name="" style="WIDTH: 66px" type="submit" value="登 录" onclick=""></div>
		</form>
	</div>
</div>
</div>
</body>
<%
Function durl(url)
domext = "comnetorgcnlaccinfohkbizmemobinametvasiakrdeorg.cnco.krcom.cnnet.cngov.cn" 
arrdom = Split(domext, "")  
durl = ""
url = LCase(url)  
If url = "" Or Len(url) = 0 Then Exit Function 
url = Replace(Replace(url, "http://", ""), "https://", "")  
s1 = InStr(url, ":") - 1    
If s1 < 0 Then s1 = InStr(url, "/") - 1  
If s1 > 0 Then url = Left(url, s1)  
s2 = Split(url, ".")(UBound(Split(url, ".")))  
If InStr(domext, s2) = 0 Then     
durl = url 
Else     
For dd = 0 To UBound(arrdom) 
	If InStr(url, "." & arrdom(dd)) > 0 Then             
		durl = Replace(url, "." & arrdom(dd) & "", "")             
        If InStr(durl, ".") = 0 Then             
			durl = url             
	 	Else             
	    	durl = Split(durl, ".")(UBound(Split(durl, "."))) & "." & arrdom(dd)             
	 	End If         
	End If     
Next 
End If 
End Function
%>