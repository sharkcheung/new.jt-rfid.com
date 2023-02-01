<!--#Include File="Include.asp"--><%
'==========================================
'文 件 名：GBookFrame.asp
'文件用途：留言框提交
'版权所有：深圳企帮
'==========================================

'定义页面变量
Id=Clng(Request.QueryString("Id"))
%><!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>留言Frame</title>
<script type="text/javascript" src="<%=SiteDir%>Js/Jquery.js"></script>
<script type="text/javascript" src="<%=SiteDir%>Js/Form.js"></script>
<script type="text/javascript" src="<%=SiteDir%>Js/Function.js"></script>
<style type="text/css">
#tab_4{
	font-family:Arial;
	border-left-width:1px;
	border-left-style:solid;
	border-right-width:1px;
	border-right-style:solid;
	border-bottom-width:0px;
	background-color: #CC99FF;
}
#tab_4 .inp3 {
	width: 160px;
	font-family: "宋体";
	font-size: 12px;
	border: 1px solid #C0D1D8;
	color: #9EACB9;
	margin-bottom: 5px;
}
#tab_4 td{font-size:12px;}
#tab_4 .tit1{border-width:1px 1px 1px 0;border-style:solid;border-color:#000;line-height:20px;height:20px;color:#009900; padding-right:5px; text-align:right}
#tab_4 .box1{padding:5px;color:#B6B6B6;	background-color:#FFFFFF;}
#tab_4 .inp1,#tab_4 .inp2{
	font-family:Arial;
	font-size:12px;
	overflow:auto;
	color:#9EACB9;
	border:1px solid #C0D1D8;
	margin-bottom:5px;
	width:200px;
	background-color: #E2EBF0;
}
#tab_4 .ioc1{margin-left:5px;}
body {
	margin:0px;
}
</style>
<script type="text/javascript">
String.prototype.Trim  = function(){return this.replace(/^\s+|\s+$/g,"");}
function baidu_query(){
	var obj = {
		param:{},
		construct:function(){
			var name, value, index;		
			var query = location.search.substr(1);
			var pairs = query.split("&");
			for(var i=0;i < pairs.length; i++){
				index = pairs[i].indexOf("=");
				if(index){
					name    = pairs[i].substr(0,index);
					value   = pairs[i].substr(index+1);
					this.param[name] = value;
				}
			}
			return this;
		},
		getParam: function(name, def){
			return this.param[name]==undefined?def:this.param[name];
		}
	};
	return obj.construct();
}
function valid_message()
{
	objMess = document.GBookDo.Fk_GBook_Content;
	objContact = document.GBookDo.Fk_GBook_Contact;
	objmail = document.GBookDo.Fk_GBook_Name;
	objMess.value = objMess.value.Trim();
	objContact.value = objContact.value.Trim();
	
	strMess = objMess.value;
	strContact = objContact.value;
	strMail = objmail.value;
	
	if (strMess.length > 200)
	{
		alert('您的留言过长，请减少至200字内，谢谢！');
		return false;
	}
	if (strContact.length > 50)
	{
		alert('您的联系方式过长，请减少至50字内，谢谢！');
		return false;
	}
	if (objmail.length > 50)
	{
		alert('您的昵称过长，请减少至50字内，谢谢！');
		return false;
	}
	if ((strMess.length < 1 || strMess==objMess.defaultValue) && strContact.length>0)
	{
		alert('您还没有填写留言内容，请在留言内容框内填写完提交，谢谢！');
		return false;
	}
	if ((strMail.length < 1 || strMail==objmail.defaultValue) && strMail.length>0)
	{
		alert('您还没有填写昵称，请在昵称框内填写完提交，谢谢！');
		return false;
	}
	if ((strContact.length < 1 || strContact==objContact.defaultValue) && strMess.length>0)
	{
		alert('您还没有填写联系方式，请在联系方式框内填写完提交，谢谢！');
		return false;
	}
	if ((strContact.length < 1 && strMess.length<1) ||  (strMess==objMess.defaultValue && strContact==objContact.defaultValue))
	{
		alert('您好，留言不能为空，谢谢！');
		return false;
	}
	return true;
}
function msee_init () {
	document.GBookDo.Fk_GBook_Content.value = document.GBookDo.Fk_GBook_Content.defaultValue;
	document.GBookDo.Fk_GBook_Name.value = document.GBookDo.Fk_GBook_Name.defaultValue;
	document.GBookDo.Fk_GBook_Contact.value = document.GBookDo.Fk_GBook_Contact.defaultValue;
}

var query = baidu_query();
msee_init
</script>
</head>

<body>
<form name="GBookDo" action="<%=SiteDir%>GBookDo.asp?Type=1&S=1" method="post" onsubmit="return valid_message();">
<table width="216"  border="0" cellpadding="0" cellspacing="0"  class="" id="tab_4">
    <tr>
      <td class="box1"><textarea name="Fk_GBook_Content" id="Fk_GBook_Content" rows="4" class="inp1" onFocus="if (this.value == this.defaultValue) this.value='';" onBlur="this.value=this.value.Trim(); if (this.value=='') this.value=this.defaultValue;">您好，感谢关注本网站！
如果您对我们的栏目感兴趣，请点击此处留言，谢谢！</textarea>
        <img src="<%=SiteDir%>Images/mail.gif" width="18" height="18">&nbsp;&nbsp;
        <input name="Fk_GBook_Name" id="Fk_GBook_Name" type="text" class="inp3" onFocus="if (this.value == this.defaultValue) this.value='';" onBlur="this.value=this.value.Trim(); if (this.value=='') this.value=this.defaultValue;" value="请输入你的昵称">
      <img src="<%=SiteDir%>Images/tel.gif" width="18" height="18">&nbsp;&nbsp;
        <input name="Fk_GBook_Contact" id="Fk_GBook_Contact" type="text" class="inp3" onFocus="if (this.value == this.defaultValue) this.value='';" onBlur="this.value=this.value.Trim(); if (this.value=='') this.value=this.defaultValue;" value="请输入你的联系方式"><br>
          <table width="192"  border="0" cellpadding="0" cellspacing="0">
            <tr>
              <td align="center" valign="bottom" style="color:#CCC;">企帮提供技术支持</td>
              <td align="right"><input type="hidden" name="Fk_GBook_Module" value="<%=Id%>" /><input type="hidden" name="Fk_GBook_Title" value="<%=Now()%>的留言" /><input name="imageField" type="image" src="<%=SiteDir%>Images/button.gif" border="0" /></td>
            </tr>
        </table>      </td>
    </tr>
</table>
</form>
</body>
</html>
<!--#Include File="Code.asp"-->
