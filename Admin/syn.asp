<!--#Include File="AdminCheck.asp"-->
<!--#Include File="create_remote_asp_utf8/v1.0/classes/SyncRequestHandler.asp"-->
<%
'==========================================
'文 件 名：Article.asp
'文件用途：内容管理拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'定义页面变量
Dim Fk_Article_Title,Fk_Article_Content,Fk_Article_Click,Fk_Article_Show,Fk_Article_Time,Fk_Article_Pic,Fk_Article_PicBig,Fk_Article_Template,Fk_Article_FileName,Fk_Article_Subject,Fk_Article_Recommend,Fk_Article_Keyword,Fk_Article_Description,Fk_Article_From,Fk_Article_Color,Fk_Article_Url,Fk_Article_Field,Fk_Article_onTop,Fk_Article_px
Dim Fk_Module_Name,Fk_Module_Id,Fk_Module_Menu,Fk_Module_Dir,Fk_Article_Module
Dim Temp2,KeyWordlist,kwdrs,ki
On Error Resume next

'===================================== 
'过滤字符 
'===================================== 
Function FilterText(t0) 
IF Len(t0)=0 Or IsNull(t0) Or IsArray(t0) Then FilterText="":Exit Function 
t0=Trim(t0) 
t0=Replace(t0,Chr(8),"")'回格 
t0=Replace(t0,Chr(9),"")'tab(水平制表符) 
t0=Replace(t0,Chr(10),"")'换行 
t0=Replace(t0,Chr(11),"")'tab(垂直制表符) 
t0=Replace(t0,Chr(12),"")'换页 
t0=Replace(t0,Chr(13),"")'回车 chr(13)&chr;(10) 回车和换行的组合 
t0=Replace(t0,Chr(22),"") 
t0=Replace(t0,Chr(32),"")'空格 SPACE 
t0=Replace(t0,Chr(33),"")'! 
t0=Replace(t0,Chr(34),"")'" 
t0=Replace(t0,Chr(35),"")'# 
t0=Replace(t0,Chr(36),"")'$ 
t0=Replace(t0,Chr(37),"")'% 
t0=Replace(t0,Chr(38),"")'& 
t0=Replace(t0,Chr(39),"")''
t0=Replace(t0,Chr(42),"")'* 
t0=Replace(t0,Chr(43),"")'+
t0=Replace(t0,Chr(59),"")'; 
t0=Replace(t0,Chr(60),"")'< 
t0=Replace(t0,Chr(61),"")'= 
t0=Replace(t0,Chr(62),"")'> 
t0=Replace(t0,Chr(64),"")'@ 
t0=Replace(t0,Chr(93),"")'] 
t0=Replace(t0,Chr(94),"")'^ 
t0=Replace(t0,Chr(96),"")'` 
t0=Replace(t0,Chr(123),"")'{
t0=Replace(t0,Chr(125),"")'} 
t0=Replace(t0,Chr(126),"")'~  
FilterText=t0 
End Function 

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call synConfirm() '添加内容表单
	Case 2
		Call synDo() '执行添加内容
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：synConfirm()
'作    用：信息同步确认页
'参    数：
'==========================================
Sub synConfirm()
	Id=Clng(Trim(Request.QueryString("Id")))
	'判断权限
	'If Not FkFun.CheckLimit("Module"&Fk_Module_Id) Then
		'Response.Write("无权限！")
		'Call FKDB.DB_Close()
		'Session.CodePage=936
		'Response.End()
	'End If
	Sqlstr="Select * From [Fk_Article] Where Fk_Article_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
%>
<script type="text/javascript" language="javascript">
	$(document).ready(function(){
		var host=window.location.host;
		if(GetCookie(host+"_synTradeName")==null || GetCookie(host+"_synTradeId")==null){return false;}
		if (GetCookie(host+"_synTradeName")!=""){
			addPreItem(GetCookie(host+"_synTradeId"),GetCookie(host+"_synTradeName"));
		}
	})
</script>
<style type="text/css">
#alertdiv{position:absolute; height:350px; width:550px; z-index:9999;left:50%; 
margin-left:-275px; padding:1px; border:1px #ccc solid; font-size:12px; 
display:none; background-color:#FFFFFF;overflow:scroll-y;} 
 
#alertdiv h2{ position:relative; height:23px; background-color:#E4E4E4; 
font-size:12px;padding:0; padding-left:5px; line-height:23px; margin:0; } 
 
#alertdiv h2 a{position:absolute; display:block;right:5px; top:3px;display:block; 
margin:0; width:16px; height:16px; margin:0; padding:0; overflow:hidden; 
background:url(images/close.gif) no-repeat; 
cursor:pointer; text-indent:-999px} 
 
.forminfo,.childType{padding:15px;} 
.forminfo td{padding:3px 0px;}

.input2{ width:20px; height:20px; line-height:20px;}
.childType{display:none;}
#ChooseType{cursor:pointer;}
.type-data{margin-right:10px; display:none;}
</style>
<div id="alertdiv">
<h2>行业类型选择<a href="javascript: closediv('alertdiv')" title="关闭">关闭</a></h2> 
<div class="forminfo"> 
</div>
 <div class="childType"></div>
</div>
<form id="ArticleAdd" name="ArticleAdd" method="post" action="syn.asp?Type=2" onsubmit="return false;">
<div id="BoxTop" style="width:98%;"><span>信息同步到企帮知道平台</span><a onclick="$('#Boxs').hide();$('select').show();"><img src="images/close3.gif"></a></div>
<div id="BoxContents" style="width:98%;">
<table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td height="25" align="right">标题：</td>
        <td><input name="syn_Title" type="text" class="Input" id="syn_Title" size="50"  value="<%=rs("Fk_Article_Title")%>"/>&nbsp; &nbsp;</td>
      </tr>
    <!--tr>
        <td height="25" align="right">类型：</td>
        <td><span class="type-data"></span><input type="button" class="Input" onclick="GetTypelist('<%=Request.Cookies("FkAdminName")%>');" id="ChooseType" value="点击选择行业类型"><input  value="" type="hidden" id="syn_Type"  name="syn_Type"/></td>
      </tr-->
       <tr>
        <td height="25" align="right" width="100">内容：</td>
        <td><textarea name="syn_Content" class="<%=bianjiqi%>" style="width:100%;" rows="15" id="syn_Content"><%=rs("Fk_Article_Content")%></textarea></td>
    </tr>
   </table>
</div>
<div id="BoxBottom" style="width:96%;">
        <input type="submit" onclick="Sends('ArticleAdd','syn.asp?Type=2',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" class="Button" name="button" id="button" value="同 步" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
end if 
End Sub

'==============================
'函 数 名：synDo
'作    用：执行同步内容
'参    数：
'==============================
Sub synDo()
	'判断权限
	'If Not FkFun.CheckLimit("Module"&Fk_Module_Id) Then
		'Response.Write("无权限！")
		'Call FKDB.DB_Close()
		'Session.CodePage=936
		'Response.End()
	'End If
	dim Syn_Title,Syn_Con,Syn_type
	Syn_Title	=FKFun.HTMLEncode(Trim(Request.Form("syn_Title")))
	Syn_Con		=FKFun.RemoveHTML(Trim(Request.Form("syn_Content")))
	'Syn_type	=FKFun.HTMLEncode(Trim(Request.Form("syn_Type")))
	'if Syn_type="" then
		'response.write "请选择要同步的行业类型"
		'response.end
	'end if
	dim reqHandler,key
	key = "85e5ffb11e1c4a8561b953a7e27a547c"
	set reqHandler = new SyncRequestHandler
	'初始化
	reqHandler.init()
	'设置密钥
	reqHandler.setKey(key)
	'-----------------------------
	'设置同步参数
	'-----------------------------
	reqHandler.setParameter "tit", Syn_Title		'标题
	reqHandler.setParameter "con", Syn_Con		'内容
	reqHandler.setParameter "typ", "-1"		'类型

	'请求的参数
	Dim Para,return,SyncUrl,host
	host=request.ServerVariables("HTTP_HOST")
	reqHandler.setParameter "hos", host		'域名
	Para  	= reqHandler.getParameters()
	SyncUrl	="http://qbknow.qb02.com/json/sync_article.asp"
	return	= reqHandler.PostHttpPage("qbknow.qb02.com",SyncUrl,Para)
	Response.Write(return)
	'Call FKFso.CreateFile("syn.txt","请求端签名："& Para&"--------------------------------------------------"&return)
		
	'Response.Write("新内容添加成功！"&Para)
End Sub

%><!--#Include File="../Code.asp"-->