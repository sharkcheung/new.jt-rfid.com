<!--#Include File="AdminCheck.asp"-->
<!--#Include File="../Class/Cls_HTML.asp"-->
<%
'==========================================
'文 件 名：Article.asp
'文件用途：内容管理拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'定义页面变量
Dim Fk_Article_Title,Fk_Article_Content,Fk_Article_Click,Fk_Article_Show,Fk_Article_Time,Fk_Article_Pic,Fk_Article_PicBig,Fk_Article_Template,Fk_Article_FileName,Fk_Article_Subject,Fk_Article_Recommend,Fk_Article_Keyword,Fk_Article_Description,Fk_Article_From,Fk_Article_Color,Fk_Article_Url,Fk_Article_Field,Fk_Article_onTop,Fk_Article_px,Fk_Article_Seotitle
Dim Fk_Module_Name,Fk_Module_Id,Fk_Module_Menu,Fk_Module_Dir,Fk_Article_Module
Dim Temp2,KeyWordlist,kwdrs,ki,host
dim appendFrom
dim Fk_Article_Copyright,Fk_Article_CopyrightInfo,Fk_Article_CopyrightFs,Fk_Article_CopyrightFt,Fk_Article_CopyrightCl
host=request.ServerVariables("HTTP_HOST")

'On Error Resume next
	dim KeyWorddat,krs
	Sqlstr="select SVkeywords from [keywordSV]"
	set krs=conn.execute(Sqlstr)
	If Not krs.Eof Then
		dim m
		m=0
		KeyWorddat=""
		do while Not krs.Eof
			if m=0 then
				KeyWorddat=FilterText(krs("SVkeywords"))
			else
				KeyWorddat=KeyWorddat&"|"&FilterText(krs("SVkeywords"))
			end if
			m=m+1
			krs.movenext
			if krs.eof then exit do
		loop
	else
		KeyWorddat=""
	end if
	krs.close
if FKFso.IsFile("KeyWordC.dat") then
	KeyWordlist=FKFso.FsoFileRead("KeyWordC.dat")
else
	KeyWordlist=KeyWorddat
	call FKFso.CreateFile("KeyWordC.dat",KeyWorddat)
end if


Function CheckFields(FieldsName,TableName)
	dim blnFlag,chkStrSql,chkStrRs
	blnFlag=False
	chkStrSql="select * from "&TableName
	Set chkStrRs=Conn.Execute(chkStrSql)
	for i = 0 to chkStrRs.Fields.Count - 1
		if lcase(chkStrRs.Fields(i).Name)=lcase(FieldsName) then
			blnFlag=True
			Exit For
		else
			blnFlag=False
		end if
	Next
	CheckFields=blnFlag
End Function

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
		Call ArticleList() '内容列表
	Case 2
		Call ArticleAddForm() '添加内容表单
	Case 3
		Call ArticleAddDo() '执行添加内容
	Case 4
		Call ArticleEditForm() '修改内容表单
	Case 5
		Call ArticleEditDo() '执行修改内容
	Case 6
		Call ArticleDelDo() '执行删除内容
	Case 7
		Call ListDelDo() '执行批量删除内容
	Case 8
		Call ArticleMove() '执行批量移动内容
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：ArticleList()
'作    用：内容列表
'参    数：
'==========================================
Sub ArticleList()
	'新功能，追加SEO title字段
	'2017年5月22日
	'middy241@163.com
	if CheckFields("Fk_Article_seotitle","Fk_Article")=false then
		conn.execute("alter table Fk_Article add column Fk_Article_seotitle varchar(255) null")
	end if
	Session("NowPage")=FkFun.GetNowUrl()
	Dim SearchStr
	SearchStr=FkFun.HTMLEncode(Trim(Request.QueryString("SearchStr")))
	Fk_Module_Id=Clng(Trim(Request.QueryString("ModuleId")))
	'判断权限
	'If Not FkFun.CheckLimit("Module"&Fk_Module_Id) Then
		'Response.Write("无权限！")
		'Call FKDB.DB_Close()
		'Session.CodePage=936
		'Response.End()
	'End If
	PageNow=Trim(Request.QueryString("Page"))
	If PageNow="" Then
		PageNow=1
	Else
		PageNow=Clng(PageNow)
	End If
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id=" & Fk_Module_Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Module_Name=Rs("Fk_Module_Name")
		Fk_Module_Menu=Rs("Fk_Module_Menu")
		Fk_Module_Dir=Rs("Fk_Module_Dir")
	Else
		PageErr=1
	End If
	Rs.Close
%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="ShowBox('Article.asp?Type=2&ModuleId=<%=Fk_Module_Id%>');">添加</a></li>
        <li><a href="javascript:void(0);" onclick="SetRContent('MainRight','<%=Session("NowPage")%>');return false">刷新</a></li>
    </ul>
</div>
<div id="ListTop">
    <%=Fk_Module_Name%>栏目&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name="SearchStr" value="<%=SearchStr%>" type="text" class="Input" id="SearchStr" style="vertical-align:middle;"/>&nbsp;<input type="button" class="Button" onclick="SetRContent('MainRight','Article.asp?Type=1&ModuleId=<%=Fk_Module_Id%>&SearchStr='+escape(document.all.SearchStr.value));" name="S" Id="S" value="  查询  "  style="vertical-align:middle;"/>&nbsp;&nbsp;请选择栏目：
<select name="D1" id="D1" onChange="window.execScript(this.options[this.selectedIndex].value);" style="vertical-align:middle;">
      <option value="alert('请选择栏目');">请选择栏目</option>
<%
Call ModuleSelectUrl(Fk_Module_Menu,0,Fk_Module_Id)
%>
</select>
</div>
<div id="ListContent">
    <form name="DelList" id="DelList" method="post" action="Article.asp?Type=7" onsubmit="return false;">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">选</td>
            <td align="left" class="ListTdTop">标题</td>
            <td align="center" class="ListTdTop" style="display:none">文件名</td>
            <td align="left" class="ListTdTop">显示</td>
            <td align="center" class="ListTdTop">点击量</td>
            <td align="center" class="ListTdTop">排序</td>
            <td align="center" class="ListTdTop">转发微博</td>
            <td align="center" class="ListTdTop">添加时间</td>
            <td align="left" class="ListTdTop">操作</td>
        </tr>
<%
	Dim Rs2,ArticleUrl,zfurl
	Set Rs2=Server.Createobject("Adodb.RecordSet")
	Sqlstr="Select * From [Fk_Article] Where Fk_Article_Module="&Fk_Module_Id&""
	If SearchStr<>"" Then
		Sqlstr=Sqlstr&" And Fk_Article_Title Like '%%"&SearchStr&"%%'"
	End If
	Sqlstr=Sqlstr&" Order By Fk_Article_Ip desc,Px desc,Fk_Article_Time Desc"
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Dim ArticleTemplate
		Rs.PageSize=PageSizes
		If PageNow>Rs.PageCount Or PageNow<=0 Then
			PageNow=1
		End If
		PageCounts=Rs.PageCount
		Rs.AbsolutePage=PageNow
		PageAll=Rs.RecordCount
		i=1
		While (Not Rs.Eof) And i<PageSizes+1
			If Rs("Fk_Article_Template")>0 Then
				Sqlstr="Select * From [Fk_Template] Where Fk_Template_Id=" & Rs("Fk_Article_Template")
				Rs2.Open Sqlstr,Conn,1,1
				If Not Rs2.Eof Then
					Fk_Article_Template=Rs2("Fk_Template_Name")
				Else
					Fk_Article_Template="未知模板"
				End If
				Rs2.Close
			Else
				Fk_Article_Template="默认模板"
			End If
			Fk_Article_Recommend=""
			If Rs("Fk_Article_Recommend")<>"" Then
				TempArr=Split(Rs("Fk_Article_Recommend"),",")
				For Each Temp In TempArr
					If Temp<>"" Then
						Sqlstr="Select * From [Fk_Recommend] Where Fk_Recommend_Id=" & Temp
						Rs2.Open Sqlstr,Conn,1,1
						If Not Rs2.Eof Then
							Fk_Article_Recommend=Fk_Article_Recommend&","&Rs2("Fk_Recommend_Name")
						End If
						Rs2.Close
					End If
				Next
			End If
			Fk_Article_Subject=""
			If Rs("Fk_Article_Subject")<>"" Then
				TempArr=Split(Rs("Fk_Article_Subject"),",")
				For Each Temp In TempArr
					If Temp<>"" Then
						Sqlstr="Select * From [Fk_Subject] Where Fk_Subject_Id=" & Temp
						Rs2.Open Sqlstr,Conn,1,1
						If Not Rs2.Eof Then
							Fk_Article_Subject=Fk_Article_Subject&","&Rs2("Fk_Subject_Name")
						End If
						Rs2.Close
					End If
				Next
			End If
			
			
			zfurl="http://"&Request.ServerVariables("Server_name")
			If Rs("Fk_Article_Url")<>"" Then
				ArticleUrl=Rs("Fk_Article_Url")
				if instr(ArticleUrl,"http://") then
					zfurl=ArticleUrl
				else
					zfurl=zfurl&ArticleUrl
				end if
			Else
				If Fk_Module_Dir<>"" Then
					ArticleUrl=Fk_Module_Dir&"/"
				Else
					ArticleUrl="Article"&Fk_Module_Id&"/"
				End If
				If Rs("Fk_Article_FileName")<>"" Then
					ArticleUrl=ArticleUrl&Rs("Fk_Article_FileName")&".html"
				Else
					ArticleUrl=ArticleUrl&Rs("Fk_Article_Id")&".html"
				End If
				If SiteHtml=1 and sitetemplate<>"wap" Then
					ArticleUrl="/html"&SiteDir&ArticleUrl
				Else
					ArticleUrl=SiteDir&sTemp&"?"&ArticleUrl
				End If
				zfurl=zfurl&ArticleUrl
			End If
%>
        <tr>
            <td height="20" align="center"><input type="checkbox" name="ListId" class="Checks" value="<%=Rs("Fk_Article_Id")%>" id="List<%=Rs("Fk_Article_Id")%>" /></td>
            <td align="left" class="td1">&nbsp;&nbsp;<%=Rs("Fk_Article_Title")%><%If Rs("Fk_Article_Color")<>"" Then%><span style="color:<%=Rs("Fk_Article_Color")%>">■</span><%End If%><%If Rs("Fk_Article_Url")<>"" Then%>[转向链接]<%End If%></td>
            <td align="center"  style="display:none"><%=Rs("Fk_Article_FileName")%></td>
            <td align="left"><%If Rs("Fk_Article_Show")=1 Then%><img src="images/caidan1.png" style="vertical-align:middle;"/><%Else%><img src="images/caidan0.png" style="vertical-align:middle;"/><%End If%><%If Rs("Fk_Article_Pic")<>"" Then%><span style="color:#F00">[图]</span><%End If%><a style="display:none;" href="javascript:void(0);" title="<%=Fk_Article_Template%> ">[模]</a><%If InStr(Fk_Article_Recommend,"推荐")>0 Then%><a href="javascript:void(0);" title="<%=Replace(Fk_Article_Recommend,",","")%> ">[推]</a><%End If%><%If trim(Rs("Fk_Article_Ip")&" ")="1" Then%><a href="javascript:void(0);" title="置顶 "><font color=blue style="font-weight:bolder;">[顶]</font></a><%End If%><%If Fk_Article_Subject<>"" Then%><a href="javascript:void(0);" title="<%=Fk_Article_Subject%> ">[专]</a><%End If%></td>
            <td align="center"><%=Rs("Fk_Article_Click")%></td>
            <td height="20" align="center"><%=Rs("Px")%></td>
            <td align="center">
			<a href="http://share.v.t.qq.com/index.php?c=share&a=index&site=<%=server.URLEncode(zfurl)%>&title=<%=server.URLEncode(Rs("Fk_Article_Title")&"("&Rs("Fk_Article_Keyword")&")")%>&pic=<%if left(Rs("Fk_Article_Pic"),4)<>"http" then
	response.write "http://"&Request.ServerVariables("Server_name")&Rs("Fk_Article_Pic")
else
	response.write Rs("Fk_Article_Pic")
end if
			%>" target="_blank"><img style="cursor:pointer;vertical-align:middle;" alt="转发到腾讯微博 " src="Images/weiboicon16.png" ></a>&nbsp;<a href="http://service.weibo.com/share/share.php?url=<%=zfurl%>&appkey=1525536596&title=<%=server.URLEncode(Rs("Fk_Article_Title")&"("&Rs("Fk_Article_Keyword")&")")%>&pic=<%if left(Rs("Fk_Article_Pic"),4)<>"http" then
	response.write "http://"&Request.ServerVariables("Server_name")&Rs("Fk_Article_Pic")
else
	response.write Rs("Fk_Article_Pic")
end if
			%>&ralateUid=" target="_blank"><img style="cursor:pointer;vertical-align:middle;" alt="转发到新浪微博 " src="Images/weiboicon16-sina.png" ></a></td>
            <td align="center"><%=Rs("Fk_Article_Time")%></td>
            <td align="left"><a title="同步到企帮知道平台 " href="javascript:void(0);" onclick="ShowBox('syn.asp?Type=1&Id=<%=Rs("Fk_Article_Id")%>');"><img src="http://image001.dgcloud01.qebang.cn/website/com.gif" style="vertical-align:middle;"></a> <a title="修改 " href="javascript:void(0);" onclick="ShowBox('Article.asp?Type=4&Id=<%=Rs("Fk_Article_Id")%>&ModuleId=<%=Fk_Module_Id%>');"><img src="images/edit.png" style="vertical-align:middle;"></a> <a title="删除 " href="javascript:void(0);" onclick="DelIt('您确认要删除“<%=Rs("Fk_Article_Title")%>”，此操作不可逆！','Article.asp?Type=6&Id=<%=Rs("Fk_Article_Id")%>','MainRight','<%=Session("NowPage")%>');"><img src="images/del.png" style="vertical-align:middle;"></a></td>
        </tr>
<%
			Rs.MoveNext
			i=i+1
		Wend
%>
        <tr>
            <td height="30" colspan="8">
            <input name="chkall" type="checkbox" id="chkall" value="select" onClick="CheckAll(this.form)" style="vertical-align:middle;"><label for="chkall" style="vertical-align:middle;">全选</label>
            <input type="submit" value="删 除" class="Button" onClick="if(confirm('此操作无法恢复！！！请慎重！！！\n\n确定要删除选中的内容吗？')){Sends('DelList','Article.asp?Type=7',0,'',0,1,'MainRight','<%=Session("NowPage")%>');}" style="vertical-align:middle;">
<select name="ArticleMove" id="ArticleMove" onchange="DelIt('确实要移动这部分内容？','Article.asp?Type=8&Id='+this.options[this.options.selectedIndex].value+'&ListId='+GetCheckbox(),'MainRight','<%=Session("NowPage")%>');" style="vertical-align:middle;">
      <option value="">转移到</option>
<%
Call ModuleSelectId(Fk_Module_Menu,0,0)
%>
</select>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%Call FKFun.ShowPageCode("Article.asp?Type=1&ModuleId="&Fk_Module_Id&"&SearchStr="&Server.URLEncode(SearchStr)&"&Page=",PageNow,PageAll,PageSizes,PageCounts)%></td>
        </tr>
<%
	Else
%>
        <tr>
            <td height="25" colspan="8" align="center">暂无记录</td>
        </tr>
<%
	End If
	Rs.Close
%>
    </table>
    </form>
</div>
<div id="ListBottom">

</div>
<%
End Sub

'==========================================
'函 数 名：ArticleAddForm()
'作    用：添加内容表单
'参    数：
'==========================================
Sub ArticleAddForm()
	
	dim rnd_num
	RANDOMIZE
	rnd_num=INT(100*RND)+1
	Fk_Module_Id=Clng(Trim(Request.QueryString("ModuleId")))
	'判断权限
	'If Not FkFun.CheckLimit("Module"&Fk_Module_Id) Then
		'Response.Write("无权限！")
		'Call FKDB.DB_Close()
		'Session.CodePage=936
		'Response.End()
	'End If
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id=" & Fk_Module_Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Module_Name=Rs("Fk_Module_Name")
		Fk_Module_Menu=Rs("Fk_Module_Menu")
	End If
	Rs.Close
	appendFrom="本文由"&SiteName&"发表。转载此文章须经作者同意，并请附上出处("&SiteName&")及本页链接。原文链接:{$originalUrl}"
%>

<script type="text/javascript" charset="utf-8" id="colorpickerjs" src="/js/colorpicker.js"></script>
<script type="text/javascript">
$(document).ready(function(){
	$("#colorpickerjs").attr("src","/js/colorpicker.js?r="+Math.random());
	set_title_color("");
	$("#Fk_Article_Color").keyup(function(){
		if($('#Fk_Article_Color').val()==""){
			set_title_color("");
		}
	})
})
	function set_title_color(color) {
		$('#Fk_Article_Title').css('color',color);
		$('#Fk_Article_Color').val(color);
	}
	function set_color(id,color) {
		$('#'+id).css('color',color);
		$('#'+id).val(color);
	}
	
	if(window.KindEditor){
		$("#Fk_Article_Pic").after(" <input type=\"button\" id=\"uploadButton\" value=\"上传\" class=\"Button\" style=\"vertical-align:middle;\"/> *上传限于gif|jpg|jpeg|png|bmp格式,文件最大允许2M");
			var editor = window.KindEditor.editor({
					fileManagerJson : '/admin/dkidtioenr/aps/file_manager_json.asp',
					uploadJson		: '/admin/dkidtioenr/aps/upload_json.asp?dir=image',
					allowFileManager : true
				});
				$('#uploadButton').click(function() {
					editor.loadPlugin('image', function() {
						editor.plugin.imageDialog({
							imageUrl : $('#Fk_Article_Pic').val(),
							clickFn : function(url) {
								$('#Fk_Article_Pic').val(url);
								editor.hideDialog();
							}
						});
					});
				});
		}
		else
		{
			$("#Fk_Article_Pic").after(" <iframe frameborder=\"0\" width=\"290\" height=\"25\" scrolling=\"No\" id=\"I2\" name=\"I2\" src=\"PicUpLoad.asp?Form=ArticleAdd&Input=Fk_Article_Pic\" style=\"vertical-align:middle\"></iframe> *上传限于gif|jpg|jpeg|png|bmp格式,文件最大允许200K");
		}
		$("#selectfs").change(function(){
			$("#Fk_Article_CopyrightInfo").css("font-size",$(this).val());
		})
		$("#selectwt").change(function(){
			$("#Fk_Article_CopyrightInfo").css("font-weight",$(this).val());
		})
		$("#selectcl").change(function(){
			$("#Fk_Article_CopyrightInfo").css("color",$(this).val());
		})
		$("input[name=Fk_Article_Copyright]").click(function(){
			if($(this).val()=='1'){
				$(".zzsm").show();
			}
			else{
				$(".zzsm").hide();
			}
		})
</script>
<style type="text/css">
.blue{color:blue}
.red{color:red}
.green{color:green}
.yellow{color:yellow}
.gray{color:gray}
.black{color:black}
.Purple{color:Purple}
.zzsm{display:none;}
</style>
<form id="ArticleAdd" name="ArticleAdd" method="post" action="Article.asp?Type=3" onsubmit="return false;">
<div id="BoxTop" style="width:98%;"><span>添加</span><a onclick="$('#Boxs').hide();$('select').show();"><img src="images/close3.gif"></a></div>
<div id="BoxContents" style="width:98%;">
<table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td height="25" align="right">标题：</td>
        <td><input name="Fk_Article_Title"<%If SiteToPinyin=1 Then%> onchange="GetPinyin('Fk_Article_FileName','ToPinyin.asp?Str='+this.value);"<%End If%> type="text" class="Input" id="Fk_Article_Title" size="50"  style="vertical-align:middle;"/><span class="colorpicker" onclick="colorpicker('colorpanel_title','set_title_color');" title="标题颜色" style="vertical-align:middle;background: url(http://image001.dgcloud01.qebang.cn/website/ext/images/icon/color.png) 0px 0px no-repeat;height: 24px;width: 24px;display: inline-block;"></span>
                           <span class="colorpanel" id="colorpanel_title" style="position:absolute;z-index:99999999;"></span> <input name="Fk_Article_Color" type="text" class="Input" id="Fk_Article_Color" style="vertical-align:middle;width:60px;" value=""/> <script type="text/javascript">
						   set_title_color("");
                           </script></td>
        <td align="right" colspan="2">排序：</td>
        <td><input name="Fk_Article_px" type="text" id="Fk_Article_px" class="Input" size="6" maxlength="6" value="0" style="vertical-align:middle;"> （仅限数字，越大越排前）</td>
      </tr>
	<tr>
		<td height="28" align="right">SEO标题：</td>
		<td colspan="3"><input name="Fk_Article_Seotitle" value="" type="text" class="Input" id="Fk_Article_Seotitle" size="60"  style="vertical-align:middle;"/></td>
	</tr>
    <tr>
        <td height="25" align="right">SEO关键词：</td>
        <td><input name="Fk_Article_Keyword" type="text" class="Input" id="Fk_Article_Keyword" size="60"  style="vertical-align:middle;"/> <input type="button" onclick="tiqu(0,'Fk_Article_Content','Fk_Article_Keyword');" class="Button" name="btnTqkeyw" id="btnTqkeyw" value="提 取"  style="vertical-align:middle;"/><input  value="<%=KeyWordlist%>" type="hidden" id="Fk_Keywordlist" /></td>
        <td align="right" colspan="2">文件名：</td>
        <td>
		<input name="Fk_Article_FileName" type="text" class="Input" id="Fk_Article_FileName" value="" size="30"  style="vertical-align:middle;"/>		&nbsp;*不可修改</td>
      </tr>
    <tr>
        <td height="25" align="right">SEO描述：</td>
        <td>
		<input name="Fk_Article_Description" type="text" class="Input" id="Fk_Article_Description" size="60"  style="vertical-align:middle;"/> <input type="button" onclick="tiqu(1,'Fk_Article_Content','Fk_Article_Description');" class="Button" name="btnTqdesc" id="btnTqdesc" value="提 取"  style="vertical-align:middle;"/></td>
        <td align="right">点击量：</td>
        <td colspan="5"><input name="Fk_Article_click" type="text" id="Fk_Article_click" class="Input" size="6" maxlength="6" value="<%=rnd_num%>" style="vertical-align:middle;"></td>
    </tr>
   <tr>
        <td height="28" align="right">图片：<br><div id="xtu" style="display:none;position:absolute;border:1px solid gray;padding:2px;background:white;left:10px;top:190px;">图片预览[<a href="#" onclick="document.getElementById('xtu').style.display='none'">关闭</a>]：<br><img id="showimg_id" src="" onclick="this.src=document.getElementById('Fk_Article_Pic').value;" style="cursor:pointer;width:400px;" title="上传后点击图片更新预览 " border="0"></div></td>
        <td colspan="6"><input name="Fk_Article_Pic" type="text" class="Input" id="Fk_Article_Pic" size="50" onmousedown="document.getElementById('showimg_id').src=this.value;document.getElementById('xtu').style.display='block'"  style="vertical-align:middle;"/><input style="display:none" name="Fk_Article_PicBig" type="text" class="Input" id="Fk_Article_PicBig" size="50"  onmousedown="document.getElementById('dtu').style.display='block';" /></td>
        
    </tr>
       <tr>
        <td height="25" align="right" width="100">内容：</td>
        <td colspan="4"><textarea name="Fk_Article_Content" class="<%=bianjiqi%>" style="width:100%;" rows="15" id="Fk_Article_Content"></textarea></td>
    </tr>
<%
	Sqlstr="Select * From [Fk_Field] Where Fk_Field_Type=0 Order By Fk_Field_Id Asc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
%>
    <tr>
        <td height="25" align="right"><%=Rs("Fk_Field_Name")%>：</td>
        <td colspan="4"><input name="Fk_Article__<%=Rs("Fk_Field_Tag")%>" type="text" class="Input" id="Fk_Article__<%=Rs("Fk_Field_Tag")%>"  style="vertical-align:middle;"/></td>
    </tr>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
    <tr>
        <td height="25" align="right">来源：</td>
        <td colspan="4">
		<input name="Fk_Article_From" type="text" class="Input" id="Fk_Article_From" value="本站" size="20"  style="vertical-align:middle;"/>　转向链接：<input name="Fk_Article_Url" type="text" class="Input" id="Fk_Article_Url" size="60"  style="vertical-align:middle;"/>&nbsp;*正常请留空</td>
    </tr>
       <tr>
        <td height="25" align="right">推荐：</td>
        <td colspan="4"><select name="Fk_Article_Recommend"  id="Fk_Article_Recommend" style="vertical-align:middle;" class="Input">
            <option value="0">无推荐</option>
            <option value="2">推荐</option>
            </select>
          <input type="checkbox" name="Fk_Article_onTop" id="Fk_Article_onTop" class="textarea" value="1"  style="vertical-align:middle;"/><label for="Fk_Article_onTop" style="vertical-align:middle;">置顶</label>
　是否显示：<input name="Fk_Article_Show" class="Input" type="radio" id="Fk_Article_Show" value="1" checked="true"  style="vertical-align:middle;"/><label for="Fk_Article_Show" style="vertical-align:middle;">显示</label>
        <input type="radio" class="Input" name="Fk_Article_Show" id="Fk_Article_Show1" value="0" style="vertical-align:middle;"/><label for="Fk_Article_Show1" style="vertical-align:middle;">不显示</label>　模板：<select name="Fk_Article_Template" class="Input" id="Fk_Article_Template" style="vertical-align:middle;">
            <option value="0"<%=FKFun.BeSelect(Fk_Article_Template,0)%>>默认模板</option>
<%
	Sqlstr="Select * From [Fk_Template] Where Not Fk_Template_Name In ('index','info','articlelist','article','productlist','product','gbook','page','subject','job','subject','top','bottom','downlist','down')"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
%>
            <option value="<%=Rs("Fk_Template_Id")%>"><%=Rs("Fk_Template_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
            </select></td>
    </tr>
    <tr>
        <td align="right">转载声明：</td><td><input name="Fk_Article_Copyright" class="Input Fk_Article_Copyright" type="radio" id="Fk_Article_Copyright" value="1" style="vertical-align:middle;"/><label for="Fk_Article_Copyright" style="vertical-align:middle;">显示</label>
        <input type="radio" class="Input Fk_Article_Copyright" checked="true" name="Fk_Article_Copyright" id="Fk_Article_Copyright1" value="0" style="vertical-align:middle;"/><label for="Fk_Article_Copyright1" style="vertical-align:middle;">不显示</label> <img src="http://image001.dgcloud01.qebang.cn/website/new.png" title="新功能：选择显示则自动在文章末尾追加转载声明，不显示则不追加转载声明" style="vertical-align:middle;"></td>
    </tr>
    <tr class="zzsm">
        <td align="right">声明内容：</td>
        <td colspan="4">*<b style="color:red">{$originalUrl}标签会自动生成当前文章链接，请勿改动</b><br><textarea name="Fk_Article_CopyrightInfo" style="border: 1px solid #D3E3F0;width:100%;padding-top:5px;padding-left:5px;height:90px;" id="Fk_Article_CopyrightInfo"><%=appendFrom%></textarea><div style="margin-top:4px;"><label for="selectfs" style="vertical-align:middle;">字号:</label> <select id="selectfs" name="selectfs" style="vertical-align:middle;"><option value="12px">默认</option><option value="10px">10px</option><option value="11px">11px</option><option value="12px">12px</option><option value="13px">13px</option><option value="14px">14px</option><option value="15px">15px</option><option value="16px">16px</option><option value="17px">17px</option><option value="18px">18px</option><option value="19px">19px</option><option value="20px">20px</option><option value="21px">21px</option><option value="22px">22px</option><option value="23px">23px</option><option value="24px">24px</option></select> <label for="selectwt" style="vertical-align:middle;">粗体:</label> <select name="selectwt" id="selectwt" style="vertical-align:middle;"><option value="normal">默认</option><option value="bolder">加粗</option></select> <label for="selectcl" style="vertical-align:middle;">文字颜色:</label> <select name="selectcl" id="selectcl" style="vertical-align:middle;"><option value="">默认</option>
					<option value="blue" class="blue">蓝色</option>
					<option value="red" class="red">红色</option>
					<option value="green" class="green">绿色</option>
					<option value="yellow" class="yellow">黄色</option>
					<option value="gray" class="gray">灰色</option>
					<option value="black" class="black">黑色</option>
					<option value="Purple" class="Purple">紫色</option></select><div></td>
    </tr>
   </table>
</div>
<div id="BoxBottom" style="width:96%;">
		<input type="hidden" name="ModuleId" value="<%=Fk_Module_Id%>" />
        <input type="submit" onclick="Sends('ArticleAdd','Article.asp?Type=3',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" class="Button" name="button" id="button" value="添 加" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="btnClose" id="btnClose" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：ArticleAddDo
'作    用：执行添加内容
'参    数：
'==============================
Sub ArticleAddDo()
	Fk_Module_Id=Trim(Request.Form("ModuleId"))
	'判断权限
	'If Not FkFun.CheckLimit("Module"&Fk_Module_Id) Then
		'Response.Write("无权限！")
		'Call FKDB.DB_Close()
		'Session.CodePage=936
		'Response.End()
	'End If
	dim Fk_Article_syn,Syn_type
	Fk_Article_Title=FKFun.HTMLEncode(Trim(Request.Form("Fk_Article_Title")))
	Fk_Article_Color=FKFun.HTMLEncode(Trim(Request.Form("Fk_Article_Color")))
	Fk_Article_Seotitle=FKFun.HTMLEncode(Trim(Request.Form("Fk_Article_Seotitle")))
	Fk_Article_Keyword=FKFun.HTMLEncode(Trim(Request.Form("Fk_Article_Keyword")))
	Fk_Article_Description=FKFun.HTMLEncode(Trim(Request.Form("Fk_Article_Description")))
	Fk_Article_Url=FKFun.HTMLEncode(Trim(Request.Form("Fk_Article_Url")))
	Fk_Article_From=FKFun.HTMLEncode(Trim(Request.Form("Fk_Article_From")))
	Fk_Article_Content=Request.Form("Fk_Article_Content")
	
	Fk_Article_Pic=FKFun.HTMLEncode(Trim(Request.Form("Fk_Article_Pic")))
	Fk_Article_PicBig=FKFun.HTMLEncode(Trim(Request.Form("Fk_Article_PicBig")))
	Fk_Article_FileName=FKFun.HTMLEncode(Trim(Request.Form("Fk_Article_FileName")))
	Fk_Article_Recommend=","&FKFun.HTMLEncode(Replace(Trim(Request.Form("Fk_Article_Recommend"))," ",""))&","
	Fk_Article_Subject=","&FKFun.HTMLEncode(Replace(Trim(Request.Form("Fk_Article_Subject"))," ",""))&","
	Fk_Article_Template=Trim(Request.Form("Fk_Article_Template"))
	Fk_Article_Show=Trim(Request.Form("Fk_Article_Show"))
	Fk_Article_onTop=Trim(Request.Form("Fk_Article_onTop"))
	Fk_Article_px=Trim(Request.Form("Fk_Article_px"))
	Fk_Article_click=Trim(Request.Form("Fk_Article_click"))
	Fk_Article_Copyright=Request.Form("Fk_Article_Copyright")
	Fk_Article_CopyrightInfo=Trim(Request.Form("Fk_Article_CopyrightInfo"))
	Fk_Article_CopyrightFs=Trim(Request.Form("selectfs"))
	Fk_Article_CopyrightFt=Trim(Request.Form("selectwt"))
	Fk_Article_CopyrightCl=Trim(Request.Form("selectcl"))
	'Fk_Article_syn=Trim(Request.Form("Fk_Article_syn"))
	'Syn_type=Trim(Request.Form("Syn_type"))
	
	
	
	dim Fk_Article_Copyright,ArticleUrl
	dim mrs,Fk_Module_Dir
	if Fk_Article_Url="" then
			set mrs=conn.execute("select Fk_Module_Dir from Fk_Module where Fk_Module_Id="&Fk_Module_Id)
			if not mrs.eof then
				Fk_Module_Dir=mrs("Fk_Module_Dir")
			end if
			mrs.close
			set mrs=nothing
			If Fk_Module_Dir<>"" Then
				ArticleUrl=Fk_Module_Dir&"/"
			Else
				ArticleUrl="Article"&Fk_Module_Id&"/"
			End If
			If Fk_Article_FileName<>"" Then
				ArticleUrl=ArticleUrl&Fk_Article_FileName&".html"
			Else
				ArticleUrl=ArticleUrl&Fk_Module_Id&".html"
			End If
			If SiteHtml=1 Then
				ArticleUrl="/html"&SiteDir&ArticleUrl
			Else
				ArticleUrl=SiteDir&sTemp&"?"&ArticleUrl
			End If
	end if
	
	
	Call FKFun.ShowString(Fk_Article_Title,1,255,0,"请输入内容标题！","内容标题不能大于255个字符！")
	Call FKFun.ShowString(Fk_Article_From,1,50,0,"请输入内容来源！","内容来源不能大于50个字符！")
	Call FKFun.ShowNum(Fk_Article_px,"排序必须为数字！")
	Call FKFun.ShowString(Fk_Article_Seotitle,0,255,2,"请输入内容SEO标题！","内容SEO标题不能大于255个字符！")
	Call FKFun.ShowString(Fk_Article_Keyword,0,255,2,"请输入内容SEO关键词！","内容SEO关键词不能大于255个字符！")
	Call FKFun.ShowString(Fk_Article_Description,0,255,2,"请输入内容SEO描述！","内容SEO描述不能大于255个字符！")
	Call FKFun.ShowString(Fk_Article_Url,0,255,2,"请输入内容转向链接！","内容转向链接不能大于255个字符！")
	If Fk_Article_Url="" Then
		Call FKFun.ShowString(Fk_Article_Content,10,1,1,"请输入内容内容，不少于10个字符！","内容内容不能大于1个字符！")
	End If
	Call FKFun.ShowString(Fk_Article_Pic,0,255,2,"请输入内容题图路径！","内容题图小图路径不能大于255个字符！")
	Call FKFun.ShowString(Fk_Article_PicBig,0,255,2,"请输入内容题图路径！","内容题图大图路径不能大于255个字符！")
	Call FKFun.ShowString(Fk_Article_FileName,0,100,2,"生成文件名不符合标准！","内容文件名不能大于100个字符！")
	'Call FKFun.ShowString(Fk_Article_FileName,2,1,1,"生成文件名不能为空","生成内容不能大于1个字符！")
	Call FKFun.ShowNum(Fk_Article_Template,"请选择模板！")
	Call FKFun.ShowNum(Fk_Article_Show,"请选择内容是否显示！")
	Call FKFun.ShowNum(Fk_Article_click,"请输入正确的点击量！")
	Call FKFun.ShowNum(Fk_Article_Copyright,"参数错误，请刷新页面！")
	Call FKFun.ShowString(Fk_Article_CopyrightInfo,0,200,0,"请输入转载声明！","转载声明内容不能大于200个字符！")
	Call FKFun.ShowString(Fk_Article_CopyrightFs,0,50,0,"请选择字体大小！","字体大小不能大于50个字符！")
	Call FKFun.ShowString(Fk_Article_CopyrightFt,0,50,0,"请选择是否加粗！","粗体样式不能大于50个字符！")
	Call FKFun.ShowString(Fk_Article_CopyrightCl,0,50,0,"请选择字体颜色！","字体颜色不能大于50个字符！")
	Call FKFun.ShowNum(Fk_Module_Id,"ModuleId系统参数错误，请刷新页面！")
	'if Fk_Article_syn="1" then			'暂时关闭
		'Call FKFun.ShowNum(Syn_type,"请选择要同步到的行业类型")
	'end if
	Sqlstr="Select * From [Fk_Field] Where Fk_Field_Type=0 Order By Fk_Field_Id Asc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
		If Fk_Article_Field="" Then
			Fk_Article_Field=Rs("Fk_Field_Tag")&"|-Fangka_Field-|"&FKFun.HTMLEncode(Trim(Request.Form("Fk_Article__"&Rs("Fk_Field_Tag"))))
		Else
			Fk_Article_Field=Fk_Article_Field&"[-Fangka_Field-]"&Rs("Fk_Field_Tag")&"|-Fangka_Field-|"&FKFun.HTMLEncode(Trim(Request.Form("Fk_Article__"&Rs("Fk_Field_Tag"))))
		End If
		Rs.MoveNext
	Wend
	Rs.Close
	If IsNumeric(Fk_Article_FileName) Then
		Response.Write("文件名不可用纯数字！")
		Call FKDB.DB_Close()
		Response.End()
	End If
	If Left(Fk_Article_FileName,4)="Info" Or Left(Fk_Article_FileName,4)="Page" Or Left(Fk_Article_FileName,5)="GBook" Or Left(Fk_Article_FileName,3)="Job" Then
		Response.Write("文件名受限，不能以一下单词开头：Info、Page、GBook、Job！")
		Call FKDB.DB_Close()
		Response.End()
	End If
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id=" & Fk_Module_Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Module_Menu=Rs("Fk_Module_Menu")
	Else
		Response.Write("栏目不存在！")
		Rs.Close
		Call FKDB.DB_Close()
		Response.End()
	End If
	Rs.Close
	If SiteDelWord=1 Then
		TempArr=Split(Trim(FKFun.UnEscape(FKFso.FsoFileRead("DelWord.dat")))," ")
		For Each Temp In TempArr
			If Temp<>"" Then
				Fk_Article_Content=Replace(Fk_Article_Content,Temp,"**")
				Fk_Article_Title=Replace(Fk_Article_Title,Temp,"**")
				Fk_Article_Keyword=Replace(Fk_Article_Keyword,Temp,"**")
				Fk_Article_Description=Replace(Fk_Article_Description,Temp,"**")
			End If
		Next
	End If
	
	
	'新功能，追加转载声明
	'2014年12月31日
	'middy241@163.com
	if CheckFields("Fk_Article_Copyright","Fk_Article")=false then
		conn.execute("alter table Fk_Article add column Fk_Article_Copyright int default 0")
		conn.execute("alter table Fk_Article add column Fk_Article_CopyrightInfo varchar(200) null")
		conn.execute("alter table Fk_Article add column Fk_Article_CopyrightFs varchar(50) null")
		conn.execute("alter table Fk_Article add column Fk_Article_CopyrightFt varchar(50) null")
		conn.execute("alter table Fk_Article add column Fk_Article_CopyrightCl varchar(50) null")
	end if
	
	Sqlstr="Select * From [Fk_Article] Where Fk_Article_Module="&Fk_Module_Id&" And (Fk_Article_Title='"&Fk_Article_Title&"'"
	If Fk_Article_FileName<>"" Then
		Sqlstr=Sqlstr&" Or Fk_Article_FileName='"&Fk_Article_FileName&"'"
	End If
	Sqlstr=Sqlstr&")"
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("Fk_Article_Title")=Fk_Article_Title
		Rs("Fk_Article_Color")=Fk_Article_Color
		Rs("Fk_Article_From")=Fk_Article_From
		Rs("Fk_Article_Seotitle")=Fk_Article_Seotitle
		Rs("Fk_Article_Keyword")=Fk_Article_Keyword
		Rs("Fk_Article_Field")=Fk_Article_Field
		Rs("Fk_Article_Description")=Fk_Article_Description
		Rs("Fk_Article_Url")=Fk_Article_Url
		Rs("Fk_Article_Show")=Fk_Article_Show
		Rs("Fk_Article_click")=Fk_Article_click
		Rs("Fk_Article_Pic")=Fk_Article_Pic
		Rs("Fk_Article_PicBig")=Fk_Article_PicBig
		Rs("Fk_Article_Content")=Fk_Article_Content
		Rs("Fk_Article_Recommend")=Fk_Article_Recommend
		Rs("Fk_Article_Subject")=Fk_Article_Subject
		Rs("Fk_Article_Module")=Fk_Module_Id
		Rs("Fk_Article_Menu")=Fk_Module_Menu
		Rs("Fk_Article_FileName")=Fk_Article_FileName
		Rs("Fk_Article_Template")=Fk_Article_Template
		Rs("Fk_Article_Ip")=Fk_Article_onTop
		
		Rs("Fk_Article_Copyright")=Fk_Article_Copyright
		Rs("Fk_Article_CopyrightInfo")=Fk_Article_CopyrightInfo
		Rs("Fk_Article_CopyrightFs")=Fk_Article_CopyrightFs
		Rs("Fk_Article_CopyrightFt")=Fk_Article_CopyrightFt
		Rs("Fk_Article_CopyrightCl")=Fk_Article_CopyrightCl
		
		Rs("Px")=Fk_Article_px
		Rs.Update()
		Application.UnLock()
		Response.Write("新内容添加成功！")
		'插入日志
		on error resume next
		dim log_content,log_ip,log_user
		log_content="添加内容：【"&Fk_Article_Title&"】"
		log_user=Request.Cookies("FkAdminName")
		
		log_ip=FKFun.getIP()
		conn.execute("insert into newTB_log (log_user,log_content,log_ip) values('"&log_user&"','"&log_content&"','"&log_ip&"')")
		'if Fk_Article_syn="1" then
			' dim reqHandler,key,host
			' host=request.ServerVariables("HTTP_HOST")
			' if host="localhost" or host="127.0.0.1" then
			' else
			' key 	= "85e5ffb11e1c4a8561b953a7e27a547c"
			' set reqHandler = new SyncRequestHandler
			' '初始化
			' reqHandler.init()
			' '设置密钥
			' reqHandler.setKey(key)
			' '-----------------------------
			' '设置同步参数
			' '-----------------------------
			' reqHandler.setParameter "tit", Fk_Article_Title		'标题
			' reqHandler.setParameter "con",FKFun.RemoveHTML(Fk_Article_Content)		'内容
			' reqHandler.setParameter "typ", "-1"		'类型

			' 'response.Cookies(host&"_synTradeName")=""
			' 'response.Cookies(host&"_synTradeId")=""
			
			' '请求的参数
			' Dim Para,return,SyncUrl
			' reqHandler.setParameter "hos", host		'域名
			' Para  	= reqHandler.getParameters()
			' SyncUrl	="http://qbknow.qb02.com/json/sync_article.asp"
			' return	= reqHandler.PostHttpPage("qbknow.qb02.com",SyncUrl,Para)
			' end if
			'Response.Write(return)
		'end if
	Else
		Response.Write("该内容标题已经被占用，请重新选择！")
	End If
	Rs.Close
End Sub

'==========================================
'函 数 名：ArticleEditForm()
'作    用：修改内容表单
'参    数：
'==========================================
Sub ArticleEditForm()

	'新功能，追加转载声明
	'2014年12月31日
	'middy241@163.com
	if CheckFields("Fk_Article_Copyright","Fk_Article")=false then
		conn.execute("alter table Fk_Article add column Fk_Article_Copyright int default 0")
		conn.execute("alter table Fk_Article add column Fk_Article_CopyrightInfo varchar(200) null")
		conn.execute("alter table Fk_Article add column Fk_Article_CopyrightFs varchar(50) null")
		conn.execute("alter table Fk_Article add column Fk_Article_CopyrightFt varchar(50) null")
		conn.execute("alter table Fk_Article add column Fk_Article_CopyrightCl varchar(50) null")
	end if

	dim Fk_Module_Dir,ArticleUrl
	Id=Clng(Request.QueryString("Id"))
	Fk_Module_Id=Clng(Trim(Request.QueryString("ModuleId")))
	'判断权限
	If Not FkFun.CheckLimit("Module"&Fk_Module_Id) Then
		'Response.Write("无权限！")
		'Call FKDB.DB_Close()
		'Session.CodePage=936
		'Response.End()
	End If
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id=" & Fk_Module_Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Module_Name=Rs("Fk_Module_Name")
		Fk_Module_Dir=Rs("Fk_Module_Dir")
		Fk_Module_Id=Rs("Fk_Module_Id")
	End If
	Rs.Close
	Sqlstr="Select * From [Fk_Article] Where Fk_Article_Id=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Article_Title=Rs("Fk_Article_Title")
		Fk_Article_Color=Rs("Fk_Article_Color")
		Fk_Article_From=Rs("Fk_Article_From")
		Fk_Article_Seotitle=Rs("Fk_Article_Seotitle")
		Fk_Article_Keyword=Rs("Fk_Article_Keyword")
		Fk_Article_Description=Rs("Fk_Article_Description")
		Fk_Article_Content=Rs("Fk_Article_Content")
		Fk_Article_Url=Rs("Fk_Article_Url")
		Fk_Article_Pic=Rs("Fk_Article_Pic")
		Fk_Article_PicBig=Rs("Fk_Article_PicBig")
		Fk_Article_Recommend=Rs("Fk_Article_Recommend")
		Fk_Article_Subject=Rs("Fk_Article_Subject")
		Fk_Article_Show=Rs("Fk_Article_Show")
		Fk_Article_Click=Rs("Fk_Article_Click")
		Fk_Article_Template=Rs("Fk_Article_Template")
		Fk_Article_FileName=Rs("Fk_Article_FileName")
		Fk_Article_onTop=trim(Rs("Fk_Article_Ip")&" ")
		Fk_Article_Copyright=trim(Rs("Fk_Article_Copyright")&" ")
		Fk_Article_CopyrightInfo=trim(Rs("Fk_Article_CopyrightInfo")&" ")
		if Fk_Article_CopyrightInfo="" then
			Fk_Article_CopyrightInfo="本文由"&SiteName&"发表。转载此文章须经作者同意，并请附上出处("&SiteName&")及本页链接。原文链接:{$originalUrl}"
		end if
		Fk_Article_CopyrightFs=trim(Rs("Fk_Article_CopyrightFs")&" ")
		Fk_Article_CopyrightFt=trim(Rs("Fk_Article_CopyrightFt")&" ")
		Fk_Article_CopyrightCl=trim(Rs("Fk_Article_CopyrightCl")&" ")
		Fk_Article_px=Rs("Px")
		If IsNull(Rs("Fk_Article_Field")) Or Rs("Fk_Article_Field")="" Then
			Fk_Article_Field=Split("-_-|-Fangka_Field-|1")
		Else
			Fk_Article_Field=Split(Rs("Fk_Article_Field"),"[-Fangka_Field-]")
		End If
	End If
	Rs.Close

%>
<script type="text/javascript" charset="utf-8" id="colorpickerjs" src="/js/colorpicker.js"></script>
<script type="text/javascript">
$(document).ready(function(){
	$("#colorpickerjs").attr("src","/js/colorpicker.js?r="+Math.random());
	set_title_color("");
	$("#Fk_Article_Color").keyup(function(){
		if($('#Fk_Article_Color').val()==""){
			set_title_color("");
		}
	})
	
	$("#Fk_Article_CopyrightInfo").css("font-size","<%=Fk_Article_CopyrightFs%>");
	$("#Fk_Article_CopyrightInfo").css("font-weight","<%=Fk_Article_CopyrightFt%>");
	$("#Fk_Article_CopyrightInfo").css("color","<%=Fk_Article_CopyrightCl%>");
	if($("input[name=Fk_Article_Copyright]:checked").val()==1){
		$(".zzsm").show();
	}

	$("#selectfs").change(function(){
		$("#Fk_Article_CopyrightInfo").css("font-size",$(this).val());
	})
	$("#selectwt").change(function(){
		$("#Fk_Article_CopyrightInfo").css("font-weight",$(this).val());
	})
	$("#selectcl").change(function(){
		$("#Fk_Article_CopyrightInfo").css("color",$(this).val());
	})
	$("input[name=Fk_Article_Copyright]").click(function(){
		if($(this).val()=='1'){
			$(".zzsm").show();
		}
		else{
			$(".zzsm").hide();
		}
	})
})
	function set_title_color(color) {
		$('#Fk_Article_Title').css('color',color);
		$('#Fk_Article_Color').val(color);
	}
	
	if(window.KindEditor){
		$("#Fk_Article_Pic").after(" <input type=\"button\" id=\"uploadButton\" value=\"上传\" class=\"Button\" style=\"vertical-align:middle;\"/> *上传限于gif|jpg|jpeg|png|bmp格式,文件最大允许2M");
			var editor = window.KindEditor.editor({
					fileManagerJson : '/admin/dkidtioenr/aps/file_manager_json.asp',
					uploadJson		: '/admin/dkidtioenr/aps/upload_json.asp?dir=image',
					allowFileManager : true
				});
				$('#uploadButton').click(function() {
					editor.loadPlugin('image', function() {
						editor.plugin.imageDialog({
							imageUrl : $('#Fk_Article_Pic').val(),
							clickFn : function(url) {
								$('#Fk_Article_Pic').val(url);
								editor.hideDialog();
							}
						});
					});
				});
				
				$('.Fk_Article_Copyright').click(function() {
					//alert($("#Fk_Article_Content").html());
					//$("#Fk_Article_Content").remove(".article_copyright_class");
				});
		}
		else
		{
			$("#Fk_Article_Pic").after(" <iframe frameborder=\"0\" width=\"290\" height=\"25\" scrolling=\"No\" id=\"I2\" name=\"I2\" src=\"PicUpLoad.asp?Form=ArticleAdd&Input=Fk_Article_Pic\" style=\"vertical-align:middle\"></iframe> *上传限于gif|jpg|jpeg|png|bmp格式,文件最大允许200K");
		}


</script>
<style type="text/css">
.blue{color:blue}
.red{color:red}
.green{color:green}
.yellow{color:yellow}
.gray{color:gray}
.black{color:black}
.Purple{color:Purple}
.zzsm{display:none;}
</style>
<form id="ArticleEdit" name="ArticleEdit" method="post" action="Article.asp?Type=5" onsubmit="return false;">
<div id="BoxTop" style="width:98%;"><span>修改</span><a onclick="$('#Boxs').hide();$('select').show();"><img src="images/close3.gif"></a></div>
<div id="BoxContents" style="width:98%;">
<table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td height="25" align="right">标题：</td>
        <td><input name="Fk_Article_Title"<%If SiteToPinyin=1 Then%> onmouseout="GetPinyin('Fk_Article_FileName','ToPinyin.asp?Str='+this.value);" onchange="GetPinyin('Fk_Article_FileName','ToPinyin.asp?Str='+this.value);"<%End If%> value="<%=Fk_Article_Title%>" type="text" class="Input" id="Fk_Article_Title" size="50"  style="vertical-align:middle;"/><span class="colorpicker" onclick="colorpicker('colorpanel_title','set_title_color');" title="标题颜色" style="vertical-align:middle;background: url(http://image001.dgcloud01.qebang.cn/website/ext/images/icon/color.png) 0px 0px no-repeat;height: 24px;width: 24px;display: inline-block;"></span>
                           <span class="colorpanel" id="colorpanel_title" style="position:absolute;z-index:99999999;"></span><input name="Fk_Article_Color" type="text" class="Input" id="Fk_Article_Color" style="vertical-align:middle;width:50px;" value=""/><script type="text/javascript">
						   set_title_color("<%=Fk_Article_Color%>");
                           </script>
        </td>
        <td colspan="2"><input type="checkbox" name="Fk_Article_Time" id="Fk_Article_Time" value="1"  style="vertical-align:middle;"/><label for="Fk_Article_Time" style="vertical-align:middle;">更新时间</label> &nbsp; &nbsp;  &nbsp; &nbsp; 点击量：<input name="Fk_Article_click" type="text" id="Fk_Article_click" class="Input" size="6" maxlength="6" value="<%=Fk_Article_Click%>" style="vertical-align:middle;"></td>
    </tr>
	<tr>
		<td height="28" align="right">SEO标题：</td>
		<td colspan="3"><input name="Fk_Article_Seotitle" value="<%=Fk_Article_Seotitle%>" type="text" class="Input" id="Fk_Article_Seotitle" size="60"  style="vertical-align:middle;"/></td>
	</tr>
    <tr>
        <td height="25" align="right">SEO关键词：</td>
        <td><input name="Fk_Article_Keyword" value="<%=Fk_Article_Keyword%>" type="text" class="Input" id="Fk_Article_Keyword" size="60"  style="vertical-align:middle;"/> 
        <input type="button" onclick="tiqu(0,'Fk_Article_Content','Fk_Article_Keyword');" class="Button" name="btntqKeyw" id="btntqKeyw" value="提 取"  style="vertical-align:middle;"/><input  value="<%=KeyWordlist%>" type="hidden" id="Fk_Keywordlist" /></td>
        <td align="right">排 &nbsp; 序：</td>
        <td><input name="Fk_Article_px" type="text" id="Fk_Article_px" class="Input" size="4" maxlength="6" value="<%if isnull(Fk_Article_px) then response.write "0" else response.write Fk_Article_px%>" style="vertical-align:middle;">  (限数字,越大越排前)</td>
    </tr>
    <tr>
        <td height="25" align="right">SEO描述：</td>
        <td><input name="Fk_Article_Description" value="<%=Fk_Article_Description%>" type="text" class="Input" size="60" id="Fk_Article_Description"  style="vertical-align:middle;"/> 
        <input type="button" onclick="tiqu(1,'Fk_Article_Content','Fk_Article_Description');" class="Button" name="btntqDesc" id="btntqDesc" value="提 取"  style="vertical-align:middle;"/></td>
		<td align="right">文件名：</td>
		<td><input name="Fk_Article_FileName" type="text" class="Input" id="Fk_Article_FileName" size="30" value="<%=Fk_Article_FileName%>"  style="vertical-align:middle;"/>*不可修改</td>
    </tr>
     <tr>
        <td height="28" align="right">图片：<br><div id="xtu" style="display:none;position:absolute;border:1px solid gray;padding:2px;background:white;left:10px;top:190px;">缩略小图预览[<a href="#" onclick="document.getElementById('xtu').style.display='none'">关闭</a>]：<br><img src="<%=Fk_Article_Pic%>" onclick="this.src=document.getElementById('Fk_Article_Pic').value;" style="cursor:pointer;" title="上传后点击图片更新预览 " border="0"></div>
		</td>
        <td colspan="3"><input name="Fk_Article_Pic" type="text" class="Input" id="Fk_Article_Pic" value="<%=Fk_Article_Pic%>" size="50" onmousedown="document.getElementById('xtu').style.display='block'"  style="vertical-align:middle;"/><input style="display:none" name="Fk_Article_PicBig" type="text" class="Input" id="Fk_Article_PicBig" value="<%=Fk_Article_PicBig%>" size="50"  onmousedown="document.getElementById('dtu').style.display='block';" /></td>
     </tr>
        <tr>
        <td height="25" width="100" align="right">内容：</td>
        <td colspan="3"><textarea name="Fk_Article_Content" class="<%=bianjiqi%>" style="width:100%;" rows="15" id="Fk_Article_Content"><%=Fk_Article_Content%></textarea></td>
    </tr>
<%
	Sqlstr="Select * From [Fk_Field] Where Fk_Field_Type=0 Order By Fk_Field_Id Asc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
		Temp2=""
		For Each Temp In Fk_Article_Field
			If Split(Temp,"|-Fangka_Field-|")(0)=Rs("Fk_Field_Tag") Then
				Temp2=FKFun.HTMLDncode(Split(Temp,"|-Fangka_Field-|")(1))
				Exit For
			End If
		Next
%>
    <tr>
        <td height="25" align="right"><%=Rs("Fk_Field_Name")%>：</td>
        <td colspan="3"><input name="Fk_Article__<%=Rs("Fk_Field_Tag")%>" value="<%=Temp2%>" type="text" class="Input" id="Fk_Article__<%=Rs("Fk_Field_Tag")%>"  style="vertical-align:middle;"/></td>
    </tr>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
    <tr>
        <td height="25" align="right">来源：</td>
        <td colspan="3"><input name="Fk_Article_From" type="text" class="Input" id="Fk_Article_From0" value="<%=Fk_Article_From%>" size="20"  style="vertical-align:middle;"/>&nbsp;　转向链接：<input name="Fk_Article_Url" type="text" class="Input" value="<%=Fk_Article_Url%>" id="Fk_Article_Url" size="60"  style="vertical-align:middle;"/>&nbsp;*正常请留空</td>
    </tr>
       <tr>
        <td height="25" align="right">推荐：</td>
        <td colspan="3"><select name="Fk_Article_Recommend" class="Input" id="Fk_Article_Recommend" style="vertical-align:middle;">
            <option value="0">无推荐</option>
            <option value="2"<%If Instr(Fk_Article_Recommend,",2,")>0 Then%> selected="selected"<%End If%>>推荐</option>
            </select>
          <label><input type="checkbox" name="Fk_Article_onTop" id="Fk_Article_onTop" class="textarea" value="1" <%if Fk_Article_onTop="1" then response.write "checked"%> style="vertical-align:middle;"/>置顶</label>
			　显示：
			<input name="Fk_Article_Show" type="radio" class="Input" id="Fk_Article_Show" value="1"<%=FKFun.BeCheck(Fk_Article_Show,1)%>  style="vertical-align:middle;"/><label for="Fk_Article_Show" style="vertical-align:middle;">显示</label>
        <input type="radio" name="Fk_Article_Show" class="Input" id="Fk_Article_Show" value="0"<%=FKFun.BeCheck(Fk_Article_Show,0)%>  style="vertical-align:middle;"/><label for="Fk_Article_Show" style="vertical-align:middle;">不显示</label>　模板：<select name="Fk_Article_Template" id="Fk_Article_Template" style="vertical-align:middle;">
            <option value="0" <%=FKFun.BeSelect(Fk_Article_Template,0)%>>默认模板</option>
<%
	Sqlstr="Select * From [Fk_Template] Where Not Fk_Template_Name In ('index','info','articlelist','article','productlist','product','gbook','page','subject','job','subject')"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
%>
            <option value="<%=Rs("Fk_Template_Id")%>"<%=FKFun.BeSelect(Fk_Article_Template,Rs("Fk_Template_Id"))%>><%=Rs("Fk_Template_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
            </select></td>
    </tr>
    <tr>
        <td align="right">转载声明：</td><td><input name="Fk_Article_Copyright" class="Input Fk_Article_Copyright" type="radio" id="Fk_Article_Copyright" value="1" style="vertical-align:middle;" <%if Fk_Article_Copyright="1" then response.write "checked=""true"""%>/><label for="Fk_Article_Copyright" style="vertical-align:middle;">显示</label>
        <input type="radio" class="Input Fk_Article_Copyright" <%if Fk_Article_Copyright="0" or Fk_Article_Copyright="" then response.write "checked=""true"""%> name="Fk_Article_Copyright" id="Fk_Article_Copyright1" value="0" style="vertical-align:middle;"/><label for="Fk_Article_Copyright1" style="vertical-align:middle;">不显示</label> <img src="http://image001.dgcloud01.qebang.cn/website/new.png" title="新功能：选择显示则自动在文章末尾追加转载声明，不显示则不追加转载声明" style="vertical-align:middle;"></td>
    </tr>
    <tr class="zzsm">
        <td align="right">声明内容：</td>
        <td colspan="4">*<b style="color:red">{$originalUrl}标签会自动生成当前文章链接，请勿改动</b><br><textarea name="Fk_Article_CopyrightInfo" style="border: 1px solid #D3E3F0;width:100%;padding-top:5px;padding-left:5px;height:90px;" id="Fk_Article_CopyrightInfo"><%=Fk_Article_CopyrightInfo%></textarea><div style="margin-top:4px;"><label for="selectfs" style="vertical-align:middle;">字号:</label> <select id="selectfs" name="selectfs" style="vertical-align:middle;"><option value="" <%if Fk_Article_CopyrightFs="" then response.write "selected"%>>默认</option><%for i=10 to 24%><option value="<%=i%>px" <%if Fk_Article_CopyrightFs=i&"px" then response.write "selected"%>><%=i&"px"%></option><%next%></select> <label for="selectwt" style="vertical-align:middle;">粗体:</label> <select name="selectwt" id="selectwt" style="vertical-align:middle;"><option value="normal" <%if Fk_Article_CopyrightFt="normal" then response.write "selected"%>>默认</option><option value="bolder" <%if Fk_Article_CopyrightFt="bolder" then response.write "selected"%>>加粗</option></select> <label for="selectcl" style="vertical-align:middle;">文字颜色:</label> <select name="selectcl" id="selectcl" style="vertical-align:middle;"><option value="" <%if Fk_Article_CopyrightCl="" then response.write "selected"%>>默认</option>
					<option value="blue" class="blue" <%if Fk_Article_CopyrightCl="blue" then response.write "selected"%>>蓝色</option>
					<option value="red" class="red" <%if Fk_Article_CopyrightCl="red" then response.write "selected"%>>红色</option>
					<option value="green" class="green" <%if Fk_Article_CopyrightCl="green" then response.write "selected"%>>绿色</option>
					<option value="yellow" class="yellow" <%if Fk_Article_CopyrightCl="yellow" then response.write "selected"%>>黄色</option>
					<option value="gray" class="gray" <%if Fk_Article_CopyrightCl="gray" then response.write "selected"%>>灰色</option>
					<option value="black" class="black" <%if Fk_Article_CopyrightCl="black" then response.write "selected"%>>黑色</option>
					<option value="Purple" class="Purple" <%if Fk_Article_CopyrightCl="Purple" then response.write "selected"%>>紫色</option></select><div></td>
    </tr>
       
</table>
</div>
<div id="BoxBottom" style="width:96%;">
		<input type="hidden" name="Id" value="<%=Id%>" />
        <input type="hidden" name="ModuleId" value="<%=Fk_Module_Id%>" />
        <input type="submit" onclick="Sends('ArticleEdit','Article.asp?Type=5',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" class="Button" name="button" id="button" value="修 改" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="btnClose" id="btnClose" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：ArticleEditDo
'作    用：执行修改内容
'参    数：
'==============================
Sub ArticleEditDo()
	dim Fk_Article_syn,Syn_type
	Fk_Module_Id=Trim(Request.Form("ModuleId"))
	'判断权限
	If Not FkFun.CheckLimit("Module"&Fk_Module_Id) Then
		'Response.Write("无权限！")
		'Call FKDB.DB_Close()
		'Session.CodePage=936
		'Response.End()
	End If
	Fk_Article_Title=FKFun.HTMLEncode(Trim(Request.Form("Fk_Article_Title")))
	Fk_Article_Color=FKFun.HTMLEncode(Trim(Request.Form("Fk_Article_Color")))
	Fk_Article_Seotitle=FKFun.HTMLEncode(Trim(Request.Form("Fk_Article_Seotitle")))
	Fk_Article_Keyword=FKFun.HTMLEncode(Trim(Request.Form("Fk_Article_Keyword")))
	Fk_Article_Description=FKFun.HTMLEncode(Trim(Request.Form("Fk_Article_Description")))
	Fk_Article_Url=FKFun.HTMLEncode(Trim(Request.Form("Fk_Article_Url")))
	Fk_Article_From=FKFun.HTMLEncode(Trim(Request.Form("Fk_Article_From")))
	Fk_Article_Pic=FKFun.HTMLEncode(Trim(Request.Form("Fk_Article_Pic")))
	Fk_Article_PicBig=FKFun.HTMLEncode(Trim(Request.Form("Fk_Article_PicBig")))
	Fk_Article_Content=Request.Form("Fk_Article_Content")
	Fk_Article_FileName=FKFun.HTMLEncode(Trim(Request.Form("Fk_Article_FileName")))
	Fk_Article_Recommend=","&FKFun.HTMLEncode(Replace(Trim(Request.Form("Fk_Article_Recommend"))," ",""))&","
	Fk_Article_Subject=","&FKFun.HTMLEncode(Replace(Trim(Request.Form("Fk_Article_Subject"))," ",""))&","
	Fk_Article_Show=Trim(Request.Form("Fk_Article_Show"))
	Fk_Article_Template=Trim(Request.Form("Fk_Article_Template"))
	Fk_Article_Time=Trim(Request.Form("Fk_Article_Time"))
	Fk_Article_onTop=Trim(Request.Form("Fk_Article_onTop"))
	Fk_Article_px=Trim(Request.Form("Fk_Article_px"))
	Fk_Article_click=Trim(Request.Form("Fk_Article_click"))
	Fk_Article_Copyright=Request.Form("Fk_Article_Copyright")
	Fk_Article_CopyrightInfo=Trim(Request.Form("Fk_Article_CopyrightInfo"))
	Fk_Article_CopyrightFs=Trim(Request.Form("selectfs"))
	Fk_Article_CopyrightFt=Trim(Request.Form("selectwt"))
	Fk_Article_CopyrightCl=Trim(Request.Form("selectcl"))
	'Fk_Article_syn=Trim(Request.Form("Fk_Article_syn"))
	'Syn_type=Trim(Request.Form("Syn_type"))
	If Fk_Article_Time="" Then
		Fk_Article_Time=0
	End If
	Id=Trim(Request.Form("Id"))
	Call FKFun.ShowString(Fk_Article_Title,1,255,0,"请输入内容标题！","内容标题不能大于255个字符！")
	Call FKFun.ShowNum(Fk_Article_px,"排序只能是数字！请重新填写")
	Call FKFun.ShowString(Fk_Article_Keyword,0,255,2,"请输入内容SEO关键词！","内容SEO关键词不能大于255个字符！")
	Call FKFun.ShowString(Fk_Article_Description,0,255,2,"请输入内容SEO描述！","内容SEO描述不能大于255个字符！")
	Call FKFun.ShowString(Fk_Article_From,1,50,2,"请输入内容来源！","内容来源不能大于50个字符！")
	Call FKFun.ShowString(Fk_Article_Url,0,255,2,"请输入内容转向链接！","内容转向链接不能大于255个字符！")
	If Fk_Article_Url="" Then
		Call FKFun.ShowString(Fk_Article_Content,10,1,1,"请输入内容内容，不少于10个字符！","内容内容不能大于1个字符！")
	End If
	Call FKFun.ShowString(Fk_Article_Pic,0,255,2,"请输入内容题图路径！","内容题图小图路径不能大于255个字符！")
	Call FKFun.ShowString(Fk_Article_PicBig,0,255,2,"请输入内容题图路径！","内容题图大图路径不能大于255个字符！")
	Call FKFun.ShowNum(Fk_Module_Id,"ModuleId系统参数错误，请刷新页面！")
	Call FKFun.ShowString(Fk_Article_FileName,0,100,2,"生成文件名不符合标准！","生成文件名不能大于100个字符！")
	'Call FKFun.ShowString(Fk_Article_FileName,2,1,1,"生成文件名不能为空","生成内容不能大于1个字符！")
	Call FKFun.ShowNum(Fk_Article_Template,"请选择模板！")
	Call FKFun.ShowNum(Fk_Article_Show,"请选择内容是否显示！")
	Call FKFun.ShowNum(Fk_Article_click,"请输入点正确的击量！")
	Call FKFun.ShowNum(Fk_Article_Copyright,"参数错误，请刷新页面！")
	Call FKFun.ShowString(Fk_Article_CopyrightInfo,0,200,0,"请输入转载声明！","转载声明内容不能大于200个字符！")
	Call FKFun.ShowString(Fk_Article_CopyrightFs,0,50,0,"请选择字体大小！","字体大小不能大于50个字符！")
	Call FKFun.ShowString(Fk_Article_CopyrightFt,0,50,0,"请选择是否加粗！","粗体样式不能大于50个字符！")
	Call FKFun.ShowString(Fk_Article_CopyrightCl,0,50,0,"请选择字体颜色！","字体颜色不能大于50个字符！")
	
	Call FKFun.ShowNum(Id,"Id系统参数错误，请刷新页面！")
	'if Fk_Article_syn="1" then
		'Call FKFun.ShowNum(Syn_type,"请选择要同步到的行业类型")
	'end if
	Sqlstr="Select * From [Fk_Field] Where Fk_Field_Type=0 Order By Fk_Field_Id Asc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
		If Fk_Article_Field="" Then
			Fk_Article_Field=Rs("Fk_Field_Tag")&"|-Fangka_Field-|"&FKFun.HTMLEncode(Trim(Request.Form("Fk_Article__"&Rs("Fk_Field_Tag"))))
		Else
			Fk_Article_Field=Fk_Article_Field&"[-Fangka_Field-]"&Rs("Fk_Field_Tag")&"|-Fangka_Field-|"&FKFun.HTMLEncode(Trim(Request.Form("Fk_Article__"&Rs("Fk_Field_Tag"))))
		End If
		Rs.MoveNext
	Wend
	Rs.Close
	If Left(Fk_Article_FileName,4)="Info" Or Left(Fk_Article_FileName,4)="Page" Or Left(Fk_Article_FileName,5)="GBook" Or Left(Fk_Article_FileName,3)="Job" Then
		Response.Write("文件名受限，不能以一下单词开头：Info、Page、GBook、Job！")
		Call FKDB.DB_Close()
		Response.End()
	End If
	If IsNumeric(Fk_Article_FileName) Then
		Response.Write("文件名不可用纯数字！")
		Call FKDB.DB_Close()
		Response.End()
	End If
	If SiteDelWord=1 Then
		TempArr=Split(Trim(FKFun.UnEscape(FKFso.FsoFileRead("DelWord.dat")))," ")
		For Each Temp In TempArr
			If Temp<>"" Then
				Fk_Article_Content=Replace(Fk_Article_Content,Temp,"**")
				Fk_Article_Title=Replace(Fk_Article_Title,Temp,"**")
				Fk_Article_Keyword=Replace(Fk_Article_Keyword,Temp,"**")
				Fk_Article_Description=Replace(Fk_Article_Description,Temp,"**")
			End If
		Next
	End If
	
	dim Fk_Article_Copyright,mrs,Fk_Module_Dir,ArticleUrl
	if Fk_Article_Url="" then
		set mrs=conn.execute("select Fk_Module_Dir from Fk_Module where Fk_Module_Id="&Fk_Module_Id)
		if not mrs.eof then
			Fk_Module_Dir=mrs("Fk_Module_Dir")
		end if
		mrs.close
		set mrs=nothing
		If Fk_Module_Dir<>"" Then
			ArticleUrl=Fk_Module_Dir&"/"
		Else
			ArticleUrl="Article"&Fk_Module_Id&"/"
		End If
		If Fk_Article_FileName<>"" Then
			ArticleUrl=ArticleUrl&Fk_Article_FileName&".html"
		Else
			ArticleUrl=ArticleUrl&Fk_Module_Id&".html"
		End If
		If SiteHtml=1 Then
			ArticleUrl="/html"&SiteDir&ArticleUrl
		Else
			ArticleUrl=SiteDir&sTemp&"?"&ArticleUrl
		End If
	end if
	
	'新功能，追加转载声明
	'2014年12月31日
	'middy241@163.com
	if CheckFields("Fk_Article_Copyright","Fk_Article")=false then
		conn.execute("alter table Fk_Article add column Fk_Article_Copyright int default 0")
		conn.execute("alter table Fk_Article add column Fk_Article_CopyrightInfo varchar(200) null")
		conn.execute("alter table Fk_Article add column Fk_Article_CopyrightFs varchar(50) null")
		conn.execute("alter table Fk_Article add column Fk_Article_CopyrightFt varchar(50) null")
		conn.execute("alter table Fk_Article add column Fk_Article_CopyrightCl varchar(50) null")
	end if
	
	Sqlstr="Select * From [Fk_Article] Where Fk_Article_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs("Fk_Article_Title")=Fk_Article_Title
		Rs("Fk_Article_Color")=Fk_Article_Color
		Rs("Fk_Article_From")=Fk_Article_From
		Rs("Fk_Article_Seotitle")=Fk_Article_Seotitle
		Rs("Fk_Article_Keyword")=Fk_Article_Keyword
		Rs("Fk_Article_Field")=Fk_Article_Field
		Rs("Fk_Article_Description")=Fk_Article_Description
		Rs("Fk_Article_Pic")=Fk_Article_Pic
		Rs("Fk_Article_PicBig")=Fk_Article_PicBig
		Rs("Fk_Article_Show")=Fk_Article_Show
		Rs("Fk_Article_click")=Fk_Article_click
		Rs("Fk_Article_Url")=Fk_Article_Url
		Rs("Fk_Article_Recommend")=Fk_Article_Recommend
		Rs("Fk_Article_Subject")=Fk_Article_Subject
		Rs("Fk_Article_Content")=Fk_Article_Content
		Rs("Fk_Article_FileName")=Fk_Article_FileName
		Rs("Fk_Article_Template")=Fk_Article_Template
		Rs("Fk_Article_Ip")=Fk_Article_onTop
		
		Rs("Fk_Article_Copyright")=Fk_Article_Copyright
		Rs("Fk_Article_CopyrightInfo")=Fk_Article_CopyrightInfo
		Rs("Fk_Article_CopyrightFs")=Fk_Article_CopyrightFs
		Rs("Fk_Article_CopyrightFt")=Fk_Article_CopyrightFt
		Rs("Fk_Article_CopyrightCl")=Fk_Article_CopyrightCl
		Rs("Px")=Fk_Article_px
		If Fk_Article_Time=1 Then
			Rs("Fk_Article_Time")=Now()
		End If
		Rs.Update()
		Application.UnLock()
		
		'插入日志
		on error resume next
		dim log_content,log_ip,log_user
		log_content="修改内容：【"&Fk_Article_Title&"】"
		log_user=Request.Cookies("FkAdminName")
		
		log_ip=FKFun.getIP()
		conn.execute("insert into newTB_log (log_user,log_content,log_ip) values('"&log_user&"','"&log_content&"','"&log_ip&"')")
		
		Response.Write("“"&Fk_Article_Title&"”修改成功！")
	Else
		Response.Write("内容不存在！")
	End If
	Rs.Close
	If SiteHtml=1 And Fk_Article_Show=1 And Fk_Article_Url="" Then
		Dim FKHTML
		Set FKHTML=New Cls_HTML
		Sqlstr="Select * From [Fk_ArticleList] Where Fk_Article_Id=" & Id
		Rs.Open Sqlstr,Conn,1,1
		Fk_Module_Dir=Rs("Fk_Module_Dir")
		Fk_Article_Module=Rs("Fk_Article_Module")
		Rs.Close
		Call FKHTML.CreatArticle(Fk_Article_Template,Fk_Article_Module,Fk_Module_Dir,Fk_Article_FileName,Fk_Article_Title,1)
	Else
		Sqlstr="Select * From [Fk_ArticleList] Where Fk_Article_Id=" & Id
		Rs.Open Sqlstr,Conn,1,1
		Fk_Module_Dir=Rs("Fk_Module_Dir")
		Fk_Article_Module=Rs("Fk_Article_Module")
		Rs.Close
		If Fk_Module_Dir<>"" Then
			Temp="../"&Fk_Module_Dir&"/"
		Else
			Temp="../Article"&Fk_Article_Module&"/"
		End If
		If Fk_Article_FileName<>"" Then
			Temp=Temp&Fk_Article_FileName&".html"
		Else
			Temp=Temp&Id&".html"
		End If
		'Call FKFso.DelFile(Temp)
	End If
	'if Fk_Article_syn="1" then
'			dim reqHandler,key
'			key 	= "85e5ffb11e1c4a8561b953a7e27a547c"
'			set reqHandler = new SyncRequestHandler
'			'初始化
'			reqHandler.init()
'			'设置密钥
'			reqHandler.setKey(key)
'			'-----------------------------
'			'设置同步参数
'			'-----------------------------
'			reqHandler.setParameter "tit", Fk_Article_Title		'标题
'			reqHandler.setParameter "con",FKFun.RemoveHTML(Fk_Article_Content)		'内容
'			reqHandler.setParameter "typ", "-1"		'类型
'
'			'请求的参数
'			Dim Para,return,SyncUrl,host
'			host=request.ServerVariables("HTTP_HOST")
'			reqHandler.setParameter "hos", host		'域名
'			Para  	= reqHandler.getParameters()
'			SyncUrl	="http://qbknow.qb02.com/json/sync_article.asp"
'			return	= reqHandler.PostHttpPage("qbknow.qb02.com",SyncUrl,Para)
'			Response.Write(return)
	'end if
End Sub
Function RegExpReplace(patrn, strng)
dim RetStr
Dim regEx, Match, Matches ' 建立变量。
Set regEx = New RegExp ' 建立正则表达式。
regEx.Pattern = patrn ' 设置模式。
regEx.IgnoreCase = True ' 设置是否区分字符大小写。
regEx.Global = True ' 设置全局可用性。
Set Matches = regEx.Execute(strng) ' 执行搜索。
For Each Match in Matches ' 遍历匹配集合。
' RetStr = RetStr & "Match found at position "
' RetStr = RetStr & Match.FirstIndex & ". Match Value is '"
'RetStr = RetStr & Match.Value & "'." & vbCRLF
strng =  replace(strng,Match.Value,"")
Next
RegExpReplace = strng
End Function 
'==============================
'函 数 名：ArticleDelDo
'作    用：执行删除内容
'参    数：
'==============================
Sub ArticleDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"Id系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_ArticleList] Where Fk_Article_Id=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		'判断权限
		If Not FkFun.CheckLimit("Module"&Rs("Fk_Article_Module")) Then
			'Response.Write("无权限！")
			'Call FKDB.DB_Close()
			'Session.CodePage=936
			'Response.End()
		End If
		Fk_Article_Title=Rs("Fk_Article_Title")
		Fk_Module_Dir=Rs("Fk_Module_Dir")
		Fk_Article_Module=Rs("Fk_Article_Module")
		Fk_Article_FileName=Rs("Fk_Article_FileName")
		Rs.Close
		If Fk_Module_Dir<>"" Then
			Temp="../"&Fk_Module_Dir&"/"
		Else
			Temp="../Article"&Fk_Article_Module&"/"
		End If
		If Fk_Article_FileName<>"" Then
			Temp=Temp&Fk_Article_FileName&".html"
		Else
			Temp=Temp&Id&".html"
		End If
		'Call FKFso.DelFile(Temp)
	Else
		Rs.Close
		Response.Write("内容不存在！")
		Call FKDB.DB_Close()
		Session.CodePage=936
		Response.End()
	End If
	Sqlstr="Select * From [Fk_Article] Where Fk_Article_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs.Delete()
		Application.UnLock()		
		Response.Write("内容删除成功！")
		'插入日志
		on error resume next
		dim log_content,log_ip,log_user
		log_content="删除内容：【"&Fk_Article_Title&"】"
		log_user=Request.Cookies("FkAdminName")
		
		log_ip=FKFun.getIP()
		conn.execute("insert into newTB_log (log_user,log_content,log_ip) values('"&log_user&"','"&log_content&"','"&log_ip&"')")
	Else
		Response.Write("内容不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：ListDelDo
'作    用：执行批量删除内容
'参    数：
'==============================
Sub ListDelDo()
	Id=Replace(Trim(Request.Form("ListId"))," ","")
	If Id="" Then
		Response.Write("请选择要删除的内容！")
		Call FKDB.DB_Close()
		Response.End()
	End If
	Sqlstr="select Fk_Article_Title From [Fk_Article] Where Fk_Article_Id In ("&Id&")"
	Rs.Open Sqlstr,Conn,1,3
	if not rs.eof then
		i=0
		do while not rs.eof
		if i=0 then
			Fk_Article_Title="【"&rs("Fk_Article_Title")&"】"
		else
			Fk_Article_Title=Fk_Article_Title&","&"【"&rs("Fk_Article_Title")&"】"
		end if
		i=i+1
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		rs.movenext
		loop
	end if
	Response.Write("内容批量删除成功！")
	'插入日志
	on error resume next
	dim log_content,log_ip,log_user
	log_content="批量删除内容："&Fk_Article_Title&""
	log_user=Request.Cookies("FkAdminName")
	
	log_ip=FKFun.getIP()
	conn.execute("insert into newTB_log (log_user,log_content,log_ip) values('"&log_user&"','"&log_content&"','"&log_ip&"')")
		
End Sub

'==============================
'函 数 名：ArticleMove
'作    用：执行批量移动内容
'参    数：
'==============================
Sub ArticleMove()
	Dim Fk_Module_Type
	Id=Replace(Trim(Request.QueryString("ListId"))," ","")
	Fk_Module_Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Fk_Module_Id,"请选择转移到的栏目！")
	If Id="" Then
		Response.Write("请选择要移动的内容！")
		Call FKDB.DB_Close()
		Response.End()
	End If
	Fk_Module_Type=1000
	Sqlstr="Select Fk_Module_Type From [Fk_Module] Where Fk_Module_Id="&Fk_Module_Id&""
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Module_Type=Rs("Fk_Module_Type")
	End If
	Rs.Close
	If Fk_Module_Type=1000 Then
		Response.Write("要移到的栏目不存在！")
		Call FKDB.DB_Close()
		Response.End()
	End If
	If Fk_Module_Type<>1 Then
		Response.Write("只能移动到内容栏目！")
		Call FKDB.DB_Close()
		Response.End()
	End If
	Sqlstr="Update [Fk_Article] Set Fk_Article_Module="&Fk_Module_Id&" Where Fk_Article_Id In ("&Id&")"
	Application.Lock()
	Conn.Execute(Sqlstr)
	Application.UnLock()
	Response.Write("内容批量移动成功！")
End Sub
%><!--#Include File="../Code.asp"-->