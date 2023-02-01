<!--#Include File="AdminCheck.asp"-->
<!--#Include File="../Class/Cls_HTML.asp"--><%
'==========================================
'文 件 名：Down.asp
'文件用途：下载管理拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'定义页面变量
Dim Fk_Down_Title,Fk_Down_Content,Fk_Down_Click,Fk_Down_Show,Fk_Down_Time,Fk_Down_Pic,Fk_Down_PicBig,Fk_Down_Template,Fk_Down_FileName,Fk_Down_Recommend,Fk_Down_Subject,Fk_Down_Keyword,Fk_Down_Description,Fk_Down_Color,Fk_Down_System,Fk_Down_Language,Fk_Down_File,Fk_Down_Url,Fk_Down_Field,Fk_Down_onTop,Fk_Down_px
Dim Fk_Module_Name,Fk_Module_Id,Fk_Module_Menu,Fk_Module_Dir,Fk_Down_Module
Dim Temp2,KeyWordlist,kwdrs,ki

On Error Resume next
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
		Call DownList() '下载列表
	Case 2
		Call DownAddForm() '添加下载表单
	Case 3
		Call DownAddDo() '执行添加下载
	Case 4
		Call DownEditForm() '修改下载表单
	Case 5
		Call DownEditDo() '执行修改下载
	Case 6
		Call DownDelDo() '执行删除下载
	Case 7
		Call ListDelDo() '执行批量删除下载
	Case 8
		Call DownMove() '执行批量移动下载
	Case Else
		Response.Write("没有找到此功能项！")
End Select

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

'==========================================
'函 数 名：DownList()
'作    用：下载列表
'参    数：
'==========================================
Sub DownList()
	'新功能，追加SEO title字段
	'2017年5月22日
	'middy241@163.com
	if CheckFields("Fk_Down_seotitle","Fk_Down")=false then
		conn.execute("alter table Fk_Down add column Fk_Down_seotitle varchar(255) null")
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
	End If
	Rs.Close
%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="ShowBox('Down.asp?Type=2&ModuleId=<%=Fk_Module_Id%>');">添加</a></li>
        <li><a href="javascript:void(0);" onclick="SetRContent('MainRight','<%=Session("NowPage")%>');">刷新</a></li>
    </ul>
</div>
<div id="ListTop">
    <%=Fk_Module_Name%>模块&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name="SearchStr" value="<%=SearchStr%>" type="text" class="Input" id="SearchStr" style="vertical-align:middle;"/>&nbsp;<input type="button" class="Button" onclick="SetRContent('MainRight','Down.asp?Type=1&ModuleId=<%=Fk_Module_Id%>&SearchStr='+escape(document.all.SearchStr.value));" name="S" Id="S" value="  查询  "  style="vertical-align:middle;"/>&nbsp;&nbsp;请选择模块：
<select name="D1" id="D1" onChange="window.execScript(this.options[this.selectedIndex].value);" style="vertical-align:middle;">
      <option value="alert('请选择模块');">请选择模块</option>
<%
Call ModuleSelectUrl(Fk_Module_Menu,0,Fk_Module_Id)
%>
</select>
</div>
<div id="ListContent">
    <form name="DelList" id="DelList" method="post" action="Down.asp?Type=7" onsubmit="return false;">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr>
            <td align="center" class="ListTdTop">选</td>
            <td align="center" class="ListTdTop">下载名称</td>
            <td align="center" style="display:none" class="ListTdTop">文件名</td>
            <td align="center" class="ListTdTop">下载参数</td>
            <td align="center" class="ListTdTop">点击量</td>
            <td align="center" class="ListTdTop">排序</td>
            <td align="center" class="ListTdTop">添加时间</td>
            <td align="center" class="ListTdTop">操作</td>
        </tr>
<%
	Dim Rs2
	Set Rs2=Server.Createobject("Adodb.RecordSet")
	Sqlstr="Select * From [Fk_Down] Where Fk_Down_Module="&Fk_Module_Id&""
	If SearchStr<>"" Then
		Sqlstr=Sqlstr&" And Fk_Down_Title Like '%%"&SearchStr&"%%'"
	End If
	Sqlstr=Sqlstr&" Order By Fk_Down_Id Desc"
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Dim DownTemplate
		Rs.PageSize=PageSizes
		If PageNow>Rs.PageCount Or PageNow<=0 Then
			PageNow=1
		End If
		PageCounts=Rs.PageCount
		Rs.AbsolutePage=PageNow
		PageAll=Rs.RecordCount
		i=1
		While (Not Rs.Eof) And i<PageSizes+1
			If Rs("Fk_Down_Template")>0 Then
				Sqlstr="Select * From [Fk_Template] Where Fk_Template_Id=" & Rs("Fk_Down_Template")
				Rs2.Open Sqlstr,Conn,1,1
				If Not Rs2.Eof Then
					Fk_Down_Template=Rs2("Fk_Template_Name")
				Else
					Fk_Down_Template="未知模板"
				End If
				Rs2.Close
			Else
				Fk_Down_Template="默认模板"
			End If
			Fk_Down_Recommend=""
			If Rs("Fk_Down_Recommend")<>"" Then
				TempArr=Split(Rs("Fk_Down_Recommend"),",")
				For Each Temp In TempArr
					If Temp<>"" Then
						Sqlstr="Select * From [Fk_Recommend] Where Fk_Recommend_Id=" & Temp
						Rs2.Open Sqlstr,Conn,1,1
						If Not Rs2.Eof Then
							Fk_Down_Recommend=Fk_Down_Recommend&","&Rs2("Fk_Recommend_Name")
						End If
						Rs2.Close
					End If
				Next
			End If
			Fk_Down_Subject=""
			If Rs("Fk_Down_Subject")<>"" Then
				TempArr=Split(Rs("Fk_Down_Subject"),",")
				For Each Temp In TempArr
					If Temp<>"" Then
						Sqlstr="Select * From [Fk_Subject] Where Fk_Subject_Id=" & Temp
						Rs2.Open Sqlstr,Conn,1,1
						If Not Rs2.Eof Then
							Fk_Down_Subject=Fk_Down_Subject&","&Rs2("Fk_Subject_Name")
						End If
						Rs2.Close
					End If
				Next
			End If
%>
        <tr>
            <td height="20" align="center"><input type="checkbox" name="ListId" class="Checks" value="<%=Rs("Fk_Down_Id")%>" id="List<%=Rs("Fk_Down_Id")%>" /></td>
            <td>&nbsp;<%=Rs("Fk_Down_Title")%><%If Rs("Fk_Down_Color")<>"" Then%><span style="color:<%=Rs("Fk_Down_Color")%>">■</span><%End If%><%If Rs("Fk_Down_Url")<>"" Then%>[转向链接]<%End If%></td>
            <td align="center" style="display:none"><%=Rs("Fk_Down_FileName")%></td>
            <td align="center"><%If Rs("Fk_Down_Show")=1 Then%><span style="color:#000">[显]</span><%Else%><span style="color:#CCC">[隐]</span><%End If%><%If Rs("Fk_Down_Pic")<>"" Then%><span style="color:#F00">[图]</span><%End If%><a href="javascript:void(0);" title="<%=Fk_Down_Template%> ">[模]</a><%If InStr(Fk_Down_Recommend,"推荐")>0 Then%><a href="javascript:void(0);" title="<%=Replace(Fk_Down_Recommend,",","")%> ">[推]</a><%End If%><%If trim(Rs("Fk_Down_Ip")&" ")="1" Then%><a href="javascript:void(0);" title="置顶 "><font color=blue style="font-weight:bolder;">[顶]</font></a><%End If%><%If Fk_Down_Subject<>"" Then%><a href="javascript:void(0);" title="<%=Fk_Down_Subject%> ">[专]</a><%End If%></td>
            <td align="center"><%=Rs("Fk_Down_Click")%></td>
            <td height="20" align="center"><%=Rs("Px")%></td>
            <td align="center"><%=Rs("Fk_Down_Time")%></td>
            <td align="center"><a href="javascript:void(0);" onclick="ShowBox('Down.asp?Type=4&ModuleId=<%=Fk_Module_Id%>&Id=<%=Rs("Fk_Down_Id")%>');"><img src="images/edit.png"></a> <a href="javascript:void(0);" onclick="DelIt('您确认要删除“<%=Rs("Fk_Down_Title")%>”，此操作不可逆！','Down.asp?Type=6&Id=<%=Rs("Fk_Down_Id")%>','MainRight','<%=Session("NowPage")%>');"><img src="images/del.png"></a></td>
        </tr>
<%
			Rs.MoveNext
			i=i+1
		Wend
		
%>
        <tr>
            <td height="30" colspan="8">
            <input name="chkall" type="checkbox" id="chkall" value="select" onClick="CheckAll(this.form)"> 全选
            <input type="submit" value="删 除" class="Button" onClick="if(confirm('此操作无法恢复！！！请慎重！！！\n\n确定要删除选中的下载吗？')){Sends('DelList','Down.asp?Type=7',0,'',0,1,'MainRight','<%=Session("NowPage")%>');}">
<select name="DownMove" id="DownMove" onchange="DelIt('确实要移动这部分下载？','Down.asp?Type=8&Id='+this.options[this.options.selectedIndex].value+'&ListId='+GetCheckbox(),'MainRight','<%=Session("NowPage")%>');">
      <option value="">转移到</option>
<%
Call ModuleSelectId(Fk_Module_Menu,0,0)
%>
</select>
            &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<%Call FKFun.ShowPageCode("Down.asp?Type=1&ModuleId="&Fk_Module_Id&"&SearchStr="&Server.URLEncode(SearchStr)&"&Page=",PageNow,PageAll,PageSizes,PageCounts)%></td>
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
'函 数 名：DownAddForm()
'作    用：添加下载表单
'参    数：
'==========================================
Sub DownAddForm()
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
%>
<script type="text/javascript" charset="utf-8" id="colorpickerjs" src="/js/colorpicker.js"></script>
<script type="text/javascript">
$(document).ready(function(){
	$("#colorpickerjs").attr("src","/js/colorpicker.js?r="+Math.random());
	set_title_color("");
})
	function set_title_color(color) {
		$('#Fk_Down_Title').css('color',color);
		$('#Fk_Down_Color').val(color);
	}
	
	
	
	if(window.KindEditor){
		$("#Fk_Down_File").after(" <input type=\"button\" id=\"uploadButton\" value=\"上传\" class=\"Button\" style=\"vertical-align:middle;\"/> *上传限于rar/zip/doc/xls/ppt格式,文件最大允许5M");
			var editor = window.KindEditor.editor({
					fileManagerJson : '/admin/dkidtioenr/aps/file_manager_json.asp',
					uploadJson		: '/admin/dkidtioenr/aps/upload_json.asp?dir=file',
					allowFileManager : true
				});
				$('#uploadButton').click(function() {
					editor.loadPlugin('insertfile', function() {
						editor.plugin.fileDialog({
							fileUrl : $('#Fk_Down_File').val(),
							clickFn : function(url) {
								$('#Fk_Down_File').val(url);
								editor.hideDialog();
							}
						});
					});
				});

		}
		else
		{
			$("#Fk_Down_File").after(" <iframe frameborder=\"0\" width=\"200\" height=\"25\" scrolling=\"No\" id=\"Fk_Down_Files\" name=\"Fk_Down_Files\" src=\"PicUpLoad.asp?Form=DownEdit&Input=Fk_Down_File\" style=\"vertical-align:middle\"></iframe> *上传限于rar/zip/doc/xls/ppt格式,文件最大允许200K");
		}
</script>
<form id="DownAdd" name="DownAdd" method="post" action="Down.asp?Type=3" onsubmit="return false;">
<div id="BoxTop" style="width:98%;"><span>添加</span><a onclick="$('#Boxs').hide();$('select').show();"><img src="images/close3.gif"></a></div>
<div id="BoxContents" style="width:98%;">
	<table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td height="28" align="right">下载名称：</td>
        <td colspan="4">&nbsp;<input name="Fk_Down_Title"<%If SiteToPinyin=1 Then%> onchange="GetPinyin('Fk_Down_FileName','ToPinyin.asp?Str='+this.value);"<%End If%> type="text" class="Input" id="Fk_Down_Title" size="50"  style="vertical-align:middle;"/> <input name="Fk_Down_Color" type="hidden" id="Fk_Down_Color" value=""/><span class="colorpicker" onclick="colorpicker('colorpanel_title','set_title_color');" title="标题颜色" style="vertical-align:middle;background: url(http://image001.dg-cloud-01.qebang.cn/website/ext/images/icon/color.png) 0px 0px no-repeat;height: 24px;width: 24px;display: inline-block;"></span>
                           <span class="colorpanel" id="colorpanel_title" style="position:absolute;z-index:99999999;"></span><script type="text/javascript">
						   set_title_color("");
                           </script> &nbsp; &nbsp; 排序：<input type="text" value="0" id="Fk_Down_px" name="Fk_Down_px" class="Input"  size="4" maxlength="6" style="vertical-align:middle;"/>（限数字，越大越前）</td>
    </tr>
    <tr>
        <td height="28" align="right">关键字：</td>
        <td>&nbsp;<input name="Fk_Down_Keyword" type="text" class="Input" id="Fk_Down_Keyword" size="30"  style="vertical-align:middle;"/> <input type="submit" onclick="tiqu(0,'Fk_Down_Content','Fk_Down_Keyword');" class="Button" name="btntqkwd" id="btntqkwd" value="提 取" style="vertical-align:middle;" /><input  value="<%=KeyWordlist%>" type="hidden" id="Fk_Keywordlist" /></td>
        <td align="right" colspan="2">描述：</td>
        <td>&nbsp;<input name="Fk_Down_Description" type="text" class="Input" id="Fk_Down_Description" size="30" style="vertical-align:middle;" /> <input type="submit" onclick="tiqu(1,'Fk_Down_Content','Fk_Down_Description');" class="Button" name="btntqdesc" id="btntqdesc" value="提 取"  style="vertical-align:middle;"/></td>
    </tr>
    <tr>
        <td height="28" align="right">适用系统：</td>
        <td>&nbsp;<input name="Fk_Down_System" value="Windows2000/xp/vista/7" type="text" class="Input" id="Fk_Down_System" size="30" style="vertical-align:middle;" /></td>
        <td align="right" colspan="2">语言：</td>
        <td>&nbsp;<input name="Fk_Down_Language" value="简体中文" type="text" class="Input" id="Fk_Down_Language" size="30"  style="vertical-align:middle;"/></td>
    </tr>
    <tr style="display:none">
        <td height="28" align="right">题图：</td>
        <td colspan="2">&nbsp;缩略小图：<input name="Fk_Down_Pic" type="text" class="Input" id="Fk_Down_Pic" size="50" /><br />
        &nbsp;正常大图：<input name="Fk_Down_PicBig" type="text" class="Input" id="Fk_Down_PicBig" size="50" />&nbsp;</td>
        <td colspan="2">
		<iframe frameborder="0" width="290" height="25" scrolling="No" id="I2" name="I2" src="PicUpLoad.asp?Form=DownAdd&Input=Fk_Down_Pic"></iframe><br><font color="red">说明：只需上传一张图，自动生成缩略小图和正常大图。</font></td>
    </tr>
    <tr>
        <td height="28" align="right">文件名：</td>
        <td colspan="4">&nbsp;<input name="Fk_Down_FileName" type="text" class="Input" id="Fk_Down_FileName"  style="vertical-align:middle;"/>&nbsp;*一旦确立不可修改</td>
    </tr>
<%
	Sqlstr="Select * From [Fk_Field] Where Fk_Field_Type=2 Order By Fk_Field_Id Asc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
%>
    <tr>
        <td height="28" align="right"><%=Rs("Fk_Field_Name")%>：</td>
        <td colspan="4">&nbsp;<input name="Fk_Down__<%=Rs("Fk_Field_Tag")%>" type="text" class="Input" id="Fk_Down__<%=Rs("Fk_Field_Tag")%>"  style="vertical-align:middle;"/></td>
    </tr>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
    <tr>
        <td height="28" align="right">转向链接：</td>
        <td colspan="4">&nbsp;<input name="Fk_Down_Url" type="text" class="Input" id="Fk_Down_Url" size="50"  style="vertical-align:middle;"/>&nbsp;*正常下载请留空</td>
    </tr>
   
    
    <tr>
        <td height="28" align="right">下载文件：</td>
        <td colspan="4">&nbsp;<input name="Fk_Down_File" type="text" class="Input" id="Fk_Down_File" size="50"  style="vertical-align:middle;"/></td>
    </tr>
    <tr>
        <td height="28" align="right" width="100">下载内容：</td>
        <td colspan="4"><textarea name="Fk_Down_Content" class="<%=bianjiqi%>" id="Fk_Down_Content" rows="8" style="width:100%;"></textarea></td>
    </tr>
     <tr>
        <td height="28" align="right">模板：</td>
        <td colspan="4">&nbsp;<select name="Fk_Down_Template" class="Input" id="Fk_Down_Template" style="vertical-align:middle;">
            <option value="0"<%=FKFun.BeSelect(Fk_Down_Template,0)%>>默认模板</option>
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
            </select>　下载显示：<input name="Fk_Down_Show" type="radio" id="Fk_Down_Show" class="Input" value="1" checked=""  style="vertical-align:middle;"/><label for="Fk_Down_Show" style="vertical-align:middle;">显示</label>
        <input type="radio" name="Fk_Down_Show" class="Input" id="Fk_Down_Show1" value="0"  style="vertical-align:middle;"/><label for="Fk_Down_Show1" style="vertical-align:middle;">不显示</label>　推荐：<select name="Fk_Down_Recommend" class="Input" size="1" id="Fk_Down_Recommend" style="vertical-align:middle;">
            <option value="0">无推荐</option>
            <option value="2">推荐</option>
            </select>
			<input type="checkbox" name="Fk_Down_onTop" id="Fk_Down_onTop" class="textarea" value="1"  style="vertical-align:middle;"/><label style="vertical-align:middle;" for="Fk_Down_onTop">置顶</label>
			　<select name="Fk_Down_Subject" class="TextArea" size="1" multiple="multiple" id="Fk_Down_Subject" style="display:none">
            <option value="0">无专题</option>
            </select> &nbsp; &nbsp; 下载量：<input name="Fk_Down_click" type="text" id="Fk_Down_click" class="Input" size="6" maxlength="6" value="<%=rnd_num%>" style="vertical-align:middle;"></td>
    </tr>
    </table>
</div>
<div id="BoxBottom" style="width:96%;">
		<input type="hidden" name="ModuleId" value="<%=Fk_Module_Id%>" />
        <input type="submit" onclick="Sends('DownAdd','Down.asp?Type=3',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" class="Button" name="button" id="button" value="添 加" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="btnclose" id="btnclose" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：DownAddDo
'作    用：执行添加下载
'参    数：
'==============================
Sub DownAddDo()
	Fk_Module_Id=Trim(Request.Form("ModuleId"))
	'判断权限
	'If Not FkFun.CheckLimit("Module"&Fk_Module_Id) Then
		'Response.Write("无权限！")
		'Call FKDB.DB_Close()
		'Session.CodePage=936
		'Response.End()
	'End If
	Fk_Down_Title=FKFun.HTMLEncode(Trim(Request.Form("Fk_Down_Title")))
	Fk_Down_Color=FKFun.HTMLEncode(Trim(Request.Form("Fk_Down_Color")))
	Fk_Down_Keyword=FKFun.HTMLEncode(Trim(Request.Form("Fk_Down_Keyword")))
	Fk_Down_Description=FKFun.HTMLEncode(Trim(Request.Form("Fk_Down_Description")))
	Fk_Down_System=FKFun.HTMLEncode(Trim(Request.Form("Fk_Down_System")))
	Fk_Down_Language=FKFun.HTMLEncode(Trim(Request.Form("Fk_Down_Language")))
	Fk_Down_Url=FKFun.HTMLEncode(Trim(Request.Form("Fk_Down_Url")))
	Fk_Down_File=FKFun.HTMLEncode(Trim(Request.Form("Fk_Down_File")))
	Fk_Down_Content=Request.Form("Fk_Down_Content")
	Fk_Down_Pic=FKFun.HTMLEncode(Trim(Request.Form("Fk_Down_Pic")))
	Fk_Down_PicBig=FKFun.HTMLEncode(Trim(Request.Form("Fk_Down_PicBig")))
	Fk_Down_Recommend=","&FKFun.HTMLEncode(Replace(Trim(Request.Form("Fk_Down_Recommend"))," ",""))&","
	Fk_Down_Subject=","&FKFun.HTMLEncode(Replace(Trim(Request.Form("Fk_Down_Subject"))," ",""))&","
	Fk_Down_FileName=FKFun.HTMLEncode(Trim(Request.Form("Fk_Down_FileName")))
	Fk_Down_Template=Trim(Request.Form("Fk_Down_Template"))
	Fk_Down_Show=Trim(Request.Form("Fk_Down_Show"))
	Fk_Down_click=Trim(Request.Form("Fk_Down_click"))
	Fk_Down_onTop=Trim(Request.Form("Fk_Down_onTop"))
	Fk_Down_px=Trim(Request.Form("Fk_Down_px"))
	Call FKFun.ShowString(Fk_Down_Title,1,255,0,"请输入下载标题！","下载标题不能大于255个字符！")
	Call FKFun.ShowNum(Fk_Down_px,"排序只能是数字！请重新填写")
	Call FKFun.ShowString(Fk_Down_Keyword,0,255,2,"请输入下载关键字！","下载关键字不能大于255个字符！")
	Call FKFun.ShowString(Fk_Down_Description,0,255,2,"请输入下载描述！","下载描述不能大于255个字符！")
	Call FKFun.ShowString(Fk_Down_System,0,255,2,"请输入适用系统！","适用系统不能大于255个字符！")
	Call FKFun.ShowString(Fk_Down_Language,0,255,2,"请输入语言版本！","语言版本不能大于255个字符！")
	Call FKFun.ShowString(Fk_Down_File,0,255,2,"请输入下载地址或上传文件！","下载地址不能大于255个字符！")
	Call FKFun.ShowString(Fk_Down_Url,0,255,2,"请输入下载转向链接！","下载转向链接不能大于255个字符！")
	If Fk_Down_Url="" Then
		Call FKFun.ShowString(Fk_Down_Content,1,1,1,"请输入下载简介！","下载内容不能大于1个字符！")
	End If
	Call FKFun.ShowString(Fk_Down_Pic,0,255,2,"请输入下载题图路径！","下载题图小图路径不能大于255个字符！")
	Call FKFun.ShowString(Fk_Down_PicBig,0,255,2,"请输入下载题图路径！","下载题图大图路径不能大于255个字符！")
	Call FKFun.ShowString(Fk_Down_FileName,0,50,2,"请输入下载文件名！","下载文件名不能大于50个字符！")
	Call FKFun.ShowNum(Fk_Down_Template,"请选择模板！")
	Call FKFun.ShowNum(Fk_Down_Show,"请选择下载是否显示！")
	Call FKFun.ShowNum(Fk_Down_click,"请输入正确的下载量！")
	Call FKFun.ShowNum(Fk_Module_Id,"ModuleId系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_Field] Where Fk_Field_Type=2 Order By Fk_Field_Id Asc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
		If Fk_Down_Field="" Then
			Fk_Down_Field=Rs("Fk_Field_Tag")&"|-Fangka_Field-|"&FKFun.HTMLEncode(Trim(Request.Form("Fk_Down__"&Rs("Fk_Field_Tag"))))
		Else
			Fk_Down_Field=Fk_Down_Field&"[-Fangka_Field-]"&Rs("Fk_Field_Tag")&"|-Fangka_Field-|"&FKFun.HTMLEncode(Trim(Request.Form("Fk_Down__"&Rs("Fk_Field_Tag"))))
		End If
		Rs.MoveNext
	Wend
	Rs.Close
	If IsNumeric(Fk_Down_FileName) Then
		Response.Write("文件名不可用纯数字！")
		Call FKDB.DB_Close()
		Response.End()
	End If
	If Left(Fk_Down_FileName,4)="Info" Or Left(Fk_Down_FileName,4)="Page" Or Left(Fk_Down_FileName,5)="GBook" Or Left(Fk_Down_FileName,3)="Job" Then
		Response.Write("文件名受限，不能以一下单词开头：Info、Page、GBook、Job！")
		Call FKDB.DB_Close()
		Response.End()
	End If
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id=" & Fk_Module_Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Module_Menu=Rs("Fk_Module_Menu")
	Else
		Response.Write("模块不存在！")
		Rs.Close
		Call FKDB.DB_Close()
		Response.End()
	End If
	Rs.Close
	If SiteDelWord=1 Then
		TempArr=Split(Trim(FKFun.UnEscape(FKFso.FsoFileRead("DelWord.dat")))," ")
		For Each Temp In TempArr
			If Temp<>"" Then
				Fk_Down_Content=Replace(Fk_Down_Content,Temp,"**")
				Fk_Down_Title=Replace(Fk_Down_Title,Temp,"**")
				Fk_Down_Keyword=Replace(Fk_Down_Keyword,Temp,"**")
				Fk_Down_Description=Replace(Fk_Down_Description,Temp,"**")
				Fk_Down_System=Replace(Fk_Down_System,Temp,"**")
				Fk_Down_Language=Replace(Fk_Down_Language,Temp,"**")
			End If
		Next
	End If
	Sqlstr="Select * From [Fk_Down] Where Fk_Down_Module="&Fk_Module_Id&" And (Fk_Down_Title='"&Fk_Down_Title&"'"
	If Fk_Down_FileName<>"" Then
		Sqlstr=Sqlstr&" Or Fk_Down_FileName='"&Fk_Down_FileName&"'"
	End If
	Sqlstr=Sqlstr&")"
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("Fk_Down_Title")=Fk_Down_Title
		Rs("Fk_Down_Color")=Fk_Down_Color
		Rs("Fk_Down_Keyword")=Fk_Down_Keyword
		Rs("Fk_Down_Description")=Fk_Down_Description
		Rs("Fk_Down_File")=Fk_Down_File
		Rs("Fk_Down_Url")=Fk_Down_Url
		Rs("Fk_Down_Field")=Fk_Down_Field
		Rs("Fk_Down_System")=Fk_Down_System
		Rs("Fk_Down_Language")=Fk_Down_Language
		Rs("Fk_Down_Show")=Fk_Down_Show
		Rs("Fk_Down_click")=Fk_Down_click
		Rs("Fk_Down_Pic")=Fk_Down_Pic
		Rs("Fk_Down_PicBig")=Fk_Down_PicBig
		Rs("Fk_Down_Content")=Fk_Down_Content
		Rs("Fk_Down_Recommend")=Fk_Down_Recommend
		Rs("Fk_Down_Subject")=Fk_Down_Subject
		Rs("Fk_Down_Module")=Fk_Module_Id
		Rs("Fk_Down_Menu")=Fk_Module_Menu
		Rs("Fk_Down_FileName")=Fk_Down_FileName
		Rs("Fk_Down_Template")=Fk_Down_Template
		Rs("Fk_Down_Ip")=Fk_Down_onTop
		Rs("Px")=Fk_Down_px
		Rs.Update()
		Application.UnLock()
		Response.Write("新下载添加成功！")
		'插入日志
		on error resume next
		dim log_content,log_ip,log_user
		log_content="添加下载：【"&Fk_Down_Title&"】"
		log_user=Request.Cookies("FkAdminName")
		
		log_ip=FKFun.getIP()
		conn.execute("insert into newTB_log (log_user,log_content,log_ip) values('"&log_user&"','"&log_content&"','"&log_ip&"')")
	Else
		Response.Write("该下载标题已经被占用，请重新选择！")
	End If
	Rs.Close
End Sub

'==========================================
'函 数 名：DownEditForm()
'作    用：修改下载表单
'参    数：
'==========================================
Sub DownEditForm()
	Id=Clng(Request.QueryString("Id"))
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
		Fk_Module_Id=Rs("Fk_Module_Id")
	End If
	Rs.Close
	Sqlstr="Select * From [Fk_Down] Where Fk_Down_Id=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		Fk_Down_Title=Rs("Fk_Down_Title")
		Fk_Down_Color=Rs("Fk_Down_Color")
		Fk_Down_Keyword=Rs("Fk_Down_Keyword")
		Fk_Down_Description=Rs("Fk_Down_Description")
		Fk_Down_Url=Rs("Fk_Down_Url")
		Fk_Down_File=Rs("Fk_Down_File")
		Fk_Down_System=Rs("Fk_Down_System")
		Fk_Down_Language=Rs("Fk_Down_Language")
		Fk_Down_Content=Rs("Fk_Down_Content")
		Fk_Down_Pic=Rs("Fk_Down_Pic")
		Fk_Down_PicBig=Rs("Fk_Down_PicBig")
		Fk_Down_Show=Rs("Fk_Down_Show")
		Fk_Down_click=Rs("Fk_Down_click")
		Fk_Down_Template=Rs("Fk_Down_Template")
		Fk_Down_FileName=Rs("Fk_Down_FileName")
		Fk_Down_Recommend=Rs("Fk_Down_Recommend")
		Fk_Down_Subject=Rs("Fk_Down_Subject")
		Fk_Down_onTop=trim(Rs("Fk_Down_Ip")&" ")
		Fk_Down_px=Rs("Px")
		If IsNull(Rs("Fk_Down_Field")) Or Rs("Fk_Down_Field")="" Then
			Fk_Down_Field=Split("-_-|-Fangka_Field-|1")
		Else
			Fk_Down_Field=Split(Rs("Fk_Down_Field"),"[-Fangka_Field-]")
		End If
	End If
	Rs.Close
%>
<script type="text/javascript" charset="utf-8" id="colorpickerjs" src="/js/colorpicker.js"></script>
<script type="text/javascript">
$(document).ready(function(){
	$("#colorpickerjs").attr("src","/js/colorpicker.js?r="+Math.random());
	set_title_color("");
})
	function set_title_color(color) {
		$('#Fk_Down_Title').css('color',color);
		$('#Fk_Down_Color').val(color);
	}

	if(window.KindEditor){
				$("#Fk_Down_File").after(" <input type=\"button\" id=\"uploadButton\" value=\"上传\" class=\"Button\" style=\"vertical-align:middle;\"/> *上传限于rar/zip/doc/xls/ppt格式,文件最大允许5M");
			var editor = window.KindEditor.editor({
					fileManagerJson : '/admin/dkidtioenr/aps/file_manager_json.asp',
					uploadJson		: '/admin/dkidtioenr/aps/upload_json.asp?dir=file',
					allowFileManager : true
				});
				$('#uploadButton').click(function() {
					editor.loadPlugin('insertfile', function() {
						editor.plugin.fileDialog({
							fileUrl : $('#Fk_Down_File').val(),
							clickFn : function(url) {
								$('#Fk_Down_File').val(url);
								editor.hideDialog();
							}
						});
					});
				});
		}
		else
		{
			$("#Fk_Down_File").after("<iframe frameborder=\"0\" width=\"200\" height=\"25\" scrolling=\"No\" id=\"Fk_Down_Files\" name=\"Fk_Down_Files\" src=\"PicUpLoad.asp?Form=DownEdit&Input=Fk_Down_File\" style=\"vertical-align:middle\"></iframe> *上传限于rar/zip/doc/xls/ppt格式,文件最大允许200K");
		}

</script>
<form id="DownEdit" name="DownEdit" method="post" action="Down.asp?Type=5" onsubmit="return false;">
<div id="BoxTop" style="width:98%;"><span>修改</span><a onclick="$('#Boxs').hide();$('select').show();"><img src="images/close3.gif"></a></div>
<div id="BoxContents" style="width:98%;">
<table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td height="28" align="right">名称：</td>
        <td colspan="4">&nbsp;<input name="Fk_Down_Title"<%If SiteToPinyin=1 And (Fk_Down_FileName="" Or IsNull(Fk_Down_FileName)) Then%> onchange="GetPinyin('Fk_Down_FileName','ToPinyin.asp?Str='+this.value);"<%End If%> value="<%=Fk_Down_Title%>" type="text" class="Input" id="Fk_Down_Title" size="50"  style="vertical-align:middle"/> <input name="Fk_Down_Color" type="hidden" id="Fk_Down_Color" value=""/><span class="colorpicker" onclick="colorpicker('colorpanel_title','set_title_color');" title="标题颜色" style="vertical-align:middle;background: url(http://image001.dgcloud01.qebang.cn/website/ext/images/icon/color.png) 0px 0px no-repeat;height: 24px;width: 24px;display: inline-block;"></span>
                           <span class="colorpanel" id="colorpanel_title" style="position:absolute;z-index:99999999;"></span><script type="text/javascript">
						   set_title_color("<%=Fk_Down_Color%>");
                           </script> &nbsp; &nbsp; 排序：<input type="text" name="Fk_Down_px" id="Fk_Down_px" class="Input" value="<%if isnull(Fk_Down_px) then response.write "0" else response.write Fk_Down_px%>" maxlength="6" size="4" style="vertical-align:middle"/>（限数字,越大越前）&nbsp; &nbsp; <input type="checkbox" name="Fk_Down_Time" id="Fk_Down_Time" value="1"  style="vertical-align:middle"/><label for="Fk_Down_Time" style="vertical-align:middle">更新时间</label></td>
    </tr>
    <tr>
        <td height="28" align="right">关键字：</td>
        <td>&nbsp;<input name="Fk_Down_Keyword" value="<%=Fk_Down_Keyword%>" type="text" class="Input" id="Fk_Down_Keyword" size="30"  style="vertical-align:middle"/> <input type="submit" onclick="tiqu(0,'Fk_Down_Content','Fk_Down_Keyword');" class="Button" name="btntqkwd" id="btntqkwd" value="提 取"  style="vertical-align:middle"/><input  value="<%=KeyWordlist%>" type="hidden" id="Fk_Keywordlist" /></td>
        <td align="right" colspan="2">描述：</td>
        <td>&nbsp;<input name="Fk_Down_Description" value="<%=Fk_Down_Description%>" type="text" class="Input" id="Fk_Down_Description" size="30"  style="vertical-align:middle"/> <input type="submit" onclick="tiqu(1,'Fk_Down_Content','Fk_Down_Description');" class="Button" name="btntqdesc" id="btntqdesc" value="提 取"  style="vertical-align:middle"/></td>
    </tr>
    <tr>
        <td height="28" align="right">适用系统：</td>
        <td>&nbsp;<input name="Fk_Down_System" value="<%=Fk_Down_System%>" type="text" class="Input" id="Fk_Down_System" size="30"  style="vertical-align:middle"/></td>
        <td align="right" colspan="2">语言：</td>
        <td>&nbsp;<input name="Fk_Down_Language" value="<%=Fk_Down_Language%>" type="text" class="Input" id="Fk_Down_Language" size="30"  style="vertical-align:middle"/></td>
    </tr>
    <tr style="display:none">
        <td height="28" align="right">题图：</td>
        <td colspan="2">&nbsp;缩略小图：<input name="Fk_Down_Pic" type="text" class="Input" id="Fk_Down_Pic" value="<%=Fk_Down_Pic%>" size="50" /><br />
        &nbsp;正常大图：<input name="Fk_Down_PicBig" type="text" class="Input" id="Fk_Down_PicBig" value="<%=Fk_Down_PicBig%>" size="50" />&nbsp;</td>
        <td colspan="2">
		<iframe frameborder="0" width="290" height="25" scrolling="No" id="I1" name="I1" src="PicUpLoad.asp?Form=DownEdit&Input=Fk_Down_Pic"></iframe><br><font color="red">说明：只需上传一张图，自动生成缩略小图和正常大图。</font></td>
    </tr>
    <tr>
        <td height="28" align="right">文件名：</td>
        <td colspan="4">&nbsp;<input name="Fk_Down_FileName" type="text" class="Input" id="Fk_Down_FileName" value="<%=Fk_Down_FileName%>"<%If Fk_Down_FileName<>"" Then%> readonly="readonly"<%End If%> size="50"  style="vertical-align:middle"/>
            *一旦确立不可修改</td>
    </tr>
<%
	Sqlstr="Select * From [Fk_Field] Where Fk_Field_Type=2 Order By Fk_Field_Id Asc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
		Temp2=""
		For Each Temp In Fk_Down_Field
			If Split(Temp,"|-Fangka_Field-|")(0)=Rs("Fk_Field_Tag") Then
				Temp2=FKFun.HTMLDncode(Split(Temp,"|-Fangka_Field-|")(1))
				Exit For
			End If
		Next
%>
    <tr>
        <td height="28" align="right"><%=Rs("Fk_Field_Name")%>：</td>
        <td colspan="4">&nbsp;<input name="Fk_Down__<%=Rs("Fk_Field_Tag")%>" value="<%=Temp2%>" type="text" class="Input" id="Fk_Down__<%=Rs("Fk_Field_Tag")%>"  style="vertical-align:middle"/></td>
    </tr>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
    <tr>
        <td height="28" align="right">转向链接：</td>
        <td colspan="4">&nbsp;<input name="Fk_Down_Url" type="text" class="Input" value="<%=Fk_Down_Url%>" id="Fk_Down_Url" size="50"  style="vertical-align:middle"/>&nbsp;*正常下载请留空</td>
    </tr>
    
    <tr>
        <td height="28" align="right">下载文件：</td>
        <td colspan="4">&nbsp;<input name="Fk_Down_File" value="<%=Fk_Down_File%>" type="text" class="Input" id="Fk_Down_File" size="50"  style="vertical-align:middle"/> 
</td>
    </tr>
    <tr>
        <td height="28" align="right" width="100">下载内容：</td>
        <td colspan="4"><textarea name="Fk_Down_Content" class="<%=bianjiqi%>" id="Fk_Down_Content" rows="8" style="width:100%;"><%=Fk_Down_Content%></textarea></td>
    </tr>
    <tr>
        <td height="28" align="right">模板：</td>
        <td colspan="4">&nbsp;<select name="Fk_Down_Template" class="Input" id="Fk_Down_Template" style="vertical-align:middle">
            <option value="0"<%=FKFun.BeSelect(Fk_Down_Template,0)%>>默认模板</option>
<%
	Sqlstr="Select * From [Fk_Template] Where Not Fk_Template_Name In ('index','info','articlelist','article','productlist','product','gbook','page','subject','job','subject')"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
%>
            <option value="<%=Rs("Fk_Template_Id")%>"<%=FKFun.BeSelect(Fk_Down_Template,Rs("Fk_Template_Id"))%>><%=Rs("Fk_Template_Name")%></option>
    <%
		Rs.MoveNext
	Wend
	Rs.Close
%>
            </select> 推荐：<select name="Fk_Down_Recommend" class="Input" size="1" id="Fk_Down_Recommend" style="vertical-align:middle">
            <option value="0">无推荐</option>
            <option value="2"<%If Instr(Fk_Down_Recommend,",2,")>0 Then%> selected="selected"<%End If%>>推荐</option>
            </select>
			<label><input type="checkbox" name="Fk_Down_onTop" id="Fk_Down_onTop" class="textarea" value="1" <%if Fk_Down_onTop="1" then response.write "checked"%> style="vertical-align:middle"/>置顶</label><select name="Fk_Down_Subject1" class="TextArea" size="1" id="Fk_Down_Subject" style="display:none">
            <option value="0">无专题</option>
            </select> &nbsp; 下载显示：<input name="Fk_Down_Show" type="radio" class="Input" id="Fk_Down_Show" value="1"<%=FKFun.BeCheck(Fk_Down_Show,1)%> checked  style="vertical-align:middle"/><label for="Fk_Down_Show"  style="vertical-align:middle">显示</label>
        <input type="radio" name="Fk_Down_Show" class="Input" id="Fk_Down_Show1" value="0"<%=FKFun.BeCheck(Fk_Down_Show,0)%>  style="vertical-align:middle"/><label for="Fk_Down_Show1" style="vertical-align:middle">不显示</label> &nbsp; &nbsp; 下载量：<input name="Fk_Down_click" type="text" id="Fk_Down_click" class="Input" size="6" maxlength="6" value="<%=Fk_Down_click%>" style="vertical-align:middle;"></td>
    </tr>
    </table>
</div>
<div id="BoxBottom" style="width:96%;">
		<input type="hidden" name="Id" value="<%=Id%>" />
        <input type="hidden" name="ModuleId" value="<%=Fk_Module_Id%>" />
        <input type="submit" onclick="Sends('DownEdit','Down.asp?Type=5',0,'',0,1,'MainRight','<%=Session("NowPage")%>');" class="Button" name="button" id="button" value="修 改" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="btnclose" id="btnclose" value="关 闭" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：DownEditDo
'作    用：执行修改下载
'参    数：
'==============================
Sub DownEditDo()
	Fk_Module_Id=Trim(Request.Form("ModuleId"))
	'判断权限
	'If Not FkFun.CheckLimit("Module"&Fk_Module_Id) Then
		'Response.Write("无权限！")
		'Call FKDB.DB_Close()
		'Session.CodePage=936
		'Response.End()
	'End If
	Fk_Down_Title=FKFun.HTMLEncode(Trim(Request.Form("Fk_Down_Title")))
	Fk_Down_Color=FKFun.HTMLEncode(Trim(Request.Form("Fk_Down_Color")))
	Fk_Down_Keyword=FKFun.HTMLEncode(Trim(Request.Form("Fk_Down_Keyword")))
	Fk_Down_Description=FKFun.HTMLEncode(Trim(Request.Form("Fk_Down_Description")))
	Fk_Down_System=FKFun.HTMLEncode(Trim(Request.Form("Fk_Down_System")))
	Fk_Down_Language=FKFun.HTMLEncode(Trim(Request.Form("Fk_Down_Language")))
	Fk_Down_Url=FKFun.HTMLEncode(Trim(Request.Form("Fk_Down_Url")))
	Fk_Down_File=FKFun.HTMLEncode(Trim(Request.Form("Fk_Down_File")))
	Fk_Down_Pic=FKFun.HTMLEncode(Trim(Request.Form("Fk_Down_Pic")))
	Fk_Down_PicBig=FKFun.HTMLEncode(Trim(Request.Form("Fk_Down_PicBig")))
	Fk_Down_Content=Request.Form("Fk_Down_Content")
	Fk_Down_Recommend=","&FKFun.HTMLEncode(Replace(Trim(Request.Form("Fk_Down_Recommend"))," ",""))&","
	Fk_Down_Subject=","&FKFun.HTMLEncode(Replace(Trim(Request.Form("Fk_Down_Subject"))," ",""))&","
	Fk_Down_FileName=FKFun.HTMLEncode(Trim(Request.Form("Fk_Down_FileName")))
	Fk_Down_Show=Trim(Request.Form("Fk_Down_Show"))
	Fk_Down_click=Trim(Request.Form("Fk_Down_click"))
	Fk_Down_Template=Trim(Request.Form("Fk_Down_Template"))
	Fk_Down_Time=Trim(Request.Form("Fk_Down_Time"))
	Fk_Down_onTop=Trim(Request.Form("Fk_Down_onTop"))
	Fk_Down_px=Trim(Request.Form("Fk_Down_px"))
	If Fk_Down_Time="" Then
		Fk_Down_Time=0
	End If
	Id=Trim(Request.Form("Id"))
	Call FKFun.ShowString(Fk_Down_Title,1,255,0,"请输入下载标题！","下载标题不能大于255个字符！")
	Call FKFun.ShowNum(Fk_Down_px,"排序只能是数字！请重新填写")
	Call FKFun.ShowString(Fk_Down_Keyword,0,255,2,"请输入下载关键字！","下载关键字不能大于255个字符！")
	Call FKFun.ShowString(Fk_Down_Description,0,255,2,"请输入下载描述！","下载描述不能大于255个字符！")
	Call FKFun.ShowString(Fk_Down_System,0,255,2,"请输入适用系统！","适用系统不能大于255个字符！")
	Call FKFun.ShowString(Fk_Down_Language,0,255,2,"请输入语言版本！","语言版本不能大于255个字符！")
	Call FKFun.ShowString(Fk_Down_File,0,255,2,"请输入下载地址或上传文件！","下载地址不能大于255个字符！")
	Call FKFun.ShowString(Fk_Down_Url,0,255,2,"请输入下载转向链接！","下载转向链接不能大于255个字符！")
	If Fk_Down_Url="" Then
		Call FKFun.ShowString(Fk_Down_Content,1,1,1,"请输入下载简介！","下载内容不能大于1个字符！")
	End If
	Call FKFun.ShowString(Fk_Down_Pic,0,255,2,"请输入下载题图路径！","下载题图小图路径不能大于255个字符！")
	Call FKFun.ShowString(Fk_Down_PicBig,0,255,2,"请输入下载题图路径！","下载题图大图路径不能大于255个字符！")
	Call FKFun.ShowNum(Fk_Module_Id,"ModuleId系统参数错误，请刷新页面！")
	Call FKFun.ShowString(Fk_Down_FileName,0,50,2,"请输入下载文件名！","下载文件名不能大于50个字符！")
	Call FKFun.ShowNum(Fk_Down_Template,"请选择模板！")
	Call FKFun.ShowNum(Fk_Down_Show,"请选择下载是否显示！")
	Call FKFun.ShowNum(Fk_Down_click,"请输入正确的下载量！")
	Call FKFun.ShowNum(Id,"Id系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_Field] Where Fk_Field_Type=2 Order By Fk_Field_Id Asc"
	Rs.Open Sqlstr,Conn,1,1
	While Not Rs.Eof
		If Fk_Down_Field="" Then
			Fk_Down_Field=Rs("Fk_Field_Tag")&"|-Fangka_Field-|"&FKFun.HTMLEncode(Trim(Request.Form("Fk_Down__"&Rs("Fk_Field_Tag"))))
		Else
			Fk_Down_Field=Fk_Down_Field&"[-Fangka_Field-]"&Rs("Fk_Field_Tag")&"|-Fangka_Field-|"&FKFun.HTMLEncode(Trim(Request.Form("Fk_Down__"&Rs("Fk_Field_Tag"))))
		End If
		Rs.MoveNext
	Wend
	Rs.Close
	If IsNumeric(Fk_Down_FileName) Then
		Response.Write("文件名不可用纯数字！")
		Call FKDB.DB_Close()
		Response.End()
	End If
	If Left(Fk_Down_FileName,4)="Info" Or Left(Fk_Down_FileName,4)="Page" Or Left(Fk_Down_FileName,5)="GBook" Or Left(Fk_Down_FileName,3)="Job" Then
		Response.Write("文件名受限，不能以一下单词开头：Info、Page、GBook、Job！")
		Call FKDB.DB_Close()
		Response.End()
	End If
	If SiteDelWord=1 Then
		TempArr=Split(Trim(FKFun.UnEscape(FKFso.FsoFileRead("DelWord.dat")))," ")
		For Each Temp In TempArr
			If Temp<>"" Then
				Fk_Down_Content=Replace(Fk_Down_Content,Temp,"**")
				Fk_Down_Title=Replace(Fk_Down_Title,Temp,"**")
				Fk_Down_Keyword=Replace(Fk_Down_Keyword,Temp,"**")
				Fk_Down_Description=Replace(Fk_Down_Description,Temp,"**")
				Fk_Down_System=Replace(Fk_Down_System,Temp,"**")
				Fk_Down_Language=Replace(Fk_Down_Language,Temp,"**")
			End If
		Next
	End If
	Sqlstr="Select * From [Fk_Down] Where Fk_Down_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs("Fk_Down_Title")=Fk_Down_Title
		Rs("Fk_Down_Color")=Fk_Down_Color
		Rs("Fk_Down_Keyword")=Fk_Down_Keyword
		Rs("Fk_Down_Description")=Fk_Down_Description
		Rs("Fk_Down_Url")=Fk_Down_Url
		Rs("Fk_Down_File")=Fk_Down_File
		Rs("Fk_Down_Field")=Fk_Down_Field
		Rs("Fk_Down_System")=Fk_Down_System
		Rs("Fk_Down_Language")=Fk_Down_Language
		Rs("Fk_Down_Pic")=Fk_Down_Pic
		Rs("Fk_Down_PicBig")=Fk_Down_PicBig
		Rs("Fk_Down_Recommend")=Fk_Down_Recommend
		Rs("Fk_Down_Subject")=Fk_Down_Subject
		Rs("Fk_Down_Show")=Fk_Down_Show
		Rs("Fk_Down_click")=Fk_Down_click
		Rs("Fk_Down_Content")=Fk_Down_Content
		Rs("Fk_Down_FileName")=Fk_Down_FileName
		Rs("Fk_Down_Template")=Fk_Down_Template
		Rs("Fk_Down_Ip")=Fk_Down_onTop
		Rs("Px")=Fk_Down_px
		If Fk_Down_Time=1 Then
			Rs("Fk_Down_Time")=Now()
		End If
		Rs.Update()
		Application.UnLock()
		Response.Write("“"&Fk_Down_Title&"”修改成功！")
		'插入日志
		on error resume next
		dim log_content,log_ip,log_user
		log_content="修改下载：【"&Fk_Down_Title&"】"
		log_user=Request.Cookies("FkAdminName")
		
		log_ip=FKFun.getIP()
		conn.execute("insert into newTB_log (log_user,log_content,log_ip) values('"&log_user&"','"&log_content&"','"&log_ip&"')")
	Else
		Response.Write("下载不存在！")
	End If
	Rs.Close
	If SiteHtml=1 And Fk_Down_Show=1 And Fk_Down_Url="" Then
		Dim FKHTML
		Set FKHTML=New Cls_HTML
		Sqlstr="Select * From [Fk_DownList] Where Fk_Down_Id=" & Id
		Rs.Open Sqlstr,Conn,1,1
		Fk_Module_Dir=Rs("Fk_Module_Dir")
		Fk_Down_Module=Rs("Fk_Down_Module")
		Rs.Close
		Call FKHTML.CreatDown(Fk_Down_Template,Fk_Down_Module,Fk_Module_Dir,Fk_Down_FileName,Fk_Down_Title,1)
	Else
		Sqlstr="Select * From [Fk_DownList] Where Fk_Down_Id=" & Id
		Rs.Open Sqlstr,Conn,1,1
		Fk_Module_Dir=Rs("Fk_Module_Dir")
		Fk_Down_Module=Rs("Fk_Down_Module")
		Rs.Close
		If Fk_Module_Dir<>"" Then
			Temp="../"&Fk_Down_Module&"/"
		Else
			Temp="../Down"&Fk_Down_Module&"/"
		End If
		If Fk_Down_FileName<>"" Then
			Temp=Temp&Fk_Down_FileName&".html"
		Else
			Temp=Temp&Id&".html"
		End If
		Call FKFso.DelFile(Temp)
	End If
End Sub

'==============================
'函 数 名：DownDelDo
'作    用：执行删除下载
'参    数：
'==============================
Sub DownDelDo()
	Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Id,"Id系统参数错误，请刷新页面！")
	Sqlstr="Select * From [Fk_DownList] Where Fk_Down_Id=" & Id
	Rs.Open Sqlstr,Conn,1,1
	If Not Rs.Eof Then
		'判断权限
		'If Not FkFun.CheckLimit("Module"&Rs("Fk_Down_Module")) Then
			'Response.Write("无权限！")
			'Call FKDB.DB_Close()
			'Session.CodePage=936
			'Response.End()
		'End If
		Fk_Module_Dir=Rs("Fk_Module_Dir")
		Fk_Down_Module=Rs("Fk_Down_Module")
		Fk_Down_FileName=Rs("Fk_Down_FileName")
		Fk_Down_Title=Rs("Fk_Down_Title")
		Rs.Close
		If Fk_Module_Dir<>"" Then
			Temp="../"&Fk_Down_Module&"/"
		Else
			Temp="../Down"&Fk_Down_Module&"/"
		End If
		If Fk_Down_FileName<>"" Then
			Temp=Temp&Fk_Down_FileName&".html"
		Else
			Temp=Temp&Id&".html"
		End If
		Call FKFso.DelFile(Temp)
	Else
		Rs.Close
		Response.Write("下载不存在！")
		Call FKDB.DB_Close()
		Session.CodePage=936
		Response.End()
	End If
	Sqlstr="Select * From [Fk_Down] Where Fk_Down_Id=" & Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		Response.Write("下载删除成功！")
		'插入日志
		on error resume next
		dim log_content,log_ip,log_user
		log_content="删除下载：【"&Fk_Down_Title&"】"
		log_user=Request.Cookies("FkAdminName")
		
		log_ip=FKFun.getIP()
		conn.execute("insert into newTB_log (log_user,log_content,log_ip) values('"&log_user&"','"&log_content&"','"&log_ip&"')")
	Else
		Response.Write("下载不存在！")
	End If
	Rs.Close
End Sub

'==============================
'函 数 名：ListDelDo
'作    用：执行批量删除下载
'参    数：
'==============================
Sub ListDelDo()
	Id=Replace(Trim(Request.Form("ListId"))," ","")
	If Id="" Then
		Response.Write("请选择要删除的下载！")
		Call FKDB.DB_Close()
		Response.End()
	End If
	Sqlstr="select Fk_Down_Title From [Fk_Down] Where Fk_Down_Id In ("&Id&")"
	
	
	Rs.Open Sqlstr,Conn,1,3
	if not rs.eof then
		i=0
		do while not rs.eof
		if i=0 then
			Fk_Down_Title="【"&rs("Fk_Down_Title")&"】"
		else
			Fk_Down_Title=Fk_Down_Title&","&"【"&rs("Fk_Down_Title")&"】"
		end if
		i=i+1
		Application.Lock()
		Rs.Delete()
		Application.UnLock()
		rs.movenext
		loop
	end if
	Response.Write("下载批量删除成功！")
	'插入日志
	on error resume next
	dim log_content,log_ip,log_user
	log_content="批量删除下载："&Fk_Down_Title&""
	log_user=Request.Cookies("FkAdminName")
	
	log_ip=FKFun.getIP()
	conn.execute("insert into newTB_log (log_user,log_content,log_ip) values('"&log_user&"','"&log_content&"','"&log_ip&"')")
End Sub

'==============================
'函 数 名：DownMove
'作    用：执行批量移动下载
'参    数：
'==============================
Sub DownMove()
	Dim Fk_Module_Type
	Id=Replace(Trim(Request.QueryString("ListId"))," ","")
	Fk_Module_Id=Trim(Request.QueryString("Id"))
	Call FKFun.ShowNum(Fk_Module_Id,"请选择转移到的模块！")
	If Id="" Then
		Response.Write("请选择要移动的下载！")
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
		Response.Write("要移到的模块不存在！")
		Call FKDB.DB_Close()
		Response.End()
	End If
	If Fk_Module_Type<>7 Then
		Response.Write("只能移动到下载模块！")
		Call FKDB.DB_Close()
		Response.End()
	End If
	Sqlstr="Update [Fk_Down] Set Fk_Down_Module="&Fk_Module_Id&" Where Fk_Down_Id In ("&Id&")"
	Application.Lock()
	Conn.Execute(Sqlstr)
	Application.UnLock()
	Response.Write("下载批量移动成功！")
End Sub
%><!--#Include File="../Code.asp"-->