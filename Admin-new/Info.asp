<!--#Include File="AdminCheck.asp"-->
<!--#Include File="../Class/Cls_HTML.asp"-->
<%
'==========================================
'文 件 名：Info.asp
'文件用途：信息管理拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'定义页面变量
Dim Fk_Module_Name,Fk_Module_Id,Fk_Module_Menu,Fk_Module_Content,Fk_Module_Template,Fk_Module_FileName,Fk_Module_Keyword,Fk_Module_Description,Fk_Module_Seotitle

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call InfoForm() '修改信息表单
	Case 2
		Call InfoDo() '执行修改信息
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：InfoForm()
'作    用：修改信息表单
'参    数：
'==========================================
Sub InfoForm()
	Fk_Module_Id=Clng(Trim(Request.QueryString("ModuleId")))
	'判断权限
	If Not FkFun.CheckLimit("Module"&Fk_Module_Id) Then
		'Response.Write("无权限！")
		'Call FKDB.DB_Close()
		'Session.CodePage=936
		'Response.End()
	End If
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id=" & Fk_Module_Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Fk_Module_Name=Rs("Fk_Module_Name")
		Fk_Module_Menu=Rs("Fk_Module_Menu")
		Fk_Module_Content=Rs("Fk_Module_Content")
		Fk_Module_Template=Rs("Fk_Module_Template")
		Fk_Module_FileName=Rs("Fk_Module_FileName")
		Fk_Module_Seotitle=Rs("Fk_Module_Seotitle")
		Fk_Module_Keyword=Rs("Fk_Module_Keyword")
		Fk_Module_Description=Rs("Fk_Module_Description")
	End If
	Rs.Close
%>
<form id="InfoEdit" name="InfoEdit" method="post" action="Info.asp?Type=2" onsubmit="return false;">
<div id="BoxContents" style="width:93%; padding-top:20px;">
<table width="100%" border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
        <td width="100" height="25" align="right">标题：</td>
        <td><input name="Fk_Module_Name" value="<%=Fk_Module_Name%>" type="text" class="Input" id="Fk_Module_Name" size="50" <%If SiteToPinyin=1 Then%> onmousemove="GetPinyin('Fk_Module_FileName','ToPinyin.asp?Str='+this.value);" <%End If%> /></td>
    </tr>
    <tr>
        <td height="25" align="right">SEO标题：</td>
        <td><input name="Fk_Module_Seotitle" value="<%=Fk_Module_Seotitle%>" type="text" class="Input" id="Fk_Module_Seotitle" size="50" /></td>
    </tr>
    <tr>
        <td height="25" align="right">SEO关键词：</td>
        <td><input name="Fk_Module_Keyword" value="<%=Fk_Module_Keyword%>" type="text" class="Input" id="Fk_Module_Keyword" size="50" /></td>
    </tr>
    <tr>
        <td height="25" align="right">SEO描述：</td>
        <td><input name="Fk_Module_Description" value="<%=Fk_Module_Description%>" type="text" class="Input" id="Fk_Module_Description" size="70" /></td>
    </tr>
    <tr>
        <td height="25" align="right">文件名：</td>
        <td><input name="Fk_Module_FileName" type="text" class="Input" id="Fk_Module_FileName" value="<%=Fk_Module_FileName%>"<%If Fk_Module_FileName<>"" Then%> readonly="readonly"<%End If%> size="50" />&nbsp;*不可修改</td>
    </tr>
 <tr>
        <td height="25" align="right">内容：</td>
        <td style="padding:10px 0 10px 10px;"><textarea  style="width:100%;" name="Fk_Module_Content" class="<%=bianjiqi%>" id="Fk_Module_Content" rows="15"><%=Fk_Module_Content%></textarea></td>
    </tr>
    <tr>
        <td align="right">模板：</td>
        <td>&nbsp;<select name="Fk_Module_Template" class="Input" id="Fk_Module_Template">
            <option value="0"<%=FKFun.BeSelect(Fk_Module_Template,0)%>>默认模板</option>
<%
	Sqlstr="Select * From [Fk_Template] Where Not Fk_Template_Name In ('index','info','articlelist','article','productlist','product','gbook','page','subject','job','subject','top','bottom','downlist','down')"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof
%>
            <option value="<%=Rs("Fk_Template_Id")%>"<%=FKFun.BeSelect(Fk_Module_Template,Rs("Fk_Template_Id"))%>><%=Rs("Fk_Template_Name")%></option>
<%
		Rs.MoveNext
	Wend
	Rs.Close
%>
            </select>
        </td>
    </tr>
   
</table>
</div>
<div id="BoxBottom" style="width:93%; margin:0 auto;" class="tcbtm">
		<input type="hidden" name="ModuleId" value="<%=Fk_Module_Id%>" />
        <input type="submit" onclick="Sends('InfoEdit','Info.asp?Type=2',0,'',0,0,'','');" class="Button" name="button" id="button" value="修 改" />
        <input type="button" onclick="layer.closeAll();$('select').show();" class="Button close" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==========================================
'函 数 名：InfoDo()
'作    用：执行修改信息
'参    数：
'==========================================
Sub InfoDo()
	Fk_Module_Id=Trim(Request.Form("ModuleId"))
	'判断权限
	If Not FkFun.CheckLimit("Module"&Fk_Module_Id) Then
		'Response.Write("无权限！")
		'Call FKDB.DB_Close()
		'Session.CodePage=936
		'Response.End()
	End If
	Fk_Module_Name=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Name")))
	Fk_Module_Content=Request.Form("Fk_Module_Content")
	Fk_Module_FileName=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_FileName")))
	Fk_Module_Template=Trim(Request.Form("Fk_Module_Template"))
	Fk_Module_Seotitle=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Seotitle")))
	Fk_Module_Keyword=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Keyword")))
	Fk_Module_Description=FKFun.HTMLEncode(Trim(Request.Form("Fk_Module_Description")))
	Call FKFun.ShowString(Fk_Module_Name,1,255,0,"请输入信息标题！","信息标题不能大于255个字符！")
	Call FKFun.ShowString(Fk_Module_Content,5,1,1,"请输入信息内容，不少于5个字符！","信息内容不能大于1个字符！")
	Call FKFun.ShowString(Fk_Module_FileName,0,50,2,"请输入模块文件名！","模块文件名不能大于50个字符！")
	Call FKFun.ShowNum(Fk_Module_Template,"请选择模块模板！")
	Call FKFun.ShowString(Fk_Module_Seotitle,0,255,2,"请输入SEO标题！","SEO标题不能大于255个字符！")
	Call FKFun.ShowString(Fk_Module_Keyword,0,255,2,"请输入SEO关键词！","SEO关键词不能大于255个字符！")
	Call FKFun.ShowString(Fk_Module_Description,0,255,2,"请输入SEO描述！","SEO描述不能大于255个字符！")
	Call FKFun.ShowNum(Fk_Module_Id,"ModuleId系统参数错误，请刷新页面！")
	If Left(Fk_Module_FileName,4)="Info" Or Left(Fk_Module_FileName,4)="Page" Or Left(Fk_Module_FileName,5)="GBook" Or Left(Fk_Module_FileName,3)="Job" Then
		Response.Write("文件名受限，不能以一下单词开头：Info、Page、GBook、Job！")
		Call FKDB.DB_Close()
		Response.End()
	End If
	Sqlstr="Select * From [Fk_Module] Where Fk_Module_Id=" & Fk_Module_Id
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs("Fk_Module_Name")=Fk_Module_Name
		Rs("Fk_Module_Seotitle")=Fk_Module_Seotitle
		Rs("Fk_Module_Keyword")=Fk_Module_Keyword
		Rs("Fk_Module_Description")=Fk_Module_Description
		Rs("Fk_Module_Content")=Fk_Module_Content
		Rs("Fk_Module_FileName")=Fk_Module_FileName
		Rs("Fk_Module_Template")=Fk_Module_Template
		Rs.Update()
		Application.UnLock()
		'插入日志
		on error resume next
		dim log_content,log_ip,log_user
		log_content="修改单页信息：【"&Fk_Module_Name&"】"
		log_user=Request.Cookies("FkAdminName")
		
		log_ip=FKFun.getIP()
		conn.execute("insert into newTB_log (log_user,log_content,log_ip) values('"&log_user&"','"&log_content&"','"&log_ip&"')")
		Response.Write(Fk_Module_Name&"修改成功！")
	Else
		Response.Write("信息不存在！")
	End If
	Rs.Close
	If SiteHtml=1 Then
		Dim FKHTML
		Set FKHTML=New Cls_HTML
		Id=Fk_Module_Id
		Call FKHTML.CreatInfo(Fk_Module_Template,Fk_Module_FileName,Fk_Module_Name,1)
	End If
End Sub
%><!--#Include File="../Code.asp"-->