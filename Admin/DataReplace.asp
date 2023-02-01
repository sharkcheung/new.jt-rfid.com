<!--#Include File="AdminCheck.asp"-->
<!--#Include File="../Inc/Md5.asp"-->
<%
'==========================================
'文 件 名：Admin.asp
'文件用途：管理员管理拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'判断权限
'If Request.Cookies("FkAdminLimitId")>0 Then
'	Response.Write("无权限！")
'	Call FKDB.DB_Close()
'	Session.CodePage=936
'	Response.End()
'End If

'定义页面变量
Dim Fk_Admin_LoginName,Fk_Admin_LoginPass1,Fk_Admin_LoginPass2,Fk_Admin_Name,Fk_Admin_User,Fk_Admin_Limit,s

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call AdminAddForm() '添加管理员表单
	Case 2
		Call DataReplaceDo() '执行删除管理员
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：AdminAddForm()
'作    用：添加管理员表单
'参    数：
'==========================================
Sub AdminAddForm()
%>
<form id="AdminAdd" name="AdminAdd" method="post" action="DataReplace.asp?Type=2" onsubmit="return false;">
<div id="BoxTop" style="width:98%;"><span>数据替换</span></div>
<div id="BoxContents" style="width:98%;">
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
            <td height="30" align="right">栏目类型：</td>
            <td>&nbsp;<select name="moduletype" class="Input" id="moduletype">
                <option value="1">新闻栏目</option>
                <option value="2">图文栏目</option>
                <option value="3">单页栏目</option>
                <option value="4">下载栏目</option>
				</select></td>
        </tr>
        <tr>
            <td height="30" align="right">要替换的内容：</td>
            <td>&nbsp;<input name="replacecontent" type="text" class="Input" id="replacecontent" /></td>
        </tr>
        <tr>
            <td height="30" align="right">替换成内容：</td>
            <td>&nbsp;<input name="replaceto" type="text" class="Input" id="replaceto" /></td>
        </tr>
	</table>
</div>
<div id="BoxBottom" style="width:96%;">
        <input type="submit" onclick="Sends('AdminAdd','DataReplace.asp?Type=2',0,'',0,1,'MainRight','DataReplace.asp?Type=1');" class="Button" name="button" id="button" value="替 换" />
</div>
</form>
<%
End Sub

'==============================
'函 数 名：DataReplaceDo
'作    用：执行替换内容
'参    数：
'==============================
Sub DataReplaceDo()
	dim moduletype,replacecontent,replaceto
	moduletype=Trim(Request.Form("moduletype"))
	replacecontent=Trim(Request.Form("replacecontent"))
	replaceto=Trim(Request.Form("replaceto"))
	Call FKFun.ShowNum(moduletype,"系统参数错误，请刷新页面！")
	Call FKFun.ShowString(replacecontent,1,100,0,"请输入要替换的内容！","替换的内容不能大于100个字符！")
	Call FKFun.ShowString(replaceto,0,100,0,"请输入要替换的内容！","替换成的内容不能大于100个字符！")
	if moduletype=1 then
		Sqlstr="select * from Fk_Article"
		Rs.open Sqlstr,conn,1,3
		while not Rs.eof 
			Rs("Fk_Article_Content") = Replace(Rs("Fk_Article_Content"),replacecontent,replaceto)
			Rs.update()
			Rs.movenext
		wend
	elseif moduletype=2 then
		Sqlstr="select * from Fk_Product"
		Rs.open Sqlstr,conn,1,3
		while not Rs.eof 
			Rs("Fk_Product_Content") = Replace(Rs("Fk_Product_Content"),replacecontent,replaceto)
			Rs.update()
			Rs.movenext
		wend
	elseif moduletype=3 then
		Sqlstr="select * from Fk_Module"
		Rs.open Sqlstr,conn,1,3
		while not Rs.eof 
			if not IsNull(Rs("Fk_Module_Content")) then
			Rs("Fk_Module_Content") = Replace(Rs("Fk_Module_Content"),replacecontent,replaceto)
			Rs.update()
			end if
			Rs.movenext
		wend
	elseif moduletype=4 then
		Sqlstr="select * from Fk_Down"
		Rs.open Sqlstr,conn,1,3
		while not Rs.eof 
			Rs("Fk_Down_Content") = Replace(Rs("Fk_Down_Content"),replacecontent,replaceto)
			Rs.update()
			Rs.movenext
		wend
	end if
	Response.Write("内容替换成功！")
	Rs.Close
End Sub
%><!--#Include File="../Code.asp"-->