<!--#Include File="AdminCheck.asp"-->
<%
'==========================================
'文 件 名：Data.asp
'文件用途：数据库操作
'版权所有：企帮网络www.qebang.cn
'==========================================

'判断权限
If Request.Cookies("FkAdminLimitId")>0 Then
	'Response.Write("无权限！")
	'Call FKDB.DB_Close()
	'Session.CodePage=936
	'Response.End()
End If

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call DataList() '操作主页面
	Case 2
		Call DoSql() '执行SQL语句
	Case 3
		Call CompactDB() '整理数据库
	Case 4
	    Call BakDB()  '备份数据库
	Case 5
	    Call ReBakDB() '还原数据库
	Case Else
		Response.Write("没有找到此功能项！")
End Select

'==========================================
'函 数 名：DataList()
'作    用：操作主页面
'参    数：
'==========================================
Sub DataList()
%>
<div id="ListNav">
    <ul>
        <li><a href="javascript:void(0);" onclick="SetRContent('MainRight','Data.asp?Type=1');return false">刷新</a></li>
    </ul>
</div>
<div id="ListTop">
    数据维护
</div>
<div id="ListContent">
    <table width="100%" bordercolor="#CCCCCC" border="1" cellspacing="0" cellpadding="0">
        <tr style="display:none">
            <td align="center" class="ListTdTop">SQL语句执行</td>
        </tr>
        <form id="Sql" name="Sql" method="post" action="Data.asp?Type=2" onsubmit="return false;">
        <tr style="display:none">
            <td height="20" align="center">Sql语句：<input name="Sqlstr" type="text" class="Input" id="Sqlstr" size="50" />&nbsp;&nbsp;<input type="submit" onclick="Sends('Sql','Data.asp?Type=2',0,'',0,1,'MainRight','Data.asp?Type=1');" class="Button" name="button" id="button" value="执 行" />&nbsp;&nbsp;<span style="color:#F00">执行SQL语句仅供特殊情况使用！</span></td>
        </tr>
        </form>
        <tr>
            <td align="center" class="ListTdTop">数据维护</td>
        </tr>
        <tr>
            <td height="30" align="center">数据整理：
                <input type="button" onclick="DelIt('是否要整理数据？','Data.asp?Type=3','MainRight','Data.asp?Type=1');" class="Button" name="button" id="button" value="整 理" /> 
			整理数据有利于提供运行效率，建议每月运行一次。</td>
        </tr>
        <tr>
            <td height="30" align="center">数据备份：
                <input type="button" onclick="DelIt('是否要备份数据？','Data.asp?Type=4','MainRight','Data.asp?Type=1');" class="Button" name="button" id="button" value="备 份" /> 
			备份数据名以当前日期为命名，建议每周运行一次。</td>
        </tr>
        <tr>
            <td height="30" align="center">数据恢复：不能直接在此恢复数据，需恢复备份数据，请联系您的客服人员。</td>
        </tr>
        <tr>
            <td align="center" class="ListTdTop">空间占用</td>
        </tr>
        <tr>
            <td height="30" align="center">占用空间：<%=GetSize(SiteDir)%>；附件占用：<%=GetSize(SiteDir&"Up")%></td>
        </tr>
        <tr>
            <td height="25" align="center">&nbsp;</td>
        </tr>
    </table>
</div>
<div id="ListBottom">

</div>
<%
End Sub

'==========================================
'函 数 名：DoSql()
'作    用：执行SQL语句
'参    数：
'==========================================
Sub DoSql()
	'Sqlstr=Request.Form("Sqlstr")
	'Application.Lock()
	'Conn.Execute(Sqlstr)
	'Application.UnLock()
	'Response.Write("SQL语句执行成功！")
End Sub

'==========================================
'函 数 名：CompactDB()
'作    用：整理数据库
'参    数：
'==========================================
Sub CompactDB()
	Call FKDB.DB_Close()
	Dim Engine
	Set Fso=CreateObject("Scri"&"pting.FileS"&"ystemO"&"bject")
	Set Engine=CreateObject("JRO.JetEngine")
	Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(SiteData),"Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&Server.MapPath(SiteDir)&"\Temp.mdb"
	Fso.CopyFile Server.MapPath(SiteDir)&"\Temp.mdb",Server.MapPath(SiteData)
	Fso.DeleteFile(Server.MapPath(SiteDir)&"\Temp.mdb")
	Set Fso = nothing
	Set Engine = nothing
	Call FKDB.DB_Open()
	
	Response.Write("数据整理成功！")
End Sub

'==========================================
'函 数 名：BakDB()
'作    用：备份数据库
'参    数：
'==========================================
Sub BakDB()
	Call FKDB.DB_Close()
	Set Fso=CreateObject("Scri"&"pting.FileS"&"ystemO"&"bject")
	Fso.CopyFile Server.MapPath(SiteData),Server.MapPath(SiteDir&SiteDBDir)&"\"&formatdatetime(now(),1)&".Data.mdb.数据库备份"   '数据库备份
	Response.Write("数据库备份成功！")
	Fso.CopyFile Server.MapPath(SiteDir&"inc")&"\Site.asp",Server.MapPath(SiteDir&SiteDBDir)&"\"&formatdatetime(now(),1)&".Site.asp.站点配置备份"   '站点配置文件备份
	Response.Write("配置文件备份成功！")
	Fso.CopyFile Server.MapPath(SiteDir&"admin")&"\KeyWord.dat",Server.MapPath(SiteDir&SiteDBDir)&"\"&formatdatetime(now(),1)&".KeyWord.dat.关键词库备份"   '关键词库文件备份
	Response.Write("关键词库备份成功！")
	Set Fso = nothing
	Call FKDB.DB_Open()
End Sub

'==========================================
'函 数 名：ReBakDB()
'作    用：恢复数据库
'参    数：
'==========================================
Sub ReBakDB()
	Dim RebakDBdate
	RebakDBdate=Request.QueryString("RebakDBdate")
	Call FKDB.DB_Close()
	Set Fso=CreateObject("Scri"&"pting.FileS"&"ystemO"&"bject")
	If RebakDBdate<>"" and Fso.FileExists(RebakDBdate&".dbbak")=True then
		Fso.CopyFile Server.MapPath(SiteDir&SiteDBDir)&"\"&RebakDBdate&".qbbak",Server.MapPath(SiteData)
		Response.Write("数据还原成功,数据已经恢复到"&RebakDBdate&",请更新缓存。")
	Else
		Response.Write("您选择的备份日期数据不存在，恢复失败！")
	End if
	Set Fso = nothing
	Call FKDB.DB_Open()
End Sub


'==========================================
'函 数 名：GetSize()
'作    用：获取空间使用
'参    数：
'==========================================
Function GetSize(Path)
	On Error Resume Next
	Dim Size,ShowSize,Paths
	Set Fso=CreateObject("Scri"&"pting.FileS"&"ystemO"&"bject")
	Paths=Path
	Paths=Server.Mappath(Paths) 		 		
	Set F=Fso.GetFolder(Paths) 		
	Size=F.Size
	ShowSize=Size&"&nbsp;Byte" 
	If Size>1024 Then
		Size=(Size/1024)
		ShowSize=Size&"&nbsp;KB"
	End If
	If Size>1024 Then
		Size=(Size/1024)
		ShowSize=FormatNumber(Size,2)&"&nbsp;MB"		
	End If
	If Size>1024 then
		Size=(Size/1024)
		ShowSize=FormatNumber(Size,2)&"&nbsp;GB"	   
	End If   
	Set Fso = nothing
	GetSize=ShowSize
End function	
%><!--#Include File="../Code.asp"-->