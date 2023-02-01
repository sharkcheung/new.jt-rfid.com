<!--#Include File="AdminCheck.asp"-->
<%
'==========================================
'文 件 名：QQ.asp
'文件用途：客服浮窗拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'判断权限
If Request.Cookies("FkAdminLimitId")>0 Then
	Response.Write("无权限！")
	Call FKDB.DB_Close()
	Session.CodePage=936
	Response.End()
End If

Dim Fk_QQ_Content

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call QQBox() '读取客服浮窗
	Case 2
		Call QQDo() '设置客服浮窗
End Select

'==========================================
'函 数 名：QQBox()
'作    用：读取客服浮窗
'参    数：
'==========================================
Sub QQBox()
	Sqlstr="Select * From [Fk_QQ]"
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Fk_QQ_Content=FKFun.HTMLDncode(Rs("Fk_QQ_Content"))
	Else
		PageErr=1
	End If
	Rs.Close
%>
<form id="QQSet" name="QQSet" method="post" action="QQ.asp?Type=2" onsubmit="return false;">
<div id="BoxTop" style="width:600px;">客服浮窗设置[按ESC关闭窗口]</div>
<div id="BoxContents" style="width:600px;">
	<table width="90%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
            <td height="30" align="right" class="MainTableTop">客服代码：</td>
            <td>&nbsp;
                <textarea name="Fk_QQ_Content" cols="50" rows="20" class="TextArea" id="Fk_QQ_Content"><%=Fk_QQ_Content%></textarea></td>
        </tr>
    </table>
</div>
<div id="BoxBottom" style="width:580px;">
        <input type="submit" onclick="Sends('QQSet','QQ.asp?Type=2',0,'',0,0,'','');" class="Button" name="button" id="button" value="设 置" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==========================================
'函 数 名：QQDo()
'作    用：设置客服浮窗
'参    数：
'==========================================
Sub QQDo()
	Fk_QQ_Content=FKFun.HTMLEncode(Request.Form("Fk_QQ_Content"))
	Sqlstr="Update [Fk_QQ] Set Fk_QQ_Content='"&Fk_QQ_Content&"'"
	Application.Lock()
	Conn.Execute(Sqlstr)
	Application.UnLock()
	Response.Write("客服浮窗内容修改成功！")
End Sub
%>
<!--#Include File="../Code.asp"-->