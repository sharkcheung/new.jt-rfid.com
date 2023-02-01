<!--#Include File="AdminCheck.asp"-->
<%
'==========================================
'文 件 名：Jpeg.asp
'文件用途：水印缩略设置拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'判断权限
If Request.Cookies("FkAdminLimitId")>0 Then
	'Response.Write("无权限！")
	'Call FKDB.DB_Close()
	'Session.CodePage=936
	'Response.End()
End If

Dim Fk_Jpeg_Water,Fk_Jpeg_Small,Fk_Jpeg_Content,JpegNow
Dim J1,J2,J3,J4,J5,J6,K1,K2,K3,K4

'获取参数
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call JpegBox() '读取缩略水印设置
	Case 2
		Call JpegDo() '缩略水印设置
	Case 3
		Call WaterTest() '水印测试
End Select

'==========================================
'函 数 名：JpegBox()
'作    用：读取缩略水印设置
'参    数：
'==========================================
Sub JpegBox()
	JpegNow=FKFun.IsObjInstalled("Persits.Jpeg")
	Sqlstr="Select * From [Fk_Jpeg]"
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		If JpegNow=False Then
			Application.Lock()
			Rs("Fk_Jpeg_Water")=0
			Rs("Fk_Jpeg_Small")=0
			Rs.Update
			Application.UnLock()
		Else
			Fk_Jpeg_Water=Rs("Fk_Jpeg_Water")
			Fk_Jpeg_Small=Rs("Fk_Jpeg_Small")
			TempArr=Split(Rs("Fk_Jpeg_Content"),"|-_-|")
		End If
	Else
		Rs.Close
		Response.Write("设置丢失！")
		Call FKDB.DB_Close()
		Session.CodePage=936
		Response.End()
	End If
	Rs.Close
%>
<form id="JpegSets" name="JpegSets" method="post" action="Jpeg.asp?Type=2" onsubmit="return false;">
<div id="BoxTop" style="width:700px;"><span>水印和缩略图设置</span><a onclick="$('#Boxs').hide();$('select').show();"><img src="images/close3.gif"></a></div>
<div id="BoxContents" style="width:700px;">
	<table width="97%" border="0" bordercolor="#999999" align="center" cellpadding="0" cellspacing="0">
<%
	If JpegNow=True Then
%>
         <tr>
            <td height="30" align="right" class="MainTableTop">自动水印：</td>
            <td>&nbsp;<input name="Fk_Jpeg_Water" class="Input" type="radio" id="Fk_Jpeg_Water" value="1"<%=FKFun.BeCheck(Fk_Jpeg_Water,1)%> />开放
            <input type="radio" name="Fk_Jpeg_Water" class="Input" id="Fk_Jpeg_Water" value="0"<%=FKFun.BeCheck(Fk_Jpeg_Water,0)%> />关闭</td>
        </tr>
         <tr>
            <td height="30" align="right" class="MainTableTop">水印参数：</td>
            <td>&nbsp;颜色：<select class="Input" name="J1" id="J1">
                    <option value="&HFFFFFF"<%=FKFun.BeSelect("&HFFFFFF",TempArr(0))%>>白色</option>
                    <option value="&HFF0000"<%=FKFun.BeSelect("&HFF0000",TempArr(0))%>>红色</option>
                    <option value="&H000000"<%=FKFun.BeSelect("&H000000",TempArr(0))%>>黑色</option>
                </select>&nbsp;字体：<select class="Input" name="J2" id="J2">
                    <option value="宋体"<%=FKFun.BeSelect("宋体",TempArr(1))%>>宋体</option>
                </select>&nbsp;加粗：<select class="Input" name="J3" id="J3">
                    <option value="False"<%=FKFun.BeSelect("False",TempArr(2))%>>不加粗</option>
                    <option value="True"<%=FKFun.BeSelect("True",TempArr(2))%>>加粗</option>
                </select>
                <br />&nbsp;左方位：<input name="J4" type="text" class="Input" id="J4" value="<%=TempArr(3)%>" size="10" />&nbsp;上方位：<input name="J5" type="text" class="Input" id="J5" value="<%=TempArr(4)%>" size="10" />
                <br />&nbsp;水印文字：<input name="J6" type="text" class="Input" id="J6" value="<%=TempArr(5)%>" size="50" />&nbsp;<input type="button" onclick="SetRContent('WaterTest','Jpeg.asp?Type=3&J1='+escape(document.all.J1.options[document.all.J1.selectedIndex].value)+'&J2='+escape(document.all.J2.options[document.all.J2.selectedIndex].value)+'&J3='+document.all.J3.options[document.all.J3.selectedIndex].value+'&J4='+document.all.J4.value+'&J5='+document.all.J5.value+'&J6='+escape(document.all.J6.value))" class="Button" name="button" id="button" value="测试效果" />
                </td>
        </tr>
         <tr>
             <td height="30" align="right" class="MainTableTop">水印效果：</td>
             <td id="WaterTest">&nbsp;设置水印参数后点测试效果</td>
         </tr>
         <tr>
            <td height="30" align="right" class="MainTableTop">自动缩略图：</td>
            <td>&nbsp;<input name="Fk_Jpeg_Small" class="Input" type="radio" id="Fk_Jpeg_Small" value="1"<%=FKFun.BeCheck(Fk_Jpeg_Small,1)%> />开启
            <input type="radio" name="Fk_Jpeg_Small" class="Input" id="Fk_Jpeg_Small" value="0"<%=FKFun.BeCheck(Fk_Jpeg_Small,0)%> />关闭</td>
        </tr>
         <tr>
            <td height="30" align="right" class="MainTableTop">缩略参数：</td>
            <td>&nbsp;缩略图限宽：<input name="K1" type="text" class="Input" id="K1" value="<%=TempArr(15)%>" size="10" />&nbsp;缩略图限高：<input name="K2" type="text" class="Input" id="K2" value="<%=TempArr(16)%>" size="10" />
            <br />&nbsp;编辑器图限宽：<input name="K3" type="text" class="Input" id="K3" value="<%=TempArr(17)%>" size="10" />&nbsp;编辑器图限高：<input name="K4" type="text" class="Input" id="K4" value="<%=TempArr(18)%>" size="10" />
                </td>
        </tr>
<%
	Else
%>
       <tr>
            <td height="30" align="right" class="MainTableTop">提醒：</td>
            <td style="color:#F00;">&nbsp;您的空间不支持AspJpeg组件，水印和缩略无法使用！</td>
        </tr>
<%
	End If
%>
    </table>
</div>
<div id="BoxBottom" style="width:680px;">
        <input type="submit" onclick="Sends('JpegSets','Jpeg.asp?Type=2',0,'',0,0,'','');" class="Button" name="button" id="button" value="设 置" />
        <input type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
</div>
</form>
<%
End Sub

'==========================================
'函 数 名：JpegDo()
'作    用：缩略水印设置
'参    数：
'==========================================
Sub JpegDo()
	Fk_Jpeg_Water=Trim(Request.Form("Fk_Jpeg_Water"))
	Fk_Jpeg_Small=Trim(Request.Form("Fk_Jpeg_Small"))
	J1=FKFun.HTMLEncode(Request.Form("J1"))
	J2=FKFun.HTMLEncode(Request.Form("J2"))
	J3=FKFun.HTMLEncode(Request.Form("J3"))
	J4=Request.Form("J4")
	J5=Request.Form("J5")
	J6=FKFun.HTMLEncode(Request.Form("J6"))
	K1=Request.Form("K1")
	K2=Request.Form("K2")
	K3=Request.Form("K3")
	K4=Request.Form("K4")
	Call FKFun.ShowString(J1,1,50,0,"请选择水印颜色！","水印颜色不能大于50个字符！")
	Call FKFun.ShowString(J2,1,50,0,"请选择水印字体！","水印字体不能大于50个字符！")
	Call FKFun.ShowString(J3,1,50,0,"请选择是否加粗！","是否加粗不能大于50个字符！")
	Call FKFun.ShowString(J6,1,50,0,"请输入水印字符！","水印字符不能大于50个字符！")
	Call FKFun.ShowNum(J4,"水印X方位必须是数字！")
	Call FKFun.ShowNum(J5,"水印Y方位必须是数字！")
	Call FKFun.ShowNum(K1,"缩略图限宽必须是数字！")
	Call FKFun.ShowNum(K2,"缩略图限高必须是数字！")
	Call FKFun.ShowNum(K3,"编辑器图限宽必须是数字！")
	Call FKFun.ShowNum(K4,"编辑器图限高必须是数字！")
	Call FKFun.ShowNum(Fk_Jpeg_Water,"请选择是否开启水印！")
	Call FKFun.ShowNum(Fk_Jpeg_Small,"请选择是否开启缩略！")
	Fk_Jpeg_Content=J1&"|-_-|"&J2&"|-_-|"&J3&"|-_-|"&J4&"|-_-|"&J5&"|-_-|"&J6&"|-_-|7|-_-|8|-_-|9|-_-|10|-_-|11|-_-|12|-_-|13|-_-|14|-_-|15|-_-|"&K1&"|-_-|"&K2&"|-_-|"&K3&"|-_-|"&K4&"|-_-|20|-_-|21|-_-|22|-_-|23|-_-|24|-_-|25"
	Sqlstr="Select * From [Fk_Jpeg]"
	Rs.Open Sqlstr,Conn,1,3
	If Not Rs.Eof Then
		Application.Lock()
		Rs("Fk_Jpeg_Water")=Fk_Jpeg_Water
		Rs("Fk_Jpeg_Small")=Fk_Jpeg_Small
		Rs("Fk_Jpeg_Content")=Fk_Jpeg_Content
		Rs.Update
		Application.UnLock()
	Else
		Rs.Close
		Response.Write("设置丢失！")
		Call FKDB.DB_Close()
		Session.CodePage=936
		Response.End()
	End If
	Rs.Close
	Response.Write("设置成功！")
End Sub

'==========================================
'函 数 名：WaterTest()
'作    用：水印测试
'参    数：
'==========================================
Sub WaterTest()
	J1=Request.QueryString("J1")
	J2=Request.QueryString("J2")
	J3=Request.QueryString("J3")
	J4=Request.QueryString("J4")
	J5=Request.QueryString("J5")
	J6=Request.QueryString("J6")
	If J1="" Or J2="" Or J3="" Or J4="" Or J5="" Or J6="" Then
		Response.Write("设置请输入完全！")
	Else
		Call FKFun.DoWater("Images/Test.jpg","Images/Test2.jpg",J1,J2,J3,J4,J5,J6)
		Response.Write("<img src=Images/Test2.jpg />")
	End If
End Sub
%><!--#Include File="../Code.asp"-->