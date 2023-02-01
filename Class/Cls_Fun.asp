<%
'==========================================
'文 件 名：Cls_Fun.asp
'文件用途：常规函数类
'==========================================

Class Cls_Fun
	Private x,y,ii
	'==============================
	'函 数 名：AlertInfo
	'作    用：错误显示函数
	'参    数：错误提示内容InfoStr，转向页面GoUrl
	'==============================
	Public Function AlertInfo(InfoStr,GoUrl)
		If GoUrl="1" Then
			Response.Write "<Script>alert('"& InfoStr &"');location.href='javascript:history.go(-1)';</Script>"
		Else
			Response.Write "<Script>alert('"& InfoStr &"');location.href='"& GoUrl &"';</Script>"
		End If
		Call FKDB.DB_Close()
		Session.CodePage=936
		Response.End()
	End Function
	
	'==============================
	'函 数 名：HTMLEncode
	'作    用：字符转换函数
	'参    数：需要转换的文本fString
	'==============================
	Public Function HTMLEncode(fString)
		If Not IsNull(fString) Then
			fString = replace(fString, ">", "&gt;")
			fString = replace(fString, "<", "&lt;")
			fString = Replace(fString, CHR(32), " ")		
			fString = Replace(fString, CHR(34), "&quot;")
			fString = Replace(fString, CHR(39), "&#39;")
			fString = Replace(fString, CHR(9), "&nbsp;")
			fString = Replace(fString, CHR(13), "")
			fString = Replace(fString, CHR(10) & CHR(10), "<p></p> ")
			fString = Replace(fString, CHR(10), "<br /> ")
			HTMLEncode = fString
		End If
	End Function
	
	'==============================
	'函 数 名：HTMLDncode
	'作    用：字符转回函数
	'参    数：需要转换的文本fString
	'==============================
	Public Function HTMLDncode(fString)
		If Not IsNull(fString) Then
			fString = Replace(fString, "&gt;",">" )
			fString = Replace(fString, "&lt;", "<")
			fString = Replace(fString, " ", CHR(32))
			fString = Replace(fString, "&nbsp;", CHR(9))
			fString = Replace(fString, "&quot;", CHR(34))
			fString = Replace(fString, "&#39;", CHR(39))
			fString = Replace(fString, "", CHR(13))
			fString = Replace(fString, "<p></p> ",CHR(10) & CHR(10) )
			fString = Replace(fString, "<br /> ",CHR(10) )
			HTMLDncode = fString
		End If
	End Function
	
	'==============================
	'函 数 名：AlertNum
	'作    用：判断是否是数字（验证字符，不为数字时的提示）
	'参    数：需进行判断的文本CheckStr，错误提示ErrStr
	'==============================
	Public Function AlertNum(CheckStr,ErrStr)
		If Not IsNumeric(CheckStr) or CheckStr="" Then
			Call AlertInfo(ErrStr,"1")
		End If
	End Function

	'==============================
	'函 数 名：AlertString
	'作    用：判断字符串长度
	'参    数：
	'需进行判断的文本CheckStr
	'限定最短ShortLen
	'限定最长LongLen
	'验证类型CheckType（0两头限制，1限制最短，2限制最长）
	'过短提示LongStr
	'过长提示LongStr，
	'==============================
	Public Function AlertString(CheckStr,ShortLen,LongLen,CheckType,ShortErr,LongErr)
		If (CheckType=0 Or CheckType=1) And StringLength(CheckStr)<ShortLen Then
			Call AlertInfo(ShortErr,"1")
		End If
		If (CheckType=0 Or CheckType=2) And StringLength(CheckStr)>LongLen Then
			Call AlertInfo(LongErr,"1")
		End If
	End Function
	
	'==============================
	'函 数 名：ShowNum
	'作    用：判断是否是数字（验证字符，不为数字时的提示）
	'参    数：需进行判断的文本CheckStr，错误提示ErrStr
	'==============================
	Public Function ShowNum(CheckStr,ErrStr)
		If Not IsNumeric(CheckStr) or CheckStr="" Then
			Response.Write(ErrStr)
			Call FKDB.DB_Close()
			Session.CodePage=936
			Response.End()
		End If
	End Function

	'==============================
	'函 数 名：ShowString
	'作    用：判断字符串长度
	'参    数：
	'需进行判断的文本CheckStr
	'限定最短ShortLen
	'限定最长LongLen
	'验证类型CheckType（0两头限制，1限制最短，2限制最长）
	'过短提示LongStr
	'过长提示LongStr，
	'==============================
	Public Function ShowString(CheckStr,ShortLen,LongLen,CheckType,ShortErr,LongErr)
		If (CheckType=0 Or CheckType=1) And StringLength(CheckStr)<ShortLen Then
			Response.Write(ShortErr)
			Call FKDB.DB_Close()
			Response.End()
		End If
		If (CheckType=0 Or CheckType=2) And StringLength(CheckStr)>LongLen Then
			Response.Write(LongErr)
			Call FKDB.DB_Close()
			Response.End()
		End If
	End Function
	
	'==============================
	'函 数 名：StringLength
	'作    用：判断字符串长度
	'参    数：需进行判断的文本Txt
	'==============================
	Private Function StringLength(Txt)
		Txt=Trim(Txt)
		x=Len(Txt)
		y=0
		For ii = 1 To x
			If Asc(Mid(Txt,ii,1))<=2 or Asc(Mid(Txt,ii,1))>255 Then
				y=y + 2
			Else
				y=y + 1
			End If
		Next
		StringLength=y
	End Function
	
	'==============================
	'函 数 名：BeSelect
	'作    用：判断select选项选中
	'参    数：Select1,Select2
	'==============================
	Public Function BeSelect(Select1,Select2)
		If Select1=Select2 Then
			BeSelect=" selected='selected'"
		End If
	End Function
	
	'==============================
	'函 数 名：BeCheck
	'作    用：判断Check选项选中
	'参    数：Check1,Check2
	'==============================
	Public Function BeCheck(Check1,Check2)
		If Check1=Check2 Then
			BeCheck=" checked='checked'"
		End If
	End Function
	
	'==============================
	'函 数 名：CheckModule
	'作    用：判断模块类型，输出名称
	'参    数：要判断的类型ModuleId
	'==============================
	Public Function CheckModule(ModuleId)
		For i=0 To UBound(FKModuleId)
			If ModuleId=Clng(FKModuleId(i)) Then
				CheckModule=FKModuleName(i)
				Exit Function
			End If
		Next
	End Function	

	'==============================
	'函 数 名：ShowPageCode
	'作    用：显示页码
	'参    数：链接PageUrl，当前页Nows，记录数AllCount，每页数量Sizes，总页数AllPage
	'==============================
	Public Function ShowPageCode(PageUrl,Nows,AllCount,Sizes,AllPage)
		If Nows>1 Then
			Response.Write("<a href=""javascript:void(0);"" onclick=""SetRContent('MainRight','"&PageUrl&"1');return false"">第一页</a>")
			Response.Write("&nbsp;")
			Response.Write("<a href=""javascript:void(0);"" onclick=""SetRContent('MainRight','"&PageUrl&(Nows-1)&"');return false"">上一页</a>")
		Else
			Response.Write("第一页")
			Response.Write("&nbsp;")
			Response.Write("上一页")
		End If
		Response.Write("&nbsp;")
		If AllPage>Nows Then
			Response.Write("<a href=""javascript:void(0);"" onclick=""SetRContent('MainRight','"&PageUrl&(Nows+1)&"');return false"">下一页</a>")
			Response.Write("&nbsp;")
			Response.Write("<a href=""javascript:void(0);"" onclick=""SetRContent('MainRight','"&PageUrl&AllPage&"');return false"">尾页</a>")
		Else
			Response.Write("下一页")
			Response.Write("&nbsp;")
			Response.Write("尾页")
		End If
		Response.Write("&nbsp;"&Sizes&"条/页&nbsp;共"&AllPage&"页/"&AllCount&"条&nbsp;当前第"&Nows&"页&nbsp;")
		Response.Write("<select name=""Change_Page"" id=""Change_Page"" onChange=""SetRContent('MainRight','"&PageUrl&"'+this.options[this.selectedIndex].value);"">")
		For i=1 To AllPage
			If i=Nows Then
				Response.Write("<option value="""&i&""" selected=""selected"">第"&i&"页</option>")
			Else
				Response.Write("<option value="""&i&""">第"&i&"页</option>")
			End If
		Next
      	Response.Write("</select>")
	End Function

	'==============================
	'函 数 名：ShowPageCodeRelated
	'作    用：显示页码
	'参    数：链接PageUrl，当前页Nows，记录数AllCount，每页数量Sizes，总页数AllPage
	'==============================
	Public Function ShowPageCodeRelated(PageUrl,Nows,AllCount,Sizes,AllPage)
		If Nows>1 Then
			Response.Write("<a href=""javascript:void(0);"" onclick=""SetRContentRelated('ListContentRelated','"&PageUrl&"1');return false"">第一页</a>")
			Response.Write("&nbsp;")
			Response.Write("<a href=""javascript:void(0);"" onclick=""SetRContentRelated('ListContentRelated','"&PageUrl&(Nows-1)&"');return false"">上一页</a>")
		Else
			Response.Write("第一页")
			Response.Write("&nbsp;")
			Response.Write("上一页")
		End If
		Response.Write("&nbsp;")
		If AllPage>Nows Then
			Response.Write("<a href=""javascript:void(0);"" onclick=""SetRContentRelated('ListContentRelated','"&PageUrl&(Nows+1)&"');return false"">下一页</a>")
			Response.Write("&nbsp;")
			Response.Write("<a href=""javascript:void(0);"" onclick=""SetRContentRelated('ListContentRelated','"&PageUrl&AllPage&"');return false"">尾页</a>")
		Else
			Response.Write("下一页")
			Response.Write("&nbsp;")
			Response.Write("尾页")
		End If
		Response.Write("&nbsp;"&Sizes&"条/页&nbsp;共"&AllPage&"页/"&AllCount&"条&nbsp;当前第"&Nows&"页&nbsp;")
		Response.Write("<select name=""Change_Page"" id=""Change_Page"" onChange=""SetRContentRelated('ListContentRelated','"&PageUrl&"'+this.options[this.selectedIndex].value);"">")
		For i=1 To AllPage
			If i=Nows Then
				Response.Write("<option value="""&i&""" selected=""selected"">第"&i&"页</option>")
			Else
				Response.Write("<option value="""&i&""">第"&i&"页</option>")
			End If
		Next
      	Response.Write("</select>")
	End Function

	'==============================
	'函 数 名：ShowPageCode
	'作    用：显示页码
	'参    数：链接PageUrl，当前页Nows，记录数AllCount，每页数量Sizes，总页数AllPage
	'==============================
	Public Function ShowPaper(PageUrl,Nows,AllCount,Sizes,AllPage)
		If Nows>1 Then
			Response.Write("<a href=""javascript:void(0);"" onclick=""ShowBox('"&PageUrl&"1');return false"">第一页</a>")
			Response.Write("&nbsp;")
			Response.Write("<a href=""javascript:void(0);"" onclick=""ShowBox('"&PageUrl&(Nows-1)&"');return false"">上一页</a>")
		Else
			Response.Write("第一页")
			Response.Write("&nbsp;")
			Response.Write("上一页")
		End If
		Response.Write("&nbsp;")
		If AllPage>Nows Then
			Response.Write("<a href=""javascript:void(0);"" onclick=""ShowBox('"&PageUrl&(Nows+1)&"');return false"">下一页</a>")
			Response.Write("&nbsp;")
			Response.Write("<a href=""javascript:void(0);"" onclick=""ShowBox('"&PageUrl&AllPage&"');return false"">尾页</a>")
		Else
			Response.Write("下一页")
			Response.Write("&nbsp;")
			Response.Write("尾页")
		End If
		Response.Write("&nbsp;"&Sizes&"条/页&nbsp;共"&AllPage&"页/"&AllCount&"条&nbsp;当前第"&Nows&"页&nbsp;")
		Response.Write("<select name=""Change_Page"" id=""Change_Page"" onChange=""ShowBox('"&PageUrl&"'+this.options[this.selectedIndex].value);"">")
		For i=1 To AllPage
			If i=Nows Then
				Response.Write("<option value="""&i&""" selected=""selected"">第"&i&"页</option>")
			Else
				Response.Write("<option value="""&i&""">第"&i&"页</option>")
			End If
		Next
      	Response.Write("</select>")
	End Function

	'==============================
	'函 数 名：GetNowUrl
	'作    用：返回当前网址
	'参    数：
	'==============================
	Public Function GetNowUrl()
		GetNowUrl=Request.ServerVariables("Script_Name")&"?"&Request.ServerVariables("QUERY_STRING")
	End Function

	'==============================
	'函 数 名：CheckLimit
	'作    用：判断权限
	'参    数：需要字符LimitStr
	'==============================
	Public Function CheckLimit(LimitStr)
		If Request.Cookies("FkAdminLimitId")>0 Then
			CheckLimit=False
			TempArr=Split(LimitStr,"|")
			For Each Temp In TempArr
				If Instr(Request.Cookies("FkAdminLimit"),","&Temp&",")>0 Then
					CheckLimit=True
				End If
			Next
		Else
			CheckLimit=True
		End If
	End Function

	'==============================
	'函 数 名：ReplaceTest
	'作    用：正则表达式，替换字符串
	'参    数：规则patrn，要替换的字符串Str，替换为字符串replStr
	'==============================
	Public Function ReplaceTest(patrn,replStr,Str)
		Dim regEx
		Set regEx = New RegExp
		regEx.Pattern = patrn
		regEx.IgnoreCase = True
		regEx.Global = True 
		ReplaceTest = regEx.Replace(Str,replStr)
	End Function 

	'==============================
	'函 数 名：RegExpTest
	'作    用：正则表达式，获取字符串
	'参    数：源字符串patrn，规则strng
	'==============================
	Public Function RegExpTest(patrn, strng)
		Dim regEx, Matchs, Matches, RetStr
		Set regEx = New RegExp 
		regEx.Pattern = patrn 
		regEx.IgnoreCase = True
		regEx.Global = True
		Set Matches = regEx.Execute(strng) 
		For Each Matchs in Matches
			RetStr = RetStr & Matchs.Value & "|-_-|"
		Next 
		RegExpTest = RetStr 
	End Function 
	
	'==============================
	'函 数 名：NoTrash
	'作    用：垃圾信息强力判断
	'参    数：要判断的信息TryStr
	'==============================
	Public Function NoTrash(TryStr)
		Dim HttpCount
		TryStr=LCase(TryStr)
		HttpCount=Clng(((Len(TryStr)-Len(Replace(TryStr,"http://","")))/7))
		If HttpCount>3 Then
			Call AlertInfo(arrTips(9),"1")
		End If
		If StringLength(TryStr)<=(Len(TryStr)*1.3) Then
			Call AlertInfo(arrTips(9),"1")
		End If
	End Function
	
	'==============================
	'函 数 名：SmsGo
	'作    用：发送短信
	'参    数：
	'==============================
	Public Function SmsGo(GoText)
		If FetionNum<>"" And FetionPass<>"" Then
			SmsGo=GetHttpPage("http://sms.api.bz/fetion.php?username="&FetionNum&"&password="&FetionPass&"&sendto="&FetionNum&"&message="&GoText&"","UTF-8")
			SmsGo="信息已发送"
		Else
			SmsGo="飞信未设！"
		End If
	End Function

	'==============================
	'函 数 名：GetHttpPage
	'作    用：获取页面源代码函数
	'参    数：网址HttpUrl，编码Cset
	'==============================
	Public Function GetHttpPage(HttpUrl,Cset)
		If IsNull(HttpUrl)=True Or HttpUrl="" Then
			GetHttpPage="A站点维护中！"
			Exit Function
		End If
		On Error Resume Next
		Dim Http
		Set Http=server.createobject("MSX"&"ML2.XML"&"HTTP")
		Http.open "GET",HttpUrl,False
		Http.Send()
		If Http.Readystate<>4 then
			Set Http=Nothing
			GetHttpPage="B站点维护中！"
			Exit function
		End if
		GetHttpPage=BytesToBSTR(Http.responseBody,Cset)
		Set Http=Nothing
		If Err.number<>0 then
			Err.Clear
			GetHttpPage="C站点维护中！"
			Exit function
		End If
	End Function

	'==============================
	'函 数 名：BytesToBstr
	'作    用：转换编码函数
	'参    数：字符串Body，编码Cset
	'==============================
	Private Function BytesToBstr(Body,Cset)
		Dim Objstream
		Set Objstream = Server.CreateObject("ado"&"d"&"b.st"&"re"&"am")
		Objstream.Type = 1
		Objstream.Mode =3
		Objstream.Open
		Objstream.Write body
		Objstream.Position = 0
		Objstream.Type = 2
		Objstream.Charset = Cset
		BytesToBstr = Objstream.ReadText 
		Objstream.Close
		set Objstream = nothing
	End Function
	
	'==============================
	'函 数 名：HtmlToJs
	'作    用：HTML转JS
	'参    数：字符串CStrs
	'==============================
	Public Function HtmlToJs(CStrs)
		Dim ToJs
		CStrs=Replace(CStrs,Chr(10),"") 
		CStrs=Replace(CStrs,Chr(32)&Chr(32),"") 
		CStrs=Split(CStrs,Chr(13))
		ToJs=""
		For i=0 To UBound(CStrs) 
		If Trim(CStrs(i)) <> "" Then 
			CStrs(i)= Replace(CStrs(i),Chr(34),Chr(39)) 
			ToJs=ToJs&"document.write("&Chr(34)&CStrs(i)&Chr(34)&");"&Chr(10) 
		End If 
		Next
		HtmlToJs=ToJs
	End Function
	
	'==============================
	'函 数 名：GetAdminDir
	'作    用：获取管理目录
	'参    数：
	'==============================
	Public Function GetAdminDir()
		If Request.ServerVariables("SERVER_PORT")<>"80" Then
			GetAdminDir="http://"&Request.ServerVariables("SERVER_NAME")&":"&Request.ServerVariables("SERVER_PORT")&Request.ServerVariables("URL")
		Else
			GetAdminDir="http://"&Request.ServerVariables("SERVER_NAME")&Request.ServerVariables("URL")
		End If
		GetAdminDir=Left(GetAdminDir,InstrRev(GetAdminDir,"/")-1)
		GetAdminDir=LCase(Mid(GetAdminDir,InstrRev(GetAdminDir,"/")+1))
	End Function
	
	Public Function UnEscape(Str)
		dim r,s,c 
		s="" 
		For r=1 to Len(Str) 
			c=Mid(Str,r,1) 
			If Mid(Str,r,2)="%u" and r<=Len(Str)-5 Then 
				If IsNumeric("&H" & Mid(Str,r+2,4)) Then 
					s = s & CHRW(CInt("&H" & Mid(Str,r+2,4))) 
					r = r+5 
				Else 
					s = s & c 
				End If 
			ElseIf c="%" and r<=Len(Str)-2 Then 
				If IsNumeric("&H" & Mid(Str,r+1,2)) Then 
					s = s & CHRW(CInt("&H" & Mid(Str,r+1,2))) 
					r = r+2 
				Else 
					s = s & c 
				End If 
			Else 
				s = s & c 
			End If 
		Next 
		UnEscape = s 
	End Function 
	
	'==============================
	'函 数 名：RemoveHTML
	'作    用：过滤HTML
	'参    数：
	'==============================
	Public Function RemoveHTML(strHTML)
		Dim objRegExp, Match, Matches 
		Set objRegExp = New Regexp 
		objRegExp.IgnoreCase = True 
		objRegExp.Global = True 
		'取闭合的<> 
		objRegExp.Pattern = "<.+?>" 
		'进行匹配 
		Set Matches = objRegExp.Execute(strHTML) 
		' 遍历匹配集合，并替换掉匹配的项目 
		For Each Match in Matches 
			strHtml=Replace(strHTML,Match.Value,"") 
		Next 
		'取特殊字符
		objRegExp.Pattern = "\&.+?;" 
		'进行匹配 
		Set Matches = objRegExp.Execute(strHTML) 
		' 遍历匹配集合，并替换掉匹配的项目 
		For Each Match in Matches 
			strHtml=Replace(strHTML,Match.Value,"") 
		Next 
		RemoveHTML=strHTML 
		Set objRegExp = Nothing 
	End Function
	
	'==============================
	'函 数 名：IsObjInstalled
	'作    用：判断组件是否安装了
	'参    数：组件名strClassString
	'==============================
	Public Function IsObjInstalled(strClassString)
		On Error Resume Next
		IsObjInstalled = False 
		Err.Clear
		Dim xTestObj
		Set xTestObj = Server.CreateObject(strClassString)
		If 0 = Err Then IsObjInstalled = True 
		Set xTestObj = Nothing
		Err.Clear
	End Function
	
	'==============================
	'函 数 名：DoWater
	'作    用：加入水印
	'参    数：水印图片JImg,保存路径JSImg,颜色JColor,字体JFont,加粗JStrong,X位置JpX,Y位置JpY,文字JText
	'==============================
	Function DoWater(JImg,JSImg,JColor,JFont,JStrong,JpX,JpY,JText)
		Dim JpegObj
		Set JpegObj = Server.CreateObject("Persits.Jpeg")
		JpegObj.Open Server.MapPath(JImg)
		JpegObj.Canvas.Font.Color = JColor
		JpegObj.Canvas.Font.Family = JFont
		JpegObj.Canvas.Font.Bold = CBool(JStrong)
		JpegObj.Canvas.Print JpX, JpY, JText
		JpegObj.Save Server.MapPath(JSImg)
		Set JpegObj = Nothing
	End Function
	
	'==============================
	'函 数 名：DoSmall
	'作    用：图片缩略
	'参    数：缩略图片JImg
	'==============================
	Function DoSmall(JImg,JSImg,JWidth,JHeight)
		Dim JpegObj
		Set JpegObj=Server.CreateObject("Persits.Jpeg")
		JpegObj.Open Server.MapPath(JImg)
		JpegObj.Width=JWidth
		JpegObj.Height=JHeight
		JpegObj.Save Server.MapPath(JSImg)
		JpegObj.Close
		Set JpegObj = Nothing
	End Function
	
	Function getIP()
		Dim sIPAddress, sHTTP_X_FORWARDED_FOR
		 sHTTP_X_FORWARDED_FOR = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
		If sHTTP_X_FORWARDED_FOR = "" Or InStr(sHTTP_X_FORWARDED_FOR, "unknown") > 0 Then
			sIPAddress = Request.ServerVariables("REMOTE_ADDR")
		ElseIf InStr(sHTTP_X_FORWARDED_FOR, ",") > 0 Then
			sIPAddress = Mid(sHTTP_X_FORWARDED_FOR, 1, InStr(sHTTP_X_FORWARDED_FOR, ",") -1)
		ElseIf InStr(sHTTP_X_FORWARDED_FOR, ";") > 0 Then
			sIPAddress = Mid(sHTTP_X_FORWARDED_FOR, 1, InStr(sHTTP_X_FORWARDED_FOR, ";") -1)
		Else
			sIPAddress = sHTTP_X_FORWARDED_FOR
		End If
		getIP = Trim(Mid(sIPAddress, 1, 15))
	End Function
	
	
End Class
%>
