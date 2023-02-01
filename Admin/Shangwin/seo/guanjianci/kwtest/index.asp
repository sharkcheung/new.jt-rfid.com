<body id="body"><%
wd=request("d")
if wd<>"" then
	openurl="http://" & Request.ServerVariables("server_name") & Request.ServerVariables("script_name")
	openurl=replace(openurl,filename,wd&".htm")
	
	Set fs = server.CreateObject("scripting.filesystemobject")
	if fs.FileExists(Server.MapPath(wd&".htm")) then  
		response.redirect openurl
		'response.write "存在"
	else
		HttpUrl="http://tool.admin2.com/related/txt.php?q="&wd&""
		Call Create_File(wd&".htm")
		html=GetHttpPage(HttpUrl)
		html=replace(html,"关键字 :     收录量:    搜索量:","")
		html=replace(html,"关键字","</span></div><div><span name='kwc' class='shover' title='单击关键词即可自动添加到关键词库 ' onclick='adk(this)'>关键字")
		html=replace(html,"收录量","</span><span>收录量")
		html=replace(html,"搜索量","</span><span>搜索量")
		html=replace(html,"关键字 :","")
		html=replace(html,"收录量:","")
		html=replace(html,"搜索量:","")
		html=replace(html,"","")
		html=replace(html,wd,"<b>"&wd&"</b>")
		html="<html oncontextmenu='return false'><head><meta content='text/html; charset=gb2312' http-equiv='Content-Type' /><script src='addkw.js' type='text/javascript'></script><link href='keycss.css' rel='stylesheet' type='text/css' /><script type='text/javascript' src='ymPrompt_ex.js'></script></head><body><div id='kwlist'><div id='tishi'></div><div><span><strong>关键词</strong></span><span><strong>收录量</strong></span><span><strong>日搜索量</strong>"&html&"</div></div></body></html>"
		Call Get_File(html,wd&".htm")
		response.redirect openurl
		'response.write "创建"
	end if
    set fso = Nothing

else
	response.write "<html oncontextmenu='return false'><script src='addkw.js' type='text/javascript'></script><link href='keycss.css' rel='stylesheet' type='text/css' /><div id='kwlist'><span id='tishi'></span></div></html>"
end if


'获取当前文件名
Function filename()
	Dim arrName,postion
	fileName=Request.ServerVariables("script_name")
	postion=InstrRev(fileName,"/")+1
	fileName=Mid(fileName,postion)
	If InStr(fileName,"?")>0 Then
		arrName=fileName
		arrName=Split(arrName,"?")
		filename=arrName(0)
	End If
End Function

'读文件
Function Read_File(FilePath)
    Set Fso = Server.Createobject("Scripting.FileSystemObject")
    Path = Server.MapPath(FilePath)
    set file = fso.opentextfile(path, 1)
    do until file.AtEndOfStream
        Get_String = file.ReadLine
        Get_String = Trim(Get_String)
        If Get_String="" or IsNull(Get_String) Then
            Get_String = "error"
        End If
        Response.write("Get_String: " & Get_String & "<br/>")
    loop
    file.close
    set file = nothing
    set fso = Nothing
End Function

'创建文件
Function Create_File(FilePath)

    Set Fso = Server.CreateObject("Scripting.FileSystemObject")    
    FilePath = Server.Mappath(FilePath)
    If Not Fso.FileExists(FilePath) then 
        Set CF=Fso.CreateTextFile(FilePath,True)
        CF.Close
        'Response.End
    End If
    Set Fso = Nothing

End Function 


'读取远程网址
Function GetHttpPage(HttpUrl)
    If IsNull(HttpUrl)=True or Len(HttpUrl)<18 or HttpUrl="$False$" Then
    GetHttpPage="$False$"
    Exit Function
    End If
    Dim Http
    Set Http=server.createobject("MSXML2.XMLHTTP")
    Http.open "GET",HttpUrl,False
    Http.Send()
    If Http.Readystate<>4 then
    Set Http=Nothing 
    GetHttpPage="$False$"
    Exit function
    End if
    if Len(Http.responseBody)<10 then
    	GetHTTPPage=""
    else
   		GetHTTPPage=bytesToBSTR(Http.responseBody,"GB2312")
    end if
    Set Http=Nothing
    If Err.number<>0 then
    Err.Clear
    End If
End Function

'将获取的源码转换为中文
Function BytesToBstr(Body,Cset)
    Dim Objstream
    Set Objstream = Server.CreateObject("adodb.stream")
    objstream.Type = 1
    objstream.Mode =3
    objstream.Open
    objstream.Write body
    objstream.Position = 0
    objstream.Type = 2
    objstream.Charset = Cset
    BytesToBstr = objstream.ReadText 
    objstream.Close
    set objstream = nothing
End Function


'写文件
Function Get_File(Get_Url,path)

    Set fso = Server.CreateObject("Scripting.FileSystemObject")
    path = server.mappath(path)
    set zy=fso.OpenTextFile(path,8,false)
    zy.writeline Get_Url 
    zy.close
    set zy = nothing
    set fso = Nothing

End Function


%>