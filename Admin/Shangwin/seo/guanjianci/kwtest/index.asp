<body id="body"><%
wd=request("d")
if wd<>"" then
	openurl="http://" & Request.ServerVariables("server_name") & Request.ServerVariables("script_name")
	openurl=replace(openurl,filename,wd&".htm")
	
	Set fs = server.CreateObject("scripting.filesystemobject")
	if fs.FileExists(Server.MapPath(wd&".htm")) then  
		response.redirect openurl
		'response.write "����"
	else
		HttpUrl="http://tool.admin2.com/related/txt.php?q="&wd&""
		Call Create_File(wd&".htm")
		html=GetHttpPage(HttpUrl)
		html=replace(html,"�ؼ��� :     ��¼��:    ������:","")
		html=replace(html,"�ؼ���","</span></div><div><span name='kwc' class='shover' title='�����ؼ��ʼ����Զ���ӵ��ؼ��ʿ� ' onclick='adk(this)'>�ؼ���")
		html=replace(html,"��¼��","</span><span>��¼��")
		html=replace(html,"������","</span><span>������")
		html=replace(html,"�ؼ��� :","")
		html=replace(html,"��¼��:","")
		html=replace(html,"������:","")
		html=replace(html,"","")
		html=replace(html,wd,"<b>"&wd&"</b>")
		html="<html oncontextmenu='return false'><head><meta content='text/html; charset=gb2312' http-equiv='Content-Type' /><script src='addkw.js' type='text/javascript'></script><link href='keycss.css' rel='stylesheet' type='text/css' /><script type='text/javascript' src='ymPrompt_ex.js'></script></head><body><div id='kwlist'><div id='tishi'></div><div><span><strong>�ؼ���</strong></span><span><strong>��¼��</strong></span><span><strong>��������</strong>"&html&"</div></div></body></html>"
		Call Get_File(html,wd&".htm")
		response.redirect openurl
		'response.write "����"
	end if
    set fso = Nothing

else
	response.write "<html oncontextmenu='return false'><script src='addkw.js' type='text/javascript'></script><link href='keycss.css' rel='stylesheet' type='text/css' /><div id='kwlist'><span id='tishi'></span></div></html>"
end if


'��ȡ��ǰ�ļ���
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

'���ļ�
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

'�����ļ�
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


'��ȡԶ����ַ
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

'����ȡ��Դ��ת��Ϊ����
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


'д�ļ�
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