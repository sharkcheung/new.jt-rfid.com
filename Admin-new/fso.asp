<%
response.charset="utf-8"
session.codepage=65001
private Function ReadFromTextFile(FileUrl,CharSet)
    dim str
    set stm=server.CreateObject("adodb.stream")
    stm.Type=2 '以本模式读取
    stm.mode=3
    stm.charset=CharSet
    stm.open
    stm.loadfromfile server.MapPath(FileUrl)
    str=stm.readtext
    stm.Close
    set stm=nothing
    ReadFromTextFile=str
End Function

private Sub WriteToTextFile(FileUrl,byval Str,CharSet)
    set stm=server.CreateObject("adodb.stream")
    stm.Type=2 '以本模式读取
    stm.mode=3
    stm.charset=CharSet
    stm.open
        stm.WriteText str
    stm.SaveToFile server.MapPath(FileUrl),2
    stm.flush
    stm.Close
    set stm=nothing
End Sub

Set fs = server.CreateObject("scripting.filesystemobject")  
'--假设你想让生成的EXCEL文件做如下的存放  
filename=Server.MapPath("skeyyword.xls")
fileurl="skeyyword.xls"
arrk=ReadFromTextFile("KeyWordC.dat","utf-8") '读取模板网页文件代码
arrk=split(arrk,"|")
'--如果原来的EXCEL文件存在的话删除它  
if fs.FileExists(filename) then  
	fs.DeleteFile(filename)  
end  if
for i=0 to ubound(arrk)
	strArr=strArr & arrk(i) & chr(13)
next
WriteToTextFile "skeyyword.xls",strArr,"utf-8"
set fs=Nothing
%>