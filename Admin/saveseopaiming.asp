<!--#Include File="Include.asp"--><%
'==========================================
'文 件 名：saveseopaiming.asp
'文件用途：保存SEO关键词排名结果
dim paimingjieguo,paimingkeywords
paimingjieguo=Trim(request("paimingjieguo"))
paimingkeywords=FKFun.HTMLEncode(Trim(request("paimingkeywords")))

if paimingjieguo<>"Baidu:0 Google:0" then
Sqlstr="Select * From [keywordSV] Where SVkeywords='"&paimingkeywords&"'"
	Rs.Open Sqlstr,Conn,1,3
	If Rs.Eof Then
		Application.Lock()
		Rs.AddNew()
		Rs("SVkeywords")=paimingkeywords
		Rs("SVpaiming")=paimingjieguo
		Rs.Update()
		Application.UnLock()
		'Response.Write("关键词排名查询结果添加成功！")
	Else
		Application.Lock()
		Rs("SVkeywords")=paimingkeywords
		Rs("SVpaiming")=paimingjieguo
		Rs.Update()
		Application.UnLock()
		'Response.Write("关键词排名查询结果修改成功！")
	End If
	Rs.Close
end if
%><!--#Include File="../Code.asp"-->