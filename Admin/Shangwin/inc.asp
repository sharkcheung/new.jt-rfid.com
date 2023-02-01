<%
Session.CodePage=65001
Response.ContentType = "text/html"
Response.Charset = "utf-8"
%>
<!-- ymPrompt组件 -->
<script type="text/javascript" src="/admin/winskin/ymPrompt.js"></script>
<link rel="stylesheet" type="text/css" href="/admin/winskin/qq/ymPrompt.css" /> 
<!-- ymPrompt组件 -->
<script type="text/javascript" src="/Js/jquery-1.4.2.min.js"></script>
<script type="text/javascript" src="/Js/jquery.form.min.js"></script>
<script type="text/javascript" src="/Js/function.js"></script>
<%
if request.Cookies("FkAdminName")="" then
%>
<script type="text/javascript">
$(document).ready(function(){
<%
	Response.Write("  tan3(""登录状态失效，请重新登录！"");")
%>

});
</script>
<%
	response.end
end if
%>

<%
'****************公共函数区******************************

'截取字符串,1.包括前后字符串，2.不包括前后字符串
Function strCut(strContent,StartStr,EndStr,CutType)
Dim S1,S2
On Error Resume Next
Select Case CutType
Case 1
  S1 = InStr(strContent,StartStr)
  S2 = InStr(S1,strContent,EndStr)+Len(EndStr)
Case 2
  S1 = InStr(strContent,StartStr)+Len(StartStr)
  S2 = InStr(S1,strContent,EndStr)
End Select
If Err Then
  strCute = "<p align='center' ><font size=-1>截取字符串出错.</font></p>"
  Err.Clear
  Exit Function
Else
  strCut = Mid(strContent,S1,S2-S1)
End If
End Function


%>