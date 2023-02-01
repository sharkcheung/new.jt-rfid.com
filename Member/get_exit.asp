<%
r=request.ServerVariables("HTTP_REFERER")
response.Cookies("login")("u_id")=""
response.Cookies("login")("u_pass")=""
session("u_id")=""
session("u_pass")=""
session.Abandon()
'response.write "<script language=javascript>window.location.href='"&r&"';</script>"
response.write "<script language=javascript>window.top.location.href='/';</script>"

response.end
%>