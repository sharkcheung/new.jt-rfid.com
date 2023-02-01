<!--#Include File="Include.asp"-->
<!--#Include File="Class/cls_showpage.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<style type="text/css">
<!--
*{font-size:12px;font-family:"lucida grande",Verdana,lucida,arial,helvetica,宋体, sans-serif;}
body{padding:5px; }
#search_content{width:98%;}
ul,li{list-style:none;padding:0px;margin:0px;}
.news-intro-list li {
	float: left;
	clear: both;
	width: 100%;
}
.news-intro-title {
	font-size: 14px;
	font-weight: bold;
	height: 30px;
	line-height: 30px;
	border-bottom: dashed 1px #ccc;
}
.news-intro-title a {
	text-decoration: none;
	float: left;
	color: #004d81;
}
.news-intro-title a:hover {
	color: #ff5801;
	text-decoration: underline;
}
.news-intro-title span {
	color: #B6C6C9;
	font-size: 12px;
	font-weight: normal;
	float: right;
}
.td_showpage{padding:5px 0px;}
.td_showpage a{color:#004d81;text-decoration:none;}
.td_showpage a:visited{color:#004d81;text-decoration:none;}
.td_showpage a:hover{color:#ff5801;text-decoration:none;}
.page_num{border:#ccc solid 1px;padding:2px 5px;margin-left:1px;}
.li_top{padding-left:10px;color:#333;font-size:14px; }
.li_top span{padding-left:4px;color:#ff5801;font-size:16px; font-weight:bolder;}
form{display:inline;padding-left:50px;}
-->
</style>
<title>搜索页面</title>
  	<script>
	function checkQuery()
	{
		var cc = document.getElementById("frm_s");
		if(cc.M_key.value=="" || cc.M_key.value=='请输入关键词')
		{
			alert("请输入关键字!");
			cc.M_key.focus();
			return false;
		}
		return true;
	}
	</script>
</head>
<body>
<div id="search_content">
  <ul class="news-intro-list">
    <%dim main_type,main_key,search_sql,search_rs,type_name,re_url,condition
re_url=request.servervariables( "HTTP_REFERER")
main_type=clng(request.QueryString("M_type"))
main_key=ChkQueryStr(trim(request.QueryString("M_key")))
if main_key="请输入关键词" or main_key="" then
	response.Write "<script language=javascript>alert('请输入关键词！');parent.window.location.href='"&re_url&"'</script>"
   	response.end
else
   	if main_type=0 then
   		condition=" and b.Fk_Module_Level=84 "
	else
   		condition=" and b.Fk_Module_Id="&main_type&" "
	end if
    
	dim type_rs
         search_sql="select a.Fk_Product_Id,a.Fk_Product_Title,a.Fk_Product_Time,b.Fk_Module_Dir from Fk_Product a inner join Fk_Module b on a.Fk_Product_Module=b.Fk_Module_Id where 1=1"
		 if main_key<>"" then
	     search_sql=search_sql&" and Fk_Product_Title like '%"&main_key&"%' or Fk_Product_Keyword like '%"&main_key&"%' or Fk_Product_Description like '%"&main_key&"%' or Fk_Product_Content like '%"&main_key&"%'"
		 end if
		 search_sql=search_sql&" order by Fk_Product_Click desc,Fk_Product_Time desc,Fk_Product_Id desc"
set type_rs=conn.execute(search_sql)
      response.Write "<li class=""li_top"">当前搜索类型: <span>"&type_name&"</span> &nbsp; &nbsp; 关键词: <span>"&main_key&"</span><form action="""" method=""get"" id=""frm_s"">"& _
			      " <input type=""text"" name=""M_key"" value=""请输入关键词"" onClick=""this.value='';""/> <input type=""submit"" name=""btn_m"" value=""搜 索""  onClick=""return checkQuery();submit();"" />"& _
			   "</form></li>"
'   set search_rs=server.CreateObject("adodb.recordset")
'   search_rs.open search_sql,conn,1,3
'   if search_rs.eof then
'      	response.Write "找不到您要搜索的信息1s!"
'   else
'   	response.Write search_rs.recordcount
'   end if
'response.end
      dim mypage
      Set mypage = new xdownpage '/创建对象
      mypage.getconn = conn '/得到数据库连接
      mypage.getsql = search_sql '/sql语句
      mypage.pagesize = 25 '/设置每一页的记录条数据为5条
      set search_rs = mypage.getrs() '/返回Recordset
	  if search_rs.eof then
         response.Write "找不到您要搜索的信息!"
	  else
      For I = 1 To mypage.pagesize
	  if not search_rs.eof then
	     response.Write "<li class=""news-intro-title""><span>"&search_rs(2)&"</span><a href=""/?"&search_rs(3)&"/"&search_rs(0)&".html"" target=""_blank"">"&search_rs(1)&"</a></li>"
      search_rs.movenext
	  else
	     exit for
	  end if
	  next
	  end if
'   end if
   mypage.showpage()
   Set mypage=nothing
   search_rs.close
   set search_rs=nothing
end if
'函数名：ChkQueryStr
'作 用：过虑查询的非法字符
'参 数：str ----原字符串
'返回值：过滤后的字符
'================================================
Public Function ChkQueryStr(ByVal str)
On Error Resume Next
If IsNull(str) Then
ChkQueryStr = ""
Exit Function
End If
str = Replace(str, "!", "")
str = Replace(str, "]", "")
str = Replace(str, "[", "")
str = Replace(str, ")", "")
str = Replace(str, "(", "")
str = Replace(str, "|", "")
str = Replace(str, "+", "")
str = Replace(str, "=", "")
str = Replace(str, "'", "''")
str = Replace(str, "%", "")
str = Replace(str, "&", "")
str = Replace(str, "#", "")
str = Replace(str, "^", "")
str = Replace(str, " ", "")
str = Replace(str, ",", "")
str = Replace(str, ".", "")
str = Replace(str, ".", "")
str = Replace(str, Chr(37), "")
str = Replace(str, Chr(0), "")
ChkQueryStr = str
End Function
%>
  </ul>
</div>
<!--#Include File="Code.asp"-->
</body>
</html>
