<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="c_func.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<!--<meta http-equiv="Content-Type" content="text/html; charset=gb2312" />-->
<link rel="stylesheet" href="http://cache.soso.com/wenwen/css/w_base_201012271648.css" type="text/css" media="screen"/>
<link rel="stylesheet" href="http://cache.soso.com/wenwen/css/search_201011241014.css" type="text/css" media="screen"/>
<script src="jquery-1.4.4.js" type="text/javascript"></script>
<script type="text/javascript">
   $(document).ready(function(){
      $("ol li:first-child").addClass("hide");
	  $("ol li").find("h3");
//	  $("#search_kcss a").hover(function(){
//	     $(this).addClass("hover");
//	  },function(){
//	     $(this).removeClass("hover");
//	  })
   })
</script>
<!--<link rel="stylesheet" href="http://cache.soso.com/30d/css/web/isoso9.css" />-->
<style type="text/css">
   *{margin:0px;padding:0px;}
   #c_content{width:960px; margin:0 auto;text-align:center;margin-top:10px;}
   #main_s_content{width:600px; margin:0 auto;}
   #main_s_content ol li{ line-height:150%;}
   #main_s_content .info,#main_s_content .ico_official,#main_s_content .ico_pic,#main_s_content .ico_expert{display:none;}
   #main_s_content .hide{display:none;}
   #search_kcss{width:600px; margin:0 auto;padding:10px;font-size:14px;border:solid #ccc 1px;margin-top:10px;margin-bottom:10px;}
   #search_kcss a{font-weight:bolder;padding:2px;}
   #search_kcss a:visited{color:#f00;}
   #search_kcss .acss a{color:#053E02;}
   .hover{background-color:#F9F9F9;border:#F1CC8D solid 1px;}
</style>
<title></title>
</head>

<body oncontextmenu="return false;">
<div id="c_content"><form name="frm_search" id="frm_search" method="get"><input type="text" name="sp" id="sp" /> &nbsp;<input type="submit" value="搜 索"/></form><!--<form name="frm_search" id="frm_search" method="get"><input type="text" name="w" id="w" /> &nbsp;<input type="submit" value="搜索"/></form>--></div>
<%
dim keyword,r,body
keyword = trim(Request.querystring("sp"))
a = trim(Request.querystring("a"))
if a="c" then 
   response.Cookies("sosearch_keyword")=""
end if
if keyword<>"" then
   cookies_key=trim(request.Cookies("sosearch_keyword"))
   if cookies_key="" then 
      response.Cookies("sosearch_keyword")=keyword
   else
      Arrcookies_key=split(cookies_key,"|||")
      if ubound(Arrcookies_key)>=0 then
		 blnkey=false
	     for i=0 to ubound(Arrcookies_key)
	        if keyword=Arrcookies_key(i) then
			   blnkey=true
		       exit for
		    end if
	     next
		 if blnkey=false then 
		    response.Cookies("sosearch_keyword")=cookies_key&"|||"&keyword
		 end if
      end if
'      response.Write instr(cookies_key,keyword)
'	  if instr(cookies_key,keyword)<=0 then
'         response.Cookies("sosearch_keyword")=cookies_key&"|||"&keyword
'	  end if
   end if
   cookies_key=request.Cookies("sosearch_keyword")
   Arrcookies_key=split(cookies_key,"|||")
   if ubound(Arrcookies_key)>=0 then
      response.Write "<div id=""search_kcss"">最近搜索关键词: "
	  for i=0 to ubound(Arrcookies_key)
	     response.Write "<a href=""?sp="&Arrcookies_key(i)&"&ch=w.search.sb"">"&Arrcookies_key(i)&"</a>"&" "
	  next
	  response.Write "<span class=""acss""><a href=""?a=c"" title=""点击清空关键词"">清空关键词</a></span></div>"
   end if
   cur_page=clng(Request.querystring("pg"))
   if cur_page="" or cur_page=0 then
   url="http://wenwen.soso.com/z/Search.e?sp=S"&server.URLEncode(keyword)&"&ch=w.search.sb"
   'url="http://www.soso.com/q?pid=s.idx&w="&keyword
   else
   url="http://wenwen.soso.com/z/Search.e?sp=S"&server.URLEncode(keyword)&"&sci=0&pg="&cur_page&""
   'url="http://www.soso.com/q?w="&keyword&"&lr=&sc=web&ch=w.p.b&num=10&gid=&cin=&site=&sf=0&sd=0&pg="&cur_page&""
   end if
   wstr=getHTTPPage(url,"utf-8")
'   response.Write wstr
'   response.end
   body=strCut(wstr,"<!--result list-->","<!--result side-->",2)
   
   'body=replace(body,"<span class=""yl1"" onFocus=""blur();"">预览</span>","")
   r = "(<span class=""solved_time"">.*?</span>)*"
   body = py_z_replace(body,r,"")
   body=replace(body,"<span>-</span>","")
   r = "(<a *? href=""/z/.*?.htm"">.*?</a>)*"
   body = py_z_replace(body,r,"")
   r = "(<a target=""_blank"" .*?>.*?</a>.*?字)*"
   body = py_z_replace(body,r,"")
   body = replace(body,"/z/Search.e?sp=S",CUrl(1)&"?sp=")
   body = replace(body,".htm?w=",".htm&w=")
   body = replace(body,"/z/q",ServerPath&"detail.asp?py=q")
   body = replace(body,"<script>setTimeout(setPosition,50);</script>","")
   body = replace(body,"_blank","_self")
   
   'body = replace(body,"http://baike.soso.com/ShowLemma.e",ServerPath&"detail.asp")

'   r = "(<cite>.*?</cite>-)*"
'   body = py_z_replace(body,r,"")
'   r="(<div class=""url""><a href=.*?>网页快照</a></div>)*"
'   body = py_z_replace(body,r,"")
'   r="(<div class=""result_summary"">-<span class=""preview"" id="".*?""></span></div>)*"
'   body = py_z_replace(body,r,"")
   response.write "<div id=""main_s_content"">"&body&"</div>"
   'response.write "<div id=""main_s_content"">"&body&turn_page&"</div>"

end if
%>
</body>
</html>
