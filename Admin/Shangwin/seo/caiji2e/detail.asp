<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<!--#include file="c_func.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<link rel="stylesheet" href="http://cache.soso.com/wenwen/css/w_base_201012271648.css" type="text/css" media="screen"/>
<link rel="stylesheet" href="http://cache.soso.com/wenwen/css/question_201101101821.css" type="text/css" media="screen"/>
<style type="text/css">
   *{margin:0px;padding:0px;}
   #main_s_content{width:720px; margin:0 auto;word-wrap:break-word;word-break:break-all;}
   #main_s_content .answer_con{ line-height:150%;color:#111;}
   #main_s_content .oso_red{color:red;font-weight:bolder;}
   #main_s_content h3{color:#222;font-weight:bolder;font-size:18px;}
   #socaiji{width:720px;margin:0 auto;}
   #socaiji a{padding:5px;font-size:16px;color:#002346;font-weight:bolder;cursor:pointer;display:block;width:40px;border:#374EA6 solid 1px;margin-top:10px;margin-bottom:10px;text-decoration:none;}
   #socaiji span{float:right;padding:5px;font-size:16px;color:#002346;font-weight:bolder;cursor:pointer;display:block;width:80px;border:#374EA6 solid 2px;margin-right:10px;margin-top:10px;}
   #popup_show{display:none;background-color:#eee;width:600px;text-align:center; color:red;font-size:14px;border:#99CC33 solid 3px;z-index:9999;}
   #popup_show p{clear:both;padding-bottom:20px;font-size:20px;font:Verdana, Geneva, sans-serif;color:red;font-weight:bold;}
   #popup_show span{FLOAT:right;display:block;cursor:pointer;padding:2px 5px;background-color:#EFCAAB;border:solid 2px #ccc;margin:1px}
   .hover{background-color:#FBE6D7;}
   #mdltype{margin:20px 20px 0px 0px;float:right}
</style>
<script src="jquery-1.4.4.js" type="text/javascript"></script>
<script type="text/javascript">
$(window).resize(function(){
   $("#popup_show").css({
      position:'absolute',
      left: ($(window).width() - $("#popup_show").outerWidth())/2,
      top: ($(window).height() - $("#popup_show").outerHeight())/2
   });
});
//初始化函数
$(window).resize();
$(document).ready(function()
{
   $("#main_s_content span").addClass("oso_red");
   $("#socaiji span").hover(
      function(){
      $(this).addClass("hover");
   },function(){
      $(this).removeClass("hover");
   }
   );
   $("#socaiji span").click(function()
   {
      $("#main_s_content span").removeClass("oso_red");
	  //alert($("#main_s_content .answer_con").html());
	  $("#main_s_content img").each(function(i){
	     if($(this).attr("src").indexOf("http://pic.wenwen.soso.com/")>-1){
	        $(this).remove();
		 }
	  });
      $.ajax(
	  {
	  type:"POST",
	  url:"sosave.asp",
	  data:"tit="+escape($("#main_s_content h3").html())+"&cont="+escape($("#main_s_content .answer_con").html())+"&Module="+$("select[name=mdltype] option[selected]").val(),
	  async:false,
	  success:function(html)
	  {
         $("#popup_show").empty();
         $("#popup_show").hide();
	     $("#popup_show").append("<span title='点击关闭'>X</span><p>"+html+"</p>");
		 $("#popup_show").show();
		 $("body").css("filter","alpha(opacity=50)")
		 $("body").css("opacity","0.5")
         $("#main_s_content span").addClass("oso_red");
	  }
      });
      $("#popup_show span").click(function(){
         $("#popup_show").empty();
         $("#popup_show").hide();
	     $("body").css("filter","alpha(opacity=100)")
	     $("body").css("opacity","1")
      })
   });
})
</script>
<title></title>
</head>
<body oncontextmenu="return false;">
<%dim py,w,body,spi,sr,w8,qf,rn,qs,ch,act,tit,cont,body1%>
<div id="socaiji"><span title="点击采集">一键采集</span><%=selectlist%><a href="javascript:;" target="_self" onclick="javascript:history.back();">返回</a></div>
<%
py = trim(Request.querystring("py"))
w = trim(Request.querystring("w"))
spi = trim(Request.querystring("spi"))
sr = trim(Request.querystring("sr"))
w8 = trim(Request.querystring("w8"))
qf = trim(Request.querystring("qf"))
rn = trim(Request.querystring("rn"))
qs = trim(Request.querystring("qs"))
ch = trim(Request.querystring("ch"))
   url="http://wenwen.soso.com/z/"&py&"?w="&w&"&spi="&spi&"&sr="&sr&"&w8="&w8&"&qf="&qf&"&rn="&rn&"&qs="&qs&"&ch="&ch&""

   wstr=getHTTPPage(url,"utf-8")
   tit=strCut(wstr,"<div class=""question_main"">","<div class=""question_tag"">",2)
   body1=strCut(wstr,"<!--------- best answers  --------->","<!--------- right ads in question page  --------->",2)
   body=strCut(body1,"<div class=""answer_con"">","<div class=""evaluation_wrap"">",1)
   if body="" then 
      body=strCut(body1,"<div class=""sloved_answer"">","<div class=""evaluation_wrap"">",1)
	  body=strCut(body,"<div class=""pump_wrap"">","<div class=""evaluation_wrap"">",1)
	  body=replace(body,"<div class=""pump_wrap"">","<div class=""answer_con"">")
   end if
   body=replace(body,"<div class=""evaluation_wrap"">","")
   'body=replace(body,w8,"<span>"&w8&"</span>")
   r = "(<img class=ed_capture src=""http://pic.wenwen.soso.com/p/.*?\..*?"">)"
   body = py_z_replace(body,r,"")

   response.write "<div id=""main_s_content"">"&tit&body&"</div>"
act=trim(request("act"))
if act="cj" then
   tit=trim(request("txt_tit"))
   tit=replace(request("txt_tit"),"%09","")
   tit=replace(request("txt_tit")," ","")
   'response.Write ltrimVBcrlf(tit)
   'response.end
   cont=trim(request("txt_cont"))
   Module=trim(request("Module"))
   call openConn()
   set rs=server.CreateObject("adodb.recordset")
   sql="select top 1 * from [FK_Article ] where [Fk_Article_Title]='"&tit&"'"
   rs.open sql,conn,1,3
   if not rs.eof then
      response.Write "<script language=javascript>alert('此信息已采集过!');history.back();</script>"   
      rs.close
      set rs=nothing
      call closeConn()
   else
   'response.Write cont
      rs.addnew
      rs("Fk_Article_Title")=tit
      rs("Fk_Article_Content")=cont
      rs("Fk_Article_From")="互联网"
      rs("Fk_Article_Module")=Module
      rs("Fk_Article_Menu")=1
      rs("Fk_Article_Recommend")=",0,"
      rs("Fk_Article_Subject")=",0,"
      rs.update
      response.Write "<script language=javascript>alert('采集成功!');history.back(-1);</script>"
      rs.close
      set rs=nothing
      call closeConn()
   end if
end if
%><div id="main_s_content" style="display:none;">
<form action="?act=cj" method="post" name="frmcj" id="frmcj">
   <input type="hidden" name="txt_tit" id="txt_tit" value="<%'=tit%>"/><br/><br/>
   <textarea name="txt_cont" id="txt_cont"><%'=body%></textarea><br/><br/>
   <input type="submit" value="采 集"/>
</form>
</div>
<div id="popup_show"></div>
</body>
</html>
