<!--#Include File="../../../CheckToken.asp"-->
<!--#include file="../caiji2/c_func.asp"-->
<%
'Option Explicit
Session.CodePage=65001
Response.ContentType = "text/html"
Response.Charset = "utf-8"
'Response.Expires=-999
Session.Timeout=999
Dim curhost,keywords,htmlpage,pgnum,keywordsgbk,t
curhost=Request.ServerVariables("SERVER_NAME")
%>
<!--#include file="../head.asp"-->
<script type="text/javascript">
	// JavaScript Document
function goTopEx(){
        var obj=document.getElementById("goTopBtn");
        function getScrollTop(){
                return document.documentElement.scrollTop;
            }
        function setScrollTop(value){
                document.documentElement.scrollTop=value;
            }    
        window.onscroll=function(){getScrollTop()>0?obj.style.display="":obj.style.display="none";}
        obj.onclick=function(){
            var goTop=setInterval(scrollMove,10);
            function scrollMove(){
                    setScrollTop(getScrollTop()/1.1);
                    if(getScrollTop()<1)clearInterval(goTop);
                }
        }
    }

	
function URLencode(sStr){
     return escape(sStr).replace(/\+/g, '%2B').replace(/\"/g,'%22').replace(/\'/g, '%27').replace(/\//g,'%2F').replace(/\#/g,'%23');
   } 
$(document).ready(function()
{
	$("p").each(function(){
		$(this).html($(this).html().replace(" ", "").replace("/\s/g",""));
	})
	/*$(".s-mod-page").find("a").each(function(){
		$(this).attr("href","&e="+URLencode($(this).attr("href")));
	})*/
	$(window).scroll(function(){
		   $("#popup_show").css({
		      position:'absolute',
		      left: ($(window).width() - $("#popup_show").outerWidth())/2,
      			top: ($(window).height() - $('#popup_show').outerHeight())/2 + $(document).scrollTop() 
		   });
    });
	$("a").click(function(){
         $("#popup_show").empty();
         $("#popup_show").hide();
	     $("#popup_show").append("<div class='tcen'><p>页面加载中，请稍候...</p></div>");
		 $("#popup_show").show();
		   $("#popup_show").css({
		      position:'absolute',
		      left: ($(window).width() - $("#popup_show").outerWidth())/2,
      			top: ($(window).height() - $('#popup_show').outerHeight())/2 + $(document).scrollTop() 
		   });
		
	})
})
</script>
<style type="text/css">
	
	#popup_show .tcen{
	background:none ;
	background-color:#fff;
	padding-top:20px; padding-bottom:10px;
	border:1px #31699C solid; 
}
.info_list, .commity_list, .talent_list, .buys_list {
margin-bottom: 10px;
}
.info_list li,.info_list dd {
overflow: hidden;
zoom: 1;
padding: 10px 0;
border-bottom: 1px solid #e5e5e5;

}
#main_s_content .info_list h3{text-align:left;}
.info_list li h3 a,.info_list dd h3 a {
font-size: 14px;
color: #1E50A2;
}
.info_list li p ,.info_list dd{
line-height: 180%;
color: #666;
}
.info_list li h3 font, .info_list li p font,.info_list dd h3 font, .info_list dd p font {
color: #c00;
font-size: 14px;
}
.info_list li p font, .info_list dd p font{font-size:12px;}
.page_mod {
text-align: center;
font-family: "微软雅黑";
clear: both;
margin-bottom: 20px;
}
.page_mod .page_start, .page_mod .page_end, .s-mod-page .page_prev,.s-mod-page .page_next{
width: 52px;
height: 28px;
line-height: 28px;
color: #9a9a9a;
border: 1px solid #e3e3e3;
text-align: center;
background: #ebebeb;
}
.page_mod span ,.s-mod-page span{
display: inline-block;
}
.page_mod .page_prev a, .page_mod .page_next a, .s-mod-page .page_prev a,.s-mod-page .page_next a{
width: 52px;
}

.page_mod a ,.s-mod-page a{
display: inline-block;
width: 28px;
height: 28px;
line-height: 28px;
border: 1px solid #ccc;
margin: 0 2px;
text-align: center;
color: #676767;
}
.page_mod a:hover,.s-mod-page a:hover {
text-decoration: none;
background: #ed494a;
border: 1px solid #d83536;
color: #FFF;
}
a:hover {
text-decoration: underline;
}


.infolist dd {
    border-bottom: 1px solid #e5e5e5;
    overflow: hidden;
    padding: 10px 0;
}
.infolist dd p {
    color: #666;
    line-height: 180%;
}
.infolist dd p span.colred {
    color: #c00;
}
.infolist dd h3 a {
    color: #1e50a2;
    font-size: 14px;
}
.infolist dd h3 span.f-red {
    color: #c00;
}
.infolist dd h3 span.txt-d {
    color: #999;
    font-weight: normal;
}
.infolist dd h3 em {
    font-weight: normal;
    margin: 0 5px;
}
.infolist dd .pic-mid {
    border: 1px solid #d9d9d9;
    float: left;
    height: 100px;
    margin-right: 10px;
    overflow: hidden;
    width: 100px;
}
.infolist dd .pic-mid .box {
    display: table-cell;
    height: 100px;
    text-align: center;
    vertical-align: middle;
    width: 100px;
}
.infolist dd p font{color:#f00;}
.page-next{display: none;}
</style>
<div id="popup_show"></div>
<DIV style="DISPLAY: none" id=goTopBtn><IMG border=0 src="http://image001.dgcloud01.qebang.cn/caiji/to_top_blue.jpg"  title="回到顶部" alt="回到顶部"></DIV>
<SCRIPT type=text/javascript>goTopEx();</SCRIPT>
<div class="page">
	<!--#include file="../nav.asp"-->
	<div class="pagemian">
     	<div class="pagemian2">
			<!--#include file="../leftlist.asp"-->
            <div class="pageright gjcyy">
            	<div class="gjcyytop">
                	<div class="fr">
                    	<div class="danxuan">
                    	    <label><input type="radio" name="RadioGroup1" value="关键词类" id="RadioGroup1_0" />关键词类</label>
                    	    <label><input type="radio" name="RadioGroup1" value="问答类" id="RadioGroup1_1" />问答类</label>
                    	    <label><input type="radio" name="RadioGroup1" value="行业新闻类" id="RadioGroup1_2" />行业新闻类</label>
                    	</div>
                        <div class="srk"><input type="text" class="Input"/><select name=""><option value="">百度百科</option><option value="">360百科</option></select><input class="wenbenbtn" type="button" value="检索"/></div>
                        <div class="textnr">使用过的关键词:<a href="#">乐从沙发</a>&nbsp;&nbsp;<a href="#">真皮沙发</a>&nbsp;&nbsp;<a href="#">沙发</a>&nbsp;&nbsp;&nbsp;&nbsp;<a href="#" class="no2">[清空]</a></div>
                    </div>
                    <div class="fl">
                    	<h3>关键词运营说明：</h3>
                        <p>关键词运营是将设定好需要运营的关键词，软件系统会自动在互联网上检索出相对应的信息。可以将有效的信息一键采集到企业官网对应的栏目中，结合企业自身特新与信息更新要求进行二次加工编辑，然后更新到企业官网栏目。</p>
                    </div>
                </div>
                <div class="gjcyybtm">
                	<ul><%
'先判断是否保存采集
if request("act")="1" and Request.Form("T1")<>"" and Request.Form("S1")<>"" and Request.Form("D1")<>"" then 
	Set Conn = Server.CreateObject("Adodb.Connection")
	ConnStr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Server.MapPath(SiteData)
	Conn.Open ConnStr
	sql = "select top 1 * from [FK_Article ] where [Fk_Article_Title]='"&Request.Form("T1")&"'"
	Set rs = Server.CreateObject("ADODB.RecordSet")
	rs.Open sql,conn,1,3
	if rs.recordcount=0 then
	rs.addnew
	rs("Fk_Article_Title")=Request.Form("T1")
	rs("Fk_Article_Content")=Request.Form("S1")
	rs("Fk_Article_From")="互联网"
	rs("Fk_Article_Module")=Request.Form("D1")
	rs("Fk_Article_Menu")=1
	rs("Fk_Article_Recommend")=",0,"
	rs("Fk_Article_Subject")=",0,"
	rs.update
	response.write("<Script language=JavaScript>tan3('<b>采集保存成功！默认该条内容前台不显示。</b><br><br>需在【网站】对应栏目找到该条内容进行编辑。<br><br>并完善关键词后勾选【显示】即可！');</Script>")
	else
	response.write("<Script language=JavaScript>tan3('<b>数据库已存在相同内容！</b><br><br>请确认是否重复采集？！');</Script>")
	end if
	if err<>0 then
	response.write("<Script language=JavaScript>tan3('<b>采集内容保存时出错了！</b><br><br>请重试或更换其它内容采集。');</Script>")
	end if
	rs.close
	set rs=nothing
	conn.close

else

	keywords=request("keywords")
	htmlpage=request("htmlpage")
	if request("kw")="" then
	keywords=request("keywords")
	else
	keywords=request("kw")
	end if
	add="<a href='?keywords="&keywords&"'>"&keywords&"</a> "

	if keywords<>"" then
		if instr(Request.Cookies(curhost&".caijiwords"),add)>0 then
		else
			Response.Cookies(curhost&".caijiwords")=add&Request.Cookies(curhost&".caijiwords")
			Response.Cookies(curhost&".caijiwords").Expires= DateAdd("y", 100, Now()) 
		end if
	end if
	if request("clear")<>"" then
	Response.Cookies(curhost&".caijiwords")=""
	end if
	'Response.Cookies("keywords")=""
	
	%>
	<div id="c_content">
	<form method="POST" name="sbmit" action="index.asp">
	请您输入您需要运营的关键词：<input type="text" name="kw" value="<%=keywords%>" size="20"> <input type="submit" value=" 关键词运营 " name="B1" onclick="this.disabled=true;this.value='正在搜索中...';sbmit.submit();">
	<a href="../caiji2/?sp=<%=keywords%>&ch=w.search.sb">转问答类运营</a> | 
	<a href="../caiji-baike/?sp=<%=keywords%>">转行业新闻运营</a>
	</form>
	</div>
	<%
	response.write "<div id=""search_kcss"">您使用过的关键词："&Request.Cookies(curhost&".caijiwords")&"　[<a href='?clear=yes'>清空</a>]</div>"
	
	if keywords<>"" then	'列表内容
		dim pagination,ee
		pgnum=Trim(request("e"))
		ee=Trim(request("ee"))
		keywordsgbk=URLEncode(keywords,936)
		'Response.write pgnum
		Url="http://s.hc360.com/?w="&keywordsgbk&"&mc=news"
		'url="http://s.hc360.com/info/"&keywords&".html"
		if pgnum<>"" Then
		'	dim kkkk
		'	kkkk=strCut(pgnum,"w=","&",2)
		'Response.write pgnum
		'	pgnum=py_z_replace(pgnum,"w=(.*?)&","w="&URLEncode(kkkk,936)&"&")
			url=pgnum
			if ee<>"" Then
				Url=Url&"&mc=news&ee="&ee&"&v=6"
			End If
			
		End If 
		html=getHTTPPage(url,"gb2312")
		if pgnum<>"" then
			pagination=strCut(html,"<div class=""s-mod-page"">","</div>",1)
			html=strCut(html,"<dl class=""info_list infolist"">","</dl>",1)
			html=replace(html,"http://info.","detail.asp?url=http://info.")
			pagination=py_z_replace(pagination,"(<div class=""page-next[\s\S]*?</div>)*","")
			pagination=replace(pagination,"http://s.hc360.com/","?keywords="&keywords&"&e=http://s.hc360.com/")
			pagination=py_z_replace(pagination,"w=(.*?)&","w="&keywords&"&")
		Else
			if instr(html,"<dl class=""info_list infolist"">")>0 Then
				pagination=strCut(html,"<div class=""s-mod-page"">","</div>",1)
				html=strCut(html,"<dl class=""info_list infolist"">","</dl>",1)
				pagination=replace(pagination,"http://s.hc360.com/","?keywords="&keywords&"&e=http://s.hc360.com/")
				pagination=py_z_replace(pagination,"w=(.*?)&","w="&keywords&"&")
			Else
				html=strCut(html,"<!-- 搜索列表 资讯 S -->","<!-- 搜索列表 资讯 E -->",1)
				pagination=strCut(html,"<div class=""page_mod"">","</div>",1)
			End If
			html=replace(html,"http://info.","detail.asp?url=http://info.")
			pagination=py_z_replace(pagination,"(onclick[\s\S]*?\);"")*","")'替换<javascript>调用
			html=py_z_replace(html,"(<div class=""pagination"">[\s\S]*?<!--pagination end-->)*","")
		End If
		'pagination=strCut(pagination,"<div class=""page_bottom"">","</div>",1)
		html=py_z_replace(html,"(<div class=""mutual_phrase"">[\s\S]*?<!--你还可以搜索 mutual_phrase end-->)*","")
		html=py_z_replace(html,"(<script[\s\S]*?</script>)*","")'替换<javascript>调用
		html=py_z_replace(html,"(<samp[\s\S]*?</samp>)*","")
		'html=py_z_replace(html,"(<span[\s\S]*?</span>)*","")
		html=py_z_replace(html,"(<em[\s\S]*?</em>)*","")
		html=py_z_replace(html,"(<div class=""picNews""[\s\S]*?</div>)*","")
		html=replace(html,"list_news_pic","list_news")
		
		'html=replace(html,cutHtml,"")
		html=replace(html,"#C60A00","#FF0000")
		html=replace(html,keywords,"<font>"&keywords&"</font>")
		'html=replace(html,"慧聪","")
		html=replace(html,"报价","<strong>无效信息</strong>")
		html=py_z_replace(html,"(onclick[\s\S]*?\);"")*","")'替换onclick
		html=py_z_replace(html,"(替换onloadJS[\s\S]*?\);"")*","")'替换onloadJS
		html=py_z_replace(html,"(<!--pagination end-->[\s\S]*?<!--news_list end-->)*","")'替换onloadJS
		html=replace(html,"target=""_blank""","")
		html=html&pagination
		re1="http://s.hc360.com/info/"&keywordsgbk&".html"
		re2="?keywords="&keywords&"&"
		html=replace(html,re1,re2)
		html=replace(html,"http://z.hc360.com","?keywords="&keywords&"&e=http://z.hc360.com")
		if html<>"" then
			response.write "<div id='main_s_content'>"
			response.write html
			response.write "</div>"
		end if
		Set html= Nothing
		Set url= Nothing
	end if
end if%>
                    	
                    </ul>
                
                	<div class="yema">
                	<div class="fr"><a href="#">首页</a>&nbsp;&nbsp;<a href="#">上一页</a>&nbsp;&nbsp;<a class="yamano" href="#">1</a>&nbsp;&nbsp;<a href="#">下一页</a>&nbsp;&nbsp;<a href="#">尾页</a>&nbsp;&nbsp;跳转至：
                        <select name="" class="baike">
                            <option value="1">1</option>
                             <option value="1">2</option>
                        </select>
                    </div>
                	<div class="fl">每页10条&nbsp;&nbsp;共36条记录</div>
					</div>
					
				</div>
			</div>
		</div>
	 </div>

</div>

</body>
</html>


<%

'=================================函数区========================================

'统计strA：字符串,strB：查找字符个数
Function strCount(strA,strB)
 lngA = Len(strA)
 lngB = Len(strB)
 lngC = Len(Replace(strA,strB,""))
 strCount = (lngA - lngC) / lngB
End Function


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


Function getHTTPPage(Path,charset)
        t = GetBody(Path)
        getHTTPPage=BytesToBstr(t,charset)
End function

Function GetBody(url) 
        on error resume next
        'Set Retrieval = CreateObject("Microsoft.XMLHTTP") 
        Set Retrieval = CreateObject("MSXML2.XMLHTTP") 
        With Retrieval 
        .Open "Get", url, False, "", "" 
        .Send 
        if Retrieval.readystate<>4 then 
			GetBody="0"
			exit function
        end if
        GetBody = .ResponseBody
        End With 
        Set Retrieval = Nothing 
End Function

'中文乱码转换
Function BytesToBstr(body,Cset)
        dim objstream
        set objstream = Server.CreateObject("adodb.stream")
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

'URL编码转换1:汉字转URL
Function URLEncode(ByVal str, ByVal codePage)
 Dim preCodePage
 preCodePage = 65001
 Session.CodePage = codePage
 URLEncode = Server.URLEncode(str)
 Session.CodePage = preCodePage
End Function

'URL编码转换2:URL转汉字
Function URLDecode(ByVal str, ByVal charset)
 Dim strm
 Set strm = Server.CreateObject("ADODB.Stream")
 With strm
 .Type = 2
 .Charset = "iso-8859-1"
 .Open
 .WriteText Unescape(str)
 .Position = 0
 .Charset = charset
 URLDecode = .ReadText(-1)
 .Close
 End With
 Set strm = Nothing
End Function


'正则替换函数
Function regExReplace(sSource,patrn, replStr) 
Dim regEx, str1 
str1 = sSource 
Set regEx = New RegExp 
regEx.Pattern = patrn 
regEx.IgnoreCase = True 
regEx.Global = True 
regExReplace = regEx.Replace(str1, replStr) 
End Function 


function newsclasslist() '获取新闻类列表函数
newsclass="<select size='1' name='D1'>"
iii=1
Set Rs=Server.Createobject("Adodb.RecordSet")
Sql="Select * From [Fk_Module] where [Fk_Module_Type]=1 and [Fk_Module_Level]=0 "
	Rs.Open Sql,Conn,1,1
	do until rs.EOF
		Fk_Module_Name=Rs("Fk_Module_Name")
		Fk_Module_id=Rs("Fk_Module_id")
		newsclass=newsclass&"<option value='"&Fk_Module_id&"'>"&Fk_Module_Name&"↓</option>"
		
		Set Rs2=Server.Createobject("Adodb.RecordSet")
		Sql2="Select * From [Fk_Module] where [Fk_Module_Type]=1 and [Fk_Module_Level]="&Fk_Module_id
		Rs2.Open Sql2,Conn,1,1
		do until rs2.EOF
		Fk_Module_Name2=Rs2("Fk_Module_Name")
		Fk_Module_id2=Rs2("Fk_Module_id")
		newsclass=newsclass&"<option value='"&Fk_Module_id2&"'>　"&Fk_Module_Name2&"</option>"
		rs2.MoveNext
		iii=iii+1
		loop
		rs2.close
		Set Rs2=Nothing
		
	rs.MoveNext
	loop
	    
	rs.close
	Set Rs=Nothing
	
	newsclass=newsclass&"</select>"
	newsclasslist=newsclass
end function
closeConn
%>