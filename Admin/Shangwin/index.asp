<!-- #include file="../../inc/config.asp" -->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml" >
<head>
<script>
preloadImg=new Image();
preloadImg.src="images/BgImg.png";
</script>
<!-- #include file=inc.asp -->
<%Dim gongnenginfo,height,host,title,url,menu,linkdata,arrlink,siteid,endtimes,starttimes,u_name,u_type
	if request("id")="" then
		response.end
	end if
	id=cint(request("id"))
	gongnenginfo=request("gongnenginfo")
	linkdata= Replace(mid(gongnenginfo,InStr(gongnenginfo,"$")+1),"$","")
	arrlink=Split(linkdata,"/")
	If Len(KfUrl)=0 then
		Call FKFso.FsoLineWriteVer("../../Inc/Site.asp",52,"KfUrl=""http://"&arrlink(0)&"""")
	End if
	If Len(TjUrl)=0 then
		Call FKFso.FsoLineWriteVer("../../Inc/Site.asp",53,"TjUrl=""http://"&arrlink(1)&"""")
	End If
	height=request("height")
	if height="" then height="535"
	host=lcase(request.servervariables("HTTP_HOST")) 
	siteid=strCut(gongnenginfo,"siteid","/",2)
	endtimes=strCut(gongnenginfo,"endtime","/",2)
	starttimes=strCut(gongnenginfo,"starttime","/",2)
	u_name=LCase(request.cookies("FkAdminName"))
	If u_name="admin" Then
		u_type=1
	Else
		u_type=0
	End If 
select case id
	case 1
		title="企帮顾问服务"
		     url="http://win.qebang.net/web/service/ViewDefault.aspx?domain="&host
		     menu=1
    
	case 2
		title="站外运营"
		if instr(gongnenginfo,"tuig1")>1 then
		     url="http://popularize.shangwin.cn/?iisid="&siteid&"&u_type="&u_type
		     menu=1
		else
			url="http://win.qebang.net/shangwin/gongnengdemo/?info="&id
			menu=0
		end if
   
	case 3
		title="销管易"
		url="http://xgysys.gz004.qebang.cn/"
		menu=1
   
	case 4
		title="效果型网站"
		if instr(gongnenginfo,"web1")>1 then
			url="/admin/_Upd_Files_.asp"
			menu=1
		else
			url="http://win.qebang.net/shangwin/gongnengdemo/?info="&id
			menu=0
		end if

	case 5
		title="效果指数统计"
		if instr(gongnenginfo,"tongji_u/")<1 then
			url=tjurl&"/user/stat.asp?id="&tjid&"&starttime="&starttimes&"&endtime="&endtimes
			menu=1
		else
			url="http://win.qebang.net/shangwin/gongnengdemo/?info="&id
			menu=0
		end if
    
	case 6
		title="访客商机中心"
		if instr(gongnenginfo,"kf_u/")<1 then
			url=kfurl&"/main.html"
			menu=1
		else
			url="http://win.qebang.net/shangwin/gongnengdemo/?info="&id
			menu=0
		end if
       
	case 7
		title="诚信通"
		if instr(gongnenginfo,"ziyuan1")>1 then
			url="http://exmail.qq.com/login"
			menu=1
		else
			url="http://win.qebang.net/shangwin/gongnengdemo/?info="&id
			menu=0
		end if
		
	case 8
		title="定位策划系统"
		if instr(gongnenginfo,"seo1")>1 then
			'url="seo/?host="&host
			url="http://win.qebang.net/CorporateIndex.aspx?iisid="&siteid
			menu=1
		else
			url="http://win.qebang.net/shangwin/gongnengdemo/?info="&id
			menu=0
		end if
			
			
	case 9
		title="人才培训"
			url="http://www.qebang.net/"
			menu=1
			
	case 14
		title="微活动"
			url="http://whd.qb06.com/Home/index/"&siteid
			menu=1
	
	case 10
		title="EC营销即时通"
		if instr(gongnenginfo,"sms1")>1 then
			url="http://win.qebang.net/shangwin/sms/index.htm"
			url="http://image001.dgcloud01.qebang.cn/ec/index.html"
			menu=1
		else
			url="http://win.qebang.net/shangwin/gongnengdemo/?info="&id
			menu=0
		end if
		
	case 11
		title="站内运营"
		if instr(gongnenginfo,"seo1")>1 then
			url="seo/caiji3/"
			menu=1
		else
			url="http://win.qebang.net/shangwin/gongnengdemo/?info="&id
			menu=0
		end if
	case 12
		title="400热线电话"
		if instr(gongnenginfo,"tel4001")>1 then
			url="http://www.wo4000.com"
			menu=1
		else
			url="http://win.qebang.net/shangwin/gongnengdemo/?info="&id
			menu=0
		end if
	case 13
		title="防恶意点击"
		if instr(gongnenginfo,"ed")>1 then
			url="http://win.qebang.net/shangwin/sms/"
			menu=1
		else
			url="http://win.qebang.net/shangwin/gongnengdemo/?info="&id
			menu=0
		end if

end select
%>

<title><%=title%></title>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<link rel="stylesheet" type="text/css" href="images/style.css" />
<script>
//强制IFRAME不跳出
//var location="";
//var navigate="";
</script>

</head>
<body scroll="no" oncontextmenu="return false;" onselectstart="return false;">
<div style="display:none;position:absolute;width:99%;height:20px;border:1px solid #FFDA46;background:#FEF998;">
<%=strCut(request.servervariables("http_user_agent"),"MSIE ",";",2)
%>
<span style="position:absolute;padding-right:5px;padding-top:3px;right:0px;">关闭</span>
</div>
<div id="ListNav">
<div class="logo"><span style="cursor:default"><%=title%></span></div>
    <ul>
<% select case id
  case 1 '顾问
    call menu1()
  case 2 '推广
    call menu2()
  case 3 '邮件
    call menu3()
  case 4 '网站
    call menu4()
  case 5 '流量
    call menu5()
  case 6 '访客商机
    call menu6()
  case 7 '资源
    call menu7()
  case 8 'SEO优化
    call menu8()
  case 9 '微活动
    call menu9()
  case 10 '短信
    call menu10()
  case 11 '采集
  	call menu11()
  case 12 '400热线
    call menu12()
  case 13 '防恶意点击
  	call menu13()
  case 14 '微活动
  	call menu14()
end select %>
<%sub menu1%>
<% if menu=1 then %>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('http://win.qebang.net/web/service/ViewDefault.aspx?domain=<%=host%>');return false;" target="nonef">顾问服务</a></li>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('http://win.qebang.net/web/service/Cost_leave.aspx?domain=<%=host%>');return false;" target="nonef">留言咨询</a></li>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('http://win.qebang.net/report/IISInfo.aspx?domain=<%=host%>');return false;" target="nonef">服务记录</a></li>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('http://win.qebang.net/report/ReportListing.aspx?domain=<%=host%>');return false;" target="nonef">报表服务</a></li>
<%else%>
<li><a href="javascript:void(0);" target="_self">尚未开通</a></li>
<%end if%>
<%end sub%>
<%sub menu2%>
<% if menu=1 then 
%>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('http://popularize.shangwin.cn/new_personal.aspx?iisid=<%=siteid%>&u_type=<%=u_type%>');return false;" target="nonef">个人中心</a></li>
<!--li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('/blog/');" target="nonef">企业博客</a></li-->
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('http://popularize.shangwin.cn/new_MonthylTask.aspx?iisid=<%=siteid%>&u_type=<%=u_type%>');return false;" target="nonef">运营任务</a></li>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('http://popularize.shangwin.cn/DailyAss.aspx?iisid=<%=siteid%>&u_type=<%=u_type%>');return false;" target="nonef">日常任务</a></li>
<!--li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('http://popularize.shangwin.cn/Gifts.aspx?iisid=<%=siteid%>&u_type=<%=u_type%>');return false;" target="nonef">积分兑换</a></li-->
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('http://popularize.shangwin.cn/Accsupervise.aspx?iisid=<%=siteid%>&u_type=<%=u_type%>');return false;" target="nonef">平台账号</a></li>
<%else%>
<li><a href="javascript:void(0);" target="_self">尚未开通</a></li>
<%end if%>
<%end sub%>
<%sub menu3%>
<% if menu=1 then %>
<!--li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('http://win.qebang.net/shangwin/qmail/');return false;" target="nonef">企业邮箱</a></li-->
<%else%>
<li><a href="javascript:void(0);" target="_self">尚未开通</a></li>
<%end if%>
<%end sub%>
<%sub menu4%>
<% if menu=1 then %>
<li><a href="/" target="_blank">网站预览</a></li>
<li><a href="http://sy.qebang.cn/admin/webshow/" target="_blank">精美样式库</a></li>
<%else%>
<li><a href="javascript:void(0);" target="_self">尚未开通</a></li>
<%end if%>
<%end sub%>
<%sub menu5%>
<% if menu=1 then %>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('<%=TjUrl%>/user/stat.asp?id=<%=tjid%>');return false;" target="nonef">综合概括</a></li>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('<%=TjUrl%>/user/stat.flow.asp?type=1&id=<%=tjid%>');return false;" target="nonef">当天24小时</a></li>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('<%=TjUrl%>/user/stat.flow.asp?type=3&id=<%=tjid%>');return false;" target="nonef">本月30天</a></li>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('<%=TjUrl%>/user/stat.flow.asp?type=4&id=<%=tjid%>');return false;" target="nonef">全年12月</a></li>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('<%=TjUrl%>/user/stat.runtime.asp?type=1&id=<%=tjid%>');return false;" target="nonef">Top100</a></li>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('<%=TjUrl%>/user/stat.visit.asp?type=4&id=<%=tjid%>');return false;" target="nonef">搜索来源</a></li>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('<%=TjUrl%>/user/stat.visit.asp?type=6&id=<%=tjid%>');return false;" target="nonef">来源关键词</a></li>

<%else%>
<li><a href="javascript:void(0);" target="_self">尚未开通</a></li>
<%end if%>
<%end sub%>
<%sub menu6%>
<% if menu=1 then %>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('<%=kfurl%>/setting/guestbook.asp?gongnenginfo=<%=gongnenginfo%>');return false;" target="nonef">访客留言</a></li>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('<%=kfurl%>/setting/set_kf.asp?gongnenginfo=<%=gongnenginfo%>');return false;" target="nonef">客服管理</a></li>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('<%=kfurl%>/setting/modi_infoa.asp?gongnenginfo=<%=gongnenginfo%>');return false;" target="nonef">信息设置</a></li>
<li style="display:none"><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('<%=kfurl%>/setting/info.asp?s=qq');" target="nonef">在线QQ</a></li>
<!--li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('<%=kfurl%>/setting/module_test.asp');return false;" target="nonef">功能开通</a></li-->
<%else%>
<li><a href="javascript:void(0);" target="_self">尚未开通</a></li>
<%end if%>
<%end sub%>
<%sub menu7%>
<% if menu=1 then %>
<%else%>
<li><a href="javascript:void(0);" target="_self">尚未开通</a></li>
<%end if%>
<%end sub%>
<%sub menu8%>
<% if menu=1 then %>
<li style="display:none"><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('seo/?host=<%=host%>');return false;" target="nonef">返 回</a></li>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('http://win.qebang.net/CorporateIndex.aspx?iisid=<%=strCut(gongnenginfo,"siteid","/",2)%>');return false;" target="nonef">定位系统</a></li>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('../file-shangwin.asp?viewstyle=1&filename=keyword&iisid=<%=strCut(gongnenginfo,"siteid","/",2)%>');return false;" target="nonef">关键词库</a></li>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('../file-shangwin.asp?filename=Word&Viewstyle=0');return false;" target="nonef">关键词内链</a></li>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('../file-shangwin.asp?filename=Moduleseo&Viewstyle=7');return false;" target="nonef">栏目SEO</a></li>
<!--li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('seo/paiming/?url=<%=host%>');return false;" target="nonef">SEO排名</a></li>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('seo/shoulu/?url=<%=host%>');return false;" target="nonef">SEO收录</a></li-->
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('../file-shangwin.asp?filename=Map&Viewstyle=1');return false;" target="nonef">SEO索引</a></li>
<%else%>
<li><a href="javascript:void(0);" target="_self">尚未开通</a></li>
<%end if%>
<%end sub%>
<%sub menu9%>
<% if menu=1 then %>
<%else%>
<li><a href="javascript:void(0);" target="_self">尚未开通</a></li>
<%end if%>
<%end sub%>
<%sub menu10%>
<% if menu=1 then %>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('index.htm');return false;" target="nonef">腾讯EC</a></li>
<%else%>
<li><a href="javascript:void(0);" target="_self">尚未开通</a></li>
<%end if%>
<%end sub%>
<%sub menu11	'采集%>
<% if menu=1 then %>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('seo/caiji3/');return false;" target="nonef">关键司运营</a></li>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('seo/caiji2/');return false;" target="nonef">问答类运营</a></li>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('seo/caiji-baike/');return false;" target="nonef">行业新闻运营</a></li>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('../file-shangwin.asp?filename=log&Viewstyle=0');return false;" target="nonef">工作记录</a></li>
<%else%>
<li><a href="javascript:void(0);" target="_self">尚未开通</a></li>
<%end if%>
<%end sub%>
<%sub menu12	'400%>
<% if menu=1 then %>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('http://www.wo4000.com');return false;" target="nonef">4000开头</a></li>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('http://www.c4006.com');return false;" target="nonef">4006开头</a></li>
<%else%>
<li><a href="javascript:void(0);" target="_self">尚未开通</a></li>
<%end if%>
<%end sub%>
<%sub menu13	'防恶意点击%>
<% if menu=1 then %>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('/');return false;" target="nonef">防恶意点击</a></li>
<%else%>
<li><a href="javascript:void(0);" target="_self">尚未开通</a></li>
<%end if%>
<%end sub%>
<%sub menu14%>
<% if menu=1 then %>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('http://whd.qb06.com/Home/index/<%=siteid%>');return false;" target="nonef">发布活动</a></li>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('http://whd.qb06.com/Home/getmypopup/<%=siteid%>');return false;" target="nonef">活动管理</a></li>
<!--li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('../file-shangwin.asp?filename=weixin/Weixin_Sucai&Viewstyle=2');return false;" target="nonef">素材库管理</a></li>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('../file-shangwin.asp?filename=weixin/Weixin_ImgText&Viewstyle=2');return false;" target="nonef">图文管理</a></li>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('../file-shangwin.asp?filename=weixin/Weixin_CustReply&Viewstyle=2');return false;" target="nonef">自定义回复</a></li>
<li><a href="javascript:void(0);" onmousedown="loadingview();return false;" onclick="openframeurl('../file-shangwin.asp?filename=weixin/Weixin_menu&Viewstyle=2');return false;" target="nonef">自定义菜单</a></li-->
<%else%>
<li><a href="javascript:void(0);" target="_self">尚未开通</a></li>
<%end if%>
<%end sub%>


    </ul>
</div>
<div id="ListTop">
<iframe id="mainiframe" style="overflow-x: hidden; overflow-y: auto;" marginwidth="0" marginheight="0" src="<%=url%>" frameborder="0" width="100%" height="<%=height%>" name="mainiframe">
</iframe>
<iframe id="nonef" name="nonef" frameborder="0" width="0" height="0" style="display:none;"></iframe>
</div>
<div id="loading" style="display:;position:absolute; top:9px; right:10px; z-index:10000; ">
	 <img alt="" src="/admin/Images/loading.gif"></div>
<script language="javascript"> 
<!-- 
 
var frame = document.getElementById("mainiframe"); 
frame.onreadystatechange = function(){ 
if( this.readyState == "complete" ) 
document.getElementById("loading").style.display="none";
} 
 
function loadingview(){
document.getElementById("loading").style.display="block";
}
//--> 
</script>
<script language="javascript"> 
<!-- 
function openframeurl(frameurl){
 document.all.mainiframe.src= frameurl;
}
//--> 
</script>
</body>
<%
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
</html>