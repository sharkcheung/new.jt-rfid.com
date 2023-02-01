<div class="pageleft">
	<ul>
		<li><a href="http://admin.qbt.qebang.com/index.php/home/operationManage/initial?<%=tokenpara%>">关键词分析</a></li>
		<li><a href="http://admin.qbt.qebang.com/index.php/home/operationManage/keywords_library?<%=tokenpara%>">关键词库</a></li>
		<li><a <%if instr(pathfilename,"moduleseo")>0 then%>style="font-weight:bold;"<%end if%> href="http://<%=cur_domain%>/admin-new/moduleseo.asp?usertype=<%=strUsertype%>&mobile=<%=strMobile%>&token=<%=strToken%>&menuid=1&type=7">关键词布局</a></li>
		<li><a <%if instr(pathfilename,"index-word")>0 then%>style="font-weight:bold;"<%end if%> href="http://<%=cur_domain%>/admin-new/shangwin-login-2.asp?op=sync_word&userType=<%=strUsertype%>&mobile=<%=strMobile%>&token=<%=strToken%>&kfurl=<%=server.urlencode(strkfurl)%>&tjurl=<%=server.urlencode(strtjurl)%>">关键词内链</a></li>
		<li><a <%if instr(pathfilename,"shangwin/seo")>0 then%>style="font-weight:bold;"<%end if%> href="http://admin.qbt.qebang.com/index.php/home/operationManage/collect?<%=tokenpara%>">关键词运营</a></li>
		<li><a href="http://<%=cur_domain%>/admin-new/Map.asp?usertype=<%=strUsertype%>&mobile=<%=strMobile%>&token=<%=strToken%>">SEO索引地图</a></li>
	</ul>
</div>