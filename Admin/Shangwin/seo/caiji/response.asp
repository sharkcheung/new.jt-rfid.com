<!--#include file=menu.asp-->
<%

pp=request("pp")
if pp="" or pp="-15" then pp="0"
ppp=pp+15
pppp=pp-15


if request("htmlpage")<>"" then  '输出详细页面


url="http://www.donews.com/"&aspname&"/"&request("htmlpage")
html=getHTTPPage(url)
html=strCut(html,"leftmain_center","责编：",2)
title=strCut(html,"<h1>","</h1>",2)
html="wxyz"&html
aaaa="wxyz"""&">"
html=replace(html,aaaa,"")
response.write "<title>"&title&"</title></head><body>"
response.write html&"</div></div>"

else    '输出列表页面
url="http://www.donews.com/"&aspname&"/index.php?pp="&pp

html=getHTTPPage(url)
html=strCut(html,"<!--pindaoneirong-->","总共:",1)
html=replace(html,"http://www.donews.com/images/indx_tmp/columnlogo.jpg","images/columnlogo.jpg") '替换暂无图片时候的现实图片
html=replace(html,"index.php","")
html=replace(html,"addFeed","")
html=replace(html,"分享到人人","企帮网络营销资讯自动采集内容标记")
html=replace(html,"http://www.donews.com/"&aspname&"/","?htmlpage=")

html=html&" "&shuzi+1&"页 15条/页　　<span class='sxye'><a href='?pp="&pppp&"'>上一页</a><a href='?pp="&ppp&"'>下一页</a></span>"

response.write "<title>"&zhutitle&"</title></head><body>"

response.write html

response.write "</div><div class='fenye'>"

call pagelist(shuzi)
response.write "</div>"

response.write "</body></html>"

end if


%>