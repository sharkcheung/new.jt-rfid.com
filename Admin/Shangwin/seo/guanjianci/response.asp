<!--#include file=menu.asp-->
<%

pp=request("pp")
if pp="" or pp="-15" then pp="0"
ppp=pp+15
pppp=pp-15


if request("htmlpage")<>"" then  '�����ϸҳ��


url="http://www.donews.com/"&aspname&"/"&request("htmlpage")
html=getHTTPPage(url)
html=strCut(html,"leftmain_center","��ࣺ",2)
title=strCut(html,"<h1>","</h1>",2)
html="wxyz"&html
aaaa="wxyz"""&">"
html=replace(html,aaaa,"")
response.write "<title>"&title&"</title></head><body>"
response.write html&"</div></div>"

else    '����б�ҳ��
url="http://www.donews.com/"&aspname&"/index.php?pp="&pp

html=getHTTPPage(url)
html=strCut(html,"<!--pindaoneirong-->","�ܹ�:",1)
html=replace(html,"http://www.donews.com/images/indx_tmp/columnlogo.jpg","images/columnlogo.jpg") '�滻����ͼƬʱ�����ʵͼƬ
html=replace(html,"index.php","")
html=replace(html,"addFeed","")
html=replace(html,"��������","�������Ӫ����Ѷ�Զ��ɼ����ݱ��")
html=replace(html,"http://www.donews.com/"&aspname&"/","?htmlpage=")

html=html&" "&shuzi+1&"ҳ 15��/ҳ����<span class='sxye'><a href='?pp="&pppp&"'>��һҳ</a><a href='?pp="&ppp&"'>��һҳ</a></span>"

response.write "<title>"&zhutitle&"</title></head><body>"

response.write html

response.write "</div><div class='fenye'>"

call pagelist(shuzi)
response.write "</div>"

response.write "</body></html>"

end if


%>