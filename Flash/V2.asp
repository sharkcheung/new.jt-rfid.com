<!--#Include File="../Include.asp"-->
var slidedata =[
<%
'==========================================
'文 件 名：V.asp
'文件用途：Flash轮换
'版权所有：企帮网络www.qebang.cn
'==========================================

Dim Height,Width,Menu,Module
Dim Pic,Text,Url,ArticleUrl,ProductUrl,DownUrl

'获取参数
Types=Clng(Request.QueryString("Type"))
Menu=Clng(Request.QueryString("Menu"))
Module=Clng(Request.QueryString("Module"))
Height=Request.QueryString("Height")
Width=Request.QueryString("Width")
If Height="" Then
	Height=100
Else
	Height=Clng(Height)
End If
If Width="" Then
	Width=100
Else
	Width=Clng(Width)
End If

If Types=1 Then
	Sqlstr="Select Top 5 * From [Fk_ProductList] Where (Fk_Product_Module="&Module&" Or Fk_Module_LevelList Like '%%,"&Module&",%%') And Fk_Product_Pic<>'' Order By Fk_Product_Id Desc"
	Rs.Open Sqlstr,Conn,1,3
	While Not Rs.Eof

				ProductUrl=Rs("Fk_Module_Dir")&"/"&Rs("Fk_Product_FileName")&".html"

			If SiteHtml=1 Then
				ProductUrl=SiteDir&ProductUrl
			Else
				ProductUrl=SiteDir&"?"&ProductUrl
			End If
				Pic=Rs("Fk_Product_Pic")
                                                        Url=ProductUrl

Text=left(Rs("Fk_Product_Title"),3)

i=i+1
%>

{
img:'<%=Pic%>',
txt:'<%=Text%>',
url:'<%=Url%>',
target:'_blank',
btimg:'http://www.umiwi.com/public/2010/2010_09_09_09_28_45_7913.jpg'
}<%
if i<>5 then
response.write ","
end if
	Rs.MoveNext
	Wend
	Rs.Close
End If
%>
];
<!--#Include File="../Code.asp"-->
