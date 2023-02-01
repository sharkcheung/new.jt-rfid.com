<!--#Include File="Include.asp"-->
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<meta http-equiv="X-UA-Compatible" content="IE=EmulateIE7" />
<title></title>
<link href="Css/Style.css" rel="stylesheet" type="text/css" />
<script type="text/javascript" src="../Js/jquery-1.4.2.min.js"></script>
<script type="text/javascript" src="../Js/function.js"></script>
<script type="text/javascript">
$(document).ready(function(){
	$(".getvisits").click()
})
</script>
<%
'==========================================
'文 件 名：KeyWord.asp
'文件用途：关键词库拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================


Dim KeyWord
dim listkeyword,TempItem,Thiswords,Host,ThiswordSHU1,ThiswordSHU2,nowpaiming
Dim Newstr,oArray,baseDomain
baseDomain=request.servervariables("http_host")

'获取参数
Types=Clng(Request.QueryString("Type"))

'增加数据库表
		Call FKDB.DB_Open()
		on error resume next
		rs.open "select * from keywordSV",conn,1,1
		if err.number<1 then
		Sqlstr="create table keywordSV(id COUNTER CONSTRAINT PrimaryKey PRIMARY KEY,SVkeywords text(255),SVci int,SVpaiming text(255),SVb1 text(255),SVb2 text(255),SVb3 text(255))"
		Conn.Execute(Sqlstr)
		end if
		rs.close


Select Case Types
	Case 1
		Call KeyWordBox() '读取关键词库
End Select
%>

<%
'==========================================
'函 数 名：KeyWordBox()
'作    用：读取关键词库
'参    数：
'==========================================
Sub KeyWordBox()
	KeyWord=FKFso.FsoFileRead("KeyWord.dat")
	
'------------------------------------关键词去重-----------------------------------------------------------
listkeyword=UnEscape(keyword)
	listkeyword=replace(listkeyword," ","")
	listkeyword=replace(listkeyword,"　","")
	listkeyword=replace(listkeyword,"｜","|")
	listkeyword=replace(listkeyword,"|||","|")
	listkeyword=replace(listkeyword,"||","|")
	listkeyword=replace(listkeyword,"&nbsp;","")
oArray = Split(listkeyword, "|")
Newstr = " " 											'这里的值是一个空格
For i=0 To UBound(oArray)
    If Instr(Newstr, " " & oArray(i) & " ") = 0 Then 	'在oArray(i)的前后加一个空格
        Newstr = Newstr & oArray(i) & " " 				'用空格分开
    End If
Next
Newstr=trim(Newstr)										'去掉首尾空格
Newstr=replace(Newstr," ","|")							'替换空格为|
KeyWord=Newstr
listkeyword=Newstr
'------------------------------------关键词去重-----------------------------------------------------------
	
%>
	<table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr id="kw1" style="display:block;">
            <td align="center">
<table id="keywordlisttable">
            <tr><td class="kwlt1 kwlid" rowspan="2">序号</td>
				<td class="kwlt1" style="min-width:150px;" rowspan="2">关键词</td>
				<td class="kwlt5" colspan="2">关键词体现密度</td>
				<td class="kwlt1" rowspan="2">内链</td>
				<td class="kwlt1" rowspan="2">SEO排名</td>
				<td class="kwlt1" rowspan="2">有效访问量</td></tr>
            <tr><td class="kwlt1 kwlt4">栏目页</td><td class="kwlt1 kwlt4">内容页</td></tr>
            <%
			Dim ubArr
			ubArr=UBoundStrToArr(listkeyword,UBound(Split(listkeyword,"|")),"|")
				For TempItem=0 To ubArr
					Thiswords=Split(listkeyword,"|")(TempItem)
					ThiswordSHU2=Chakeywordci(Thiswords,2)
					ThiswordSHU1=Chakeywordci(Thiswords,1)
					nowpaiming=Chanowpaiming(Thiswords)
					if ThiswordSHU2>100 then
					response.write "<tr><td class='kwlt2 kwlid'>"&TempItem+1&"</td><td class='kwlt2' id='kwlist"&TempItem&"' style='min-width:150px;'><b>"&Thiswords&"</b></td><td class='kwlt3 viewpbg'><div class='viewpl' style='width:"&ThiswordSHU1*3&"px;max-width:190px;'></div><span>"&ThiswordSHU1&"</span></td><td class='kwlt3 viewpbg'><div class='viewpl' style='width:"&ThiswordSHU2*1&"px;max-width:190px;'></div><span>"&ThiswordSHU2&"</span></td><td class='kwlt3'><img src='images/caidan"&ChakeywordNLink(Thiswords)&".png' /></td><td class='kwlt3' title='查询【"&Thiswords&"】的排名情况 '><a class='kwlt6' name='chapaimingasd' id='chaciarea"&TempItem+1&"' onclick=""GetRank('"&baseDomain&"','"&server.URLEncode(Thiswords)&"',"&TempItem+1&");""  href='javascript:void(0);'>"&nowpaiming&"</a></td><td class='kwlt3' title='查询【"&Thiswords&"】的访问量 '><a class='kwlt6 getvisits' name='chavisits' id='chavisits"&TempItem+1&"' onclick=""GetVisits("&Tjid&",'"&Thiswords&"',"&TempItem+1&","&ubArr&");return false;""  href='javascript:void(0);'>查询访问量</a></td></tr>"
					elseif ThiswordSHU2>50 and ThiswordSHU2<100 then
					response.write "<tr><td class='kwlt2 kwlid'>"&TempItem+1&"</td><td class='kwlt2' id='kwlist"&TempItem&"' style='min-width:150px;'><b>"&Thiswords&"</b></td><td class='kwlt3 viewpbg'><div class='viewpl' style='width:"&ThiswordSHU1*3&"px;max-width:190px;'></div><span>"&ThiswordSHU1&"</span></td><td class='kwlt3 viewpbg'><div class='viewpl' style='width:"&ThiswordSHU2*2&"px;max-width:190px;'></div><span>"&ThiswordSHU2&"</span></td><td class='kwlt3'><img src='images/caidan"&ChakeywordNLink(Thiswords)&".png' /></td><td class='kwlt3' title='查询【"&Thiswords&"】的排名情况 '><a class='kwlt6' name='chapaimingasd' id='chaciarea"&TempItem+1&"' onclick=""GetRank('"&baseDomain&"','"&server.URLEncode(Thiswords)&"',"&TempItem+1&");""  href='javascript:void(0);'>"&nowpaiming&"</a></td><td class='kwlt3' title='查询【"&Thiswords&"】的访问量 '><a class='kwlt6 getvisits' name='chavisits' id='chavisits"&TempItem+1&"' onclick=""GetVisits("&Tjid&",'"&Thiswords&"',"&TempItem+1&","&ubArr&");return false;""  href='javascript:void(0);'>查询访问量</a></td></tr>"
					else
					response.write "<tr><td class='kwlt2 kwlid'>"&TempItem+1&"</td><td class='kwlt2' id='kwlist"&TempItem&"' style='min-width:150px;'><b>"&Thiswords&"</b></td><td class='kwlt3 viewpbg'><div class='viewpl' style='width:"&ThiswordSHU1*3&"px;max-width:190px;'></div><span>"&ThiswordSHU1&"</span></td><td class='kwlt3 viewpbg'><div class='viewpl' style='width:"&ThiswordSHU2*3&"px;max-width:190px;'></div><span>"&ThiswordSHU2&"</span></td><td class='kwlt3'><img src='images/caidan"&ChakeywordNLink(Thiswords)&".png' /></td><td class='kwlt3' title='查询【"&Thiswords&"】的排名情况 '><a class='kwlt6' name='chapaimingasd' id='chaciarea"&TempItem+1&"' onclick=""GetRank('"&baseDomain&"','"&server.URLEncode(Thiswords)&"',"&TempItem+1&");""  href='javascript:void(0);'>"&nowpaiming&"</a></td><td class='kwlt3 getvisits' title='查询【"&Thiswords&"】的访问量 '><a class='kwlt6 getvisits' name='chavisits' id='chavisits"&TempItem+1&"' onclick=""GetVisits("&Tjid&",'"&Thiswords&"',"&TempItem+1&","&ubArr&");return false;""  href='javascript:void(0);'>查询访问量</a></td></tr>"
					end if
				Next 
				response.write "<tr><td colspan=""8"" class=""kwlt2 kwlid"">关键词数："&TempItem&" 个；总有效访问量：<span id=""tvisits"">0</span></td></tr>"
			if TempItem>100 then 
				response.write "<tr><td class='tixinginfo2' colspan='6'><b>提醒：</b>关键词数量过多，不利SEO优化，20～100个为宜。<b>建议</b>：通过<span onclick=""viewciku();"">关键词库设置</span>删除部分关键词。</td></tr>"
			end if
			if TempItem<20 then 
				response.write "<tr><td class='tixinginfo2' colspan='6'><b>提醒：</b>关键词数量过少，不利SEO优化，20～100个为宜。<b>建议</b>：通过<span onclick=""location.href='shangwin/seo/guanjianci/?words='"">关键词联想</span>或者<span onclick=""viewciku();"">关键词库设置</span>添加关键词。</td></tr>"
			end if
            %>
            </table></td>
        </tr>
         <tr id="kw2" style="display:none;">
            <td align="center"><a href="javascript:void(0);" onclick="viewcilie();" class="keywdleft-a">关键词库列表</a><a href="javascript:void(0);" onclick="viewciku();" class="keywdleft-b">关键词库设置</a><%If Request.Cookies("FkAdminLimitId")=0 Then%><a onclick="openaddcat('http://tongji2010.qebang.cn/user/k.asp?type=6&id=<%=Tjid%>&time=2011-1-1&time2=<%=date()%>',630,475,'搜索引擎关键词来源');return false;" class="keywdleft-a" href="javascript:void(0);">搜索来源关键词</a><%end if%>
            <div class="tixinginfo"><b>间隔</b>：关键词与关键词之间用半角状态下的 | 符号隔开，最后一个关键词不需要 | 符号。<br><b>个数</b>：关键词在20～100个为宜。</div><textarea name="KeyWord" cols="99%" style="width:99%;" rows="10" class="TextArea" id="KeyWord"><%=KeyWord%></textarea><br /></td>
        </tr>
    </table>
<%
End Sub


'************************* 
'函数:UBoundStrToArr 
'作用:检测原字符串转换为数组的最大下标值 
'参数:cCheckStr(需要检测的字符串) 
' cUBoundArr(生成数组的最大下标值) 
' cSpaceStr(间隔字符串) 
'返回:数组的最大下标值 
'************************ 
Public Function UBoundStrToArr(ByVal cCheckStr,ByVal cUBoundArr,ByVal cSpaceStr) 
On Error Resume Next

If Instr(cCheckStr,cSpaceStr)=0 Then 
UBoundStrToArr=cUBoundArr 
Exit Function 
End If 
Dim TempSpaceStr,UBoundValue 
TempSpaceStr=Mid(cCheckStr,Len(cCheckStr)-Len(cSpaceStr)+1) '获取字符串右侧间隔字符 
If TempSpaceStr=cSpaceStr Then '如果字符串最右侧存在间隔字符,则下标值需要-1 
UBoundValue=cUBoundArr-1 
Else 
UBoundValue=cUBoundArr 
End If 
UBoundStrToArr=UBoundValue 
End Function 


'********查询关键词在数据库某个表某个字段中出现的次数**********
Function Chakeywordci(Keywordsrt,Keywordslei)
 dim RSC1,RSC2,RSC3,SqlChastr
 RSC1=0:RSC2=0:RSC3=0
 select case Keywordslei
 case 2
	SqlChastr="Select Fk_Article_Title,Fk_Article_Keyword From Fk_Article Where Fk_Article_Title Like '%%"&Keywordsrt&"%%' or Fk_Article_Keyword Like '%%"&Keywordsrt&"%%' "
	Rs.Open SqlChastr,Conn,1,1
	RSC1=Rs.RecordCount
	Rs.Close
	
	SqlChastr="Select Fk_Product_Title,Fk_Product_Keyword From Fk_Product Where Fk_Product_Title Like '%%"&Keywordsrt&"%%' or Fk_Product_Keyword Like '%%"&Keywordsrt&"%%' "
	Rs.Open SqlChastr,Conn,1,1
	RSC2=Rs.RecordCount
	Rs.Close
case 1 	
	SqlChastr="Select Fk_Module_Keyword From Fk_Module Where Fk_Module_Keyword Like '%%"&Keywordsrt&"%%' "
	Rs.Open SqlChastr,Conn,1,1
	RSC3=Rs.RecordCount
	Rs.Close
end select
	Chakeywordci=RSC1+RSC2+RSC3
End function

'****查询关键词是否有做内链**********
Function ChakeywordNLink(Keywordsrt)
	dim SqlChastr
	SqlChastr="Select Fk_Word_Name From Fk_Word Where Fk_Word_Name Like '%%"&Keywordsrt&"%%' "
	Rs.Open SqlChastr,Conn,1,1
		if not Rs.eof then
			ChakeywordNLink=1
		else
			ChakeywordNLink=0
		end if
	Rs.Close
End function

'****查询关键词在数据库中的排名记录**********
Function Chanowpaiming(Keywordsrt)
	dim SqlChastr
	SqlChastr="Select SVkeywords,SVpaiming From [keywordSV] Where SVkeywords='"&Keywordsrt&"' "
	Rs.Open SqlChastr,Conn,1,1
		if not Rs.eof then
			Chanowpaiming=Rs("SVpaiming")
		else
			Chanowpaiming="查询排名"
		end if
	Rs.Close
End function

'****查询strA中strB出现的次数**********
Function strCount(strA,strB)
lngA = Len(strA)
lngB = Len(strB)
lngC = Len(Replace(strA, strB, ""))
strCount = (lngA - lngC) / lngB
End Function


'页面结束
KeyWord=""
SqlChastr=""
listkeyword=""
Newstr=""
set Chakeywordci=nothing
set rs=nothing
%><!--#Include File="../Code.asp"-->