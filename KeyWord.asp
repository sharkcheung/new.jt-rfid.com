<!--#Include File="AdminCheck.asp"-->
<%
'==========================================
'文 件 名：KeyWord.asp
'文件用途：关键词库拉取页面
'版权所有：企帮网络www.qebang.cn
'==========================================

'判断权限
If Request.Cookies("FkAdminLimitId")>0 Then
	'Response.Write("无权限！")
	'Call FKDB.DB_Close()
	'Session.CodePage=936
	'Response.End()
End If

Dim KeyWord
dim listkeyword,TempItem,Thiswords,Host,ThiswordSHU1,ThiswordSHU2,nowpaiming
Dim Newstr,oArray


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
	Case 2
		Call KeyWordDo() '设置关键词库
End Select
%>

<div id="Boxs" style="display:none">
  <div id="BoxsContent">
    <div id="BoxContent"> </div>
  </div>
  <div id="AlphaBox" onClick="$('select').show();$('#Boxs').hide()"></div>
</div>
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
<form id="KeyWordSet" name="KeyWordSet" method="post" action="KeyWord.asp?Type=2" onsubmit="return false;">
  <div id="BoxTop" style="width:98%;"><span>关键词库管理</span><a style="display:none;" onclick="$('#Boxs').hide();$('select').show();"><img src="images/close3.gif"></a> </div>
  <div id="BoxContents" style="width:98%;">
    <table width="97%" border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="10" align="center"></td>
      </tr>
      <tr id="kw1" style="display:block;">
        <td align="center"><a href="javascript:void(0);" onclick="viewcilie();" class="keywdleft-b">关键词库列表</a><a href="javascript:void(0);" onclick="viewciku();" class="keywdleft-a">关键词库设置</a>
          <div class="tixinginfo"><b>优化频率</b>：数值大表示该关键词在各个页面的关键词设置中出现的频率高；数值小的关键词应进行相应的内容采集。 <br>
            <b>内链</b>：关键词内链不宜过多，应挑选5～10个核心关键词进行内链设置。 <br>
            <b>排名</b>：需正式顶级域名登录软件，类似abc.qebang.net二级域名登录查询结果无效。</div>
          <table id="keywordlisttable">
            <tr>
              <td class="kwlt1 kwlid" rowspan="2">序号</td>
              <td class="kwlt1" style="min-width:150px;" rowspan="2">关键词</td>
              <td class="kwlt5" colspan="2">关键词体现密度</td>
              <td class="kwlt1" rowspan="2">内链</td>
              <td class="kwlt1" rowspan="2"><a href="javascript:void(0);" onclick="$('.getranks').click();return false;" title="点击查询SEO排名">SEO排名</a></td>
              <td class="kwlt1" rowspan="2"><a href="javascript:void(0);" onclick="$('.getvisits').click();return false;" title="点击查询有效访问量">有效访问量</a></td>
              <td class="kwlt1" rowspan="2">内容采集</td>
            </tr>
            <tr>
              <td class="kwlt1 kwlt4">栏目页</td>
              <td class="kwlt1 kwlt4">内容页</td>
            </tr>
            <%
			Dim ubArr
			ubArr=UBoundStrToArr(listkeyword,UBound(Split(listkeyword,"|")),"|")
				For TempItem=0 To ubArr
					Thiswords=Split(listkeyword,"|")(TempItem)
					ThiswordSHU2=Chakeywordci(Thiswords,2)
					ThiswordSHU1=Chakeywordci(Thiswords,1)
					nowpaiming=Chanowpaiming(Thiswords)
					if ThiswordSHU2>100 then
					response.write "<tr><td class='kwlt2 kwlid'>"&TempItem+1&"</td><td class='kwlt2' id='kwlist"&TempItem&"' style='min-width:150px;'><b>"&Thiswords&"</b></td><td class='kwlt3 viewpbg'><div class='viewpl' style='width:"&ThiswordSHU1*3&"px;max-width:190px;'></div><span>"&ThiswordSHU1&"</span></td><td class='kwlt3 viewpbg'><div class='viewpl' style='width:"&ThiswordSHU2*1&"px;max-width:190px;'></div><span>"&ThiswordSHU2&"</span></td><td class='kwlt3'><img src='images/caidan"&ChakeywordNLink(Thiswords)&".png' /></td><td class='kwlt3' title='查询【"&Thiswords&"】的排名情况 '><a class='kwlt6 getranks' name='chapaimingasd' id='chaciarea"&TempItem+1&"' onclick=""chakeywordspaiming('"&Thiswords&"',"&TempItem+1&");""  href='javascript:void(0);'>"&nowpaiming&"</a></td><td class='kwlt3' title='查询【"&Thiswords&"】的访问量 '><a class='kwlt6 getvisits' name='chavisits' id='chavisits"&TempItem+1&"' onclick=""GetVisits("&Tjid&",'"&Thiswords&"',"&TempItem+1&","&ubArr&");return false;""  href='javascript:void(0);'>查询访问量</a></td><td class='kwlt3' title='采集包含【"&Thiswords&"】的内容 '><a onclick=""openaddcat('shangwin/seo/caiji3/?keywords="&Thiswords&"',934,475,'采集包含【"&Thiswords&"】的内容');return false;"" href='javascript:void(0);'>内容采集</a></td></tr>"
					elseif ThiswordSHU2>50 and ThiswordSHU2<100 then
					response.write "<tr><td class='kwlt2 kwlid'>"&TempItem+1&"</td><td class='kwlt2' id='kwlist"&TempItem&"' style='min-width:150px;'><b>"&Thiswords&"</b></td><td class='kwlt3 viewpbg'><div class='viewpl' style='width:"&ThiswordSHU1*3&"px;max-width:190px;'></div><span>"&ThiswordSHU1&"</span></td><td class='kwlt3 viewpbg'><div class='viewpl' style='width:"&ThiswordSHU2*2&"px;max-width:190px;'></div><span>"&ThiswordSHU2&"</span></td><td class='kwlt3'><img src='images/caidan"&ChakeywordNLink(Thiswords)&".png' /></td><td class='kwlt3' title='查询【"&Thiswords&"】的排名情况 '><a class='kwlt6 getranks' name='chapaimingasd' id='chaciarea"&TempItem+1&"' onclick=""chakeywordspaiming('"&Thiswords&"',"&TempItem+1&");""  href='javascript:void(0);'>"&nowpaiming&"</a></td><td class='kwlt3' title='查询【"&Thiswords&"】的访问量 '><a class='kwlt6 getvisits' name='chavisits' id='chavisits"&TempItem+1&"' onclick=""GetVisits("&Tjid&",'"&Thiswords&"',"&TempItem+1&","&ubArr&");return false;""  href='javascript:void(0);'>查询访问量</a></td><td class='kwlt3' title='采集包含【"&Thiswords&"】的内容 '><a onclick=""openaddcat('shangwin/seo/caiji3/?keywords="&Thiswords&"',934,475,'采集包含【"&Thiswords&"】的内容');return false;"" href='javascript:void(0);'>内容采集</a></td></tr>"
					else
					response.write "<tr><td class='kwlt2 kwlid'>"&TempItem+1&"</td><td class='kwlt2' id='kwlist"&TempItem&"' style='min-width:150px;'><b>"&Thiswords&"</b></td><td class='kwlt3 viewpbg'><div class='viewpl' style='width:"&ThiswordSHU1*3&"px;max-width:190px;'></div><span>"&ThiswordSHU1&"</span></td><td class='kwlt3 viewpbg'><div class='viewpl' style='width:"&ThiswordSHU2*3&"px;max-width:190px;'></div><span>"&ThiswordSHU2&"</span></td><td class='kwlt3'><img src='images/caidan"&ChakeywordNLink(Thiswords)&".png' /></td><td class='kwlt3' title='查询【"&Thiswords&"】的排名情况 '><a class='kwlt6 getranks' name='chapaimingasd' id='chaciarea"&TempItem+1&"' onclick=""chakeywordspaiming('"&Thiswords&"',"&TempItem+1&");""  href='javascript:void(0);'>"&nowpaiming&"</a></td><td class='kwlt3 getvisits' title='查询【"&Thiswords&"】的访问量 '><a class='kwlt6 getvisits' name='chavisits' id='chavisits"&TempItem+1&"' onclick=""GetVisits("&Tjid&",'"&Thiswords&"',"&TempItem+1&","&ubArr&");return false;""  href='javascript:void(0);'>查询访问量</a></td><td class='kwlt3' title='采集包含【"&Thiswords&"】的内容 '><a onclick=""openaddcat('shangwin/seo/caiji3/?keywords="&Thiswords&"',934,475,'采集包含【"&Thiswords&"】的内容');return false;"" href='javascript:void(0);'>内容采集</a></td></tr>"
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
        <td align="center"><a href="javascript:void(0);" onclick="viewcilie();" class="keywdleft-a">关键词库列表</a><a href="javascript:void(0);" onclick="viewciku();" class="keywdleft-b">关键词库设置</a>
          <%If Request.Cookies("FkAdminLimitId")=0 Then%>
          <a onclick="openaddcat('http://localhost:83/user/k.asp?type=6&id=<%=Tjid%>&time=2011-1-1&time2=<%=date()%>',630,475,'搜索引擎关键词来源');return false;" class="keywdleft-a" href="javascript:void(0);">搜索来源关键词</a>
          <%end if%>
          <div class="tixinginfo"><b>间隔</b>：关键词与关键词之间用半角状态下的 | 符号隔开，最后一个关键词不需要 | 符号。<br>
            <b>个数</b>：关键词在20～100个为宜。</div>
          <textarea name="KeyWord" cols="99%" style="width:99%;" rows="10" class="TextArea" id="KeyWord"><%=KeyWord%></textarea>
          <br /></td>
      </tr>
    </table>
  </div>
  <div id="BoxBottom" style="width:96%;">
    <input type="submit" onclick="chkKwNums();" class="Button" name="button" id="buttonset" style="display:none;" value="保 存" />
    <input style="display:none;" type="button" onclick="$('#Boxs').hide();$('select').show();" class="Button" name="button" id="button" value="关 闭" />
  </div>
</form>
<script language="javascript"> 
<!-- 
function chkKwNums(){
	var kwd=$("#KeyWord").val();
	kwdLen=kwd.split("|").length-1;
	if(kwdLen>100){
		alert("关键词建议不超过100个，当前关键词个数为"+ kwdLen +"个，请将关键词设置在100个以内！");
	}
	else{
		Sends('KeyWordSet','KeyWord.asp?Type=2',1,'file-shangwin.asp?filename=keyword&Viewstyle=1',0,0,'','');
	}
}

/*$(document).ready(function(){
	$(".getvisits").click()
})*/

function viewciku(){
document.getElementById("kw1").style.display="none";
document.getElementById("kw2").style.display="block";
document.getElementById("buttonset").style.display="block";
} 
 
function viewcilie(){
document.getElementById("kw1").style.display="block";
document.getElementById("kw2").style.display="none";
document.getElementById("buttonset").style.display="none";
}
//--> 
</script>
<%
End Sub


'==========================================
'函 数 名：KeyWordDo()
'作    用：设置关键词库
'参    数：
'==========================================
Sub KeyWordDo()
	KeyWord=Request("KeyWord")
	Call FKFso.CreateFile("KeyWord.dat",KeyWord)
	Response.Write("关键词库修改成功！")
End Sub


Function Easp_Escape(ByVal str)
	Dim i,c,a,s : s = ""
	If isnull(str) Then Easp_Escape = "" : Exit Function
	For i = 1 To Len(str)
		c = Mid(str,i,1)
		a = ASCW(c)
		If (a>=48 and a<=57) or (a>=65 and a<=90) or (a>=97 and a<=122) Then
			s = s & c
		ElseIf InStr("@*_+-./",c)>0 Then
			s = s & c
		ElseIf a>0 and a<16 Then
			s = s & "%0" & Hex(a)
		ElseIf a>=16 and a<256 Then
			s = s & "%" & Hex(a)
		Else
			s = s & "%u" & Hex(a)
		End If
	Next
	Easp_Escape = s
End Function

Sub Shuffle (ByRef arrInput)
    'declare local variables:
    Dim arrIndices, iSize, x
    Dim arrOriginal

    'calculate size of given array:
    iSize = UBound(arrInput)+1

    'build array of random indices:
    arrIndices = RandomNoDuplicates(0, iSize-1, iSize)

    'copy:
    arrOriginal = CopyArray(arrInput)

    'shuffle:
    For x=0 To UBound(arrIndices)
        arrInput(x) = arrOriginal(arrIndices(x))
    Next
End Sub

Function CopyArray (arr)
    Dim result(), x
    ReDim result(UBound(arr))
    For x=0 To UBound(arr)
        If IsObject(arr(x)) Then
            Set result(x) = arr(x)
        Else
            result(x) = arr(x)
        End If
    Next
    CopyArray = result
End Function

Function RandomNoDuplicates (iMin, iMax, iElements)
    'this function will return array with "iElements" elements, each of them is random
    'integer in the range "iMin"-"iMax", no duplicates.

    'make sure we won't have infinite loop:
    If (iMax-iMin+1)>iElements Then
        Exit Function
    End If

    'declare local variables:
    Dim RndArr(), x, curRand
    Dim iCount, arrValues()

    'build array of values:
    Redim arrValues(iMax-iMin)
    For x=iMin To iMax
        arrValues(x-iMin) = x
    Next

    'initialize array to return:
    Redim RndArr(iElements-1)

    'reset:
    For x=0 To UBound(RndArr)
        RndArr(x) = iMin-1
    Next

    'initialize random numbers generator engine:
    Randomize
    iCount=0

    'loop until the array is full:
    Do Until iCount>=iElements
        'create new random number:
        curRand = arrValues(CLng((Rnd*(iElements-1))+1)-1)

        'check if already has duplicate, put it in array if not
        If Not(InArray(RndArr, curRand)) Then
            RndArr(iCount)=curRand
            iCount=iCount+1
        End If

        'maybe user gave up by now...
        If Not(Response.IsClientConnected) Then
            Exit Function
        End If
    Loop

    'assign the array as return value of the function:
    RandomNoDuplicates = RndArr
End Function

Function InArray(arr, val)
    Dim x
    InArray=True
    For x=0 To UBound(arr)
        If arr(x)=val Then
            Exit Function
        End If
    Next
    InArray=False
End Function


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
%>
<!--#Include File="../Code.asp"-->
