<!--#Include File="AdminCheck.asp"-->
<!--#Include File="../member/func2.asp"-->
<%dim action,songid,songsubject,songidorder,SongFei,songkey
Types=Clng(Request.QueryString("Type"))

Select Case Types
	Case 1
		Call songhuoList() '送货方式列表
	Case 2
		Call songhuoEditDo() '执行送货方式修改
	Case 3
		Call songhuoAddDo() '执行送货方式修改
	Case 4
		Call songhuoDelDo() '执行送货方式修改
	Case Else
		Response.Write("没有找到此功能项！")
End Select

sub songhuoEditDo()
songid=request("id")
songsubject=trim(request("subject"))
songidorder=request("songidorder")
SongFei=request("SongFei")
songkey=request("key")
if songsubject="" then
	response.Write "请输入送货方式！"
	response.end
end if
if SongFei="" then
	response.Write "请输入费用！"
	response.end
end if
set rs=server.CreateObject("adodb.recordset")

rs.open "select * from Iheeo_Delivery where songid="&songid,conn,1,3
rs("SongName")=songsubject
rs("SongList")=songidorder
rs("SongFei")=SongFei
rs("SongKey")=songkey
rs.update
rs.close
set rs=nothing
response.write "成功修改送货方式！"
end sub
'/////添加送货方式
sub songhuoAddDo()
songsubject=trim(request("subject"))
songidorder=request("songidorder")
SongFei=request("SongFei")
songkey=request("key")
if songsubject="" then
	response.Write "请输入送货方式！"
	response.end
end if
if SongFei="" then
	response.Write "请输入费用！"
	response.end
end if

rs.open "select * from Iheeo_Delivery",conn,1,3
rs.addnew
rs("SongName")=songsubject
rs("SongList")=songidorder
rs("SongFei")=SongFei
rs("SongKey")=songkey
rs.update
rs.close
set rs=nothing
response.write "成功添加送货方式！"
end sub

'/////删除送货方式
sub songhuoDelDo()
songid=request("id")
conn.execute "delete from Iheeo_Delivery where songid="&songid
response.write "送货方式删除成功！"
end sub

sub songhuoList()
Session("NowPage")=FkFun.GetNowUrl()%>

<div id="ListNav">
  <ul>
    <li><a href="javascript:void(0);" onclick="SetRContent('MainRight','<%=Session("NowPage")%>');return false">刷新</a></li>
  </ul>
</div>
<div id="ListTop"> 送货方式设置 </div>
<div id="ListContent">
  <table width="100%" border="0" align="center" cellpadding="2" cellspacing="0">
          <tr >
            <td width="30%" align="center" >送货方式</td>
            <td width="10%" align="center" >排 序</td>
            <td width="15%" align="center" > 费 用</td>
            <td width="27%" align="center" > KEY</td>
            <td width="18%" align="center" >操 作</td>
          </tr>
          <%'dim i,j
		set rs=server.CreateObject("adodb.recordset")
		rs.open "select * from Iheeo_Delivery order by SongID",conn,1,1
		i=rs.recordcount
		  dim j
		  j=0
		do while not rs.eof
		j=j+1%>
          <tr>
            <form name="form1"  method="post">
              <td  align="center"><input class="Input" name="subject" type="text" id="subject<%=j%>" size="20" value=<%=trim(rs("SongName"))%>>
              </td>
              <td  align="center"><input class="Input" name="songidorder" type="text" id="songidorder<%=j%>" size="6" value=<%=rs("SongList")%> onKeyPress	= "return regInput(this,	/^[0-9]*$/,		String.fromCharCode(event.keyCode))"
		onpaste		= "return regInput(this,	/^[0-9]*$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^[0-9]*$/,		event.dataTransfer.getData('Text'))">
              </td>
              <td  align="center"><input class="Input" name="SongFei" type="text" id="SongFei<%=j%>" size="6" value=<%=rs("SongFei")%>>
                元</td>
              <td  align="center"><input class="Input" name="key" type="text" id="key<%=j%>" size="6" value=<%=rs("SongKey")%>></td>
              <td  STYLE='PADDING-LEFT: 20px'><input type="button" onclick="DelIt('','songhuo.asp?Type=2&id=<%=rs("SongID")%>&subject='+escape($('#subject<%=j%>').val())+'&songidorder='+$('#songidorder<%=j%>').val()+'&SongFei='+$('#SongFei<%=j%>').val()+'&key='+$('#key<%=j%>').val(),'MainRight','<%=Session("NowPage")%>');" name="Submit" class="Button" value="修 改">
                &nbsp;<a onclick="DelIt('您确认要删除，此操作不可逆！','songhuo.asp?Type=4&id=<%=rs("SongID")%>','MainRight','<%=Session("NowPage")%>');return false;" href="javascript:void(0);"><font color="#FF0000">删除</font></a></td>
            </form>
          </tr>
          <%rs.movenext
			loop
			rs.close
			set rs=nothing%>
          <tr>
            <td colspan="5"  align="left" > &nbsp; &nbsp; &nbsp; 添加送货方式</td>
          </tr>
          <tr>
            <form name="form2" method="post">
              <td  align="center"><input class="Input" name="subject_add" type="text" id="subject_add" size="20">
              </td>
              <td  align="center"><input class="Input" name="songidorder_add" type="text" id="songidorder_add" value=<%=i+1%> size="6" onKeyPress= "return regInput(this,	/^[0-9]*$/,		String.fromCharCode(event.keyCode))"
		onpaste		= "return regInput(this,	/^[0-9]*$/,		window.clipboardData.getData('Text'))"
		ondrop		= "return regInput(this,	/^[0-9]*$/,		event.dataTransfer.getData('Text'))">
              </td>
              <td  align="center"><input class="Input" name="SongFei_add" type="text" id="SongFei_add" size="6">
                元</td>
              <td  align="center"><input class="Input" name="key_add" type="text" id="key_add" size="6"></td>
              <td  STYLE='PADDING-LEFT: 20px'><input type="button" name="Submit3" onclick="DelIt('','songhuo.asp?Type=3&subject='+escape($('#subject_add').val())+'&songidorder='+$('#songidorder_add').val()+'&SongFei='+$('#SongFei_add').val()+'&key='+$('#key_add').val(),'MainRight','<%=Session("NowPage")%>');return false;" href="javascript:void(0);" class="Button" value="添 加">
              </td>
            </form>
          </tr>
        </table>
</div>
<div id="ListBottom">
</div>
<script>
	function regInput(obj, reg, inputStr)
	{
		var docSel	= document.selection.createRange()
		if (docSel.parentElement().tagName != "INPUT")	return false
		oSel = docSel.duplicate()
		oSel.text = ""
		var srcRange	= obj.createTextRange()
		oSel.setEndPoint("StartToStart", srcRange)
		var str = oSel.text + inputStr + srcRange.text.substr(oSel.text.length)
		return reg.test(str)
	}
</script>
<%end sub%>
