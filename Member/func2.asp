<%
function get_count(id,intKey)
   on error resume next
   dim get_rs,get_sql
   select case intKey
      case 0
	     table="u_order"
      case 1
	     table="u_shopcart"
   end select
	  get_sql="select count(id) from "&table&" where u_id="&id&""
	  set get_rs=conn.execute(get_sql)
	  if get_rs.eof then
	     get_count=0
	  else
	     get_count="<font color=red><b>"&get_rs(0)&"</b></font>"
	  end if
	  get_rs.close
	  set get_rs=nothing
end function
'***********工作地*************
Function Hireworkadds(str)
if str<>"" and not isnull(str) then
	Select Case str
		Case "0000":Hireworkadds="不限"
		Case Else
		str=split(str,",")
			for i=0 to ubound(str)
			province_name=""
			capital_name=""
			city_name=""
			workadd=trim(str(i))
			if workadd<>"" then
			  mystring=split(workadd,"*")
			  length=ubound(mystring)
			  Set Hwasrs=Server.CreateObject("adodb.recordset")
			  select case length
			   case "0"
				if trim(mystring(0))<>"" then
				province=trim(mystring(0))
				Hwassql="select province_city from job_provinceandcity where id="&province&""
				Hwasrs.open Hwassql,conn,1,1
					if not Hwasrs.eof then
					 province_name=Hwasrs("province_city")
					end if
				Hwasrs.close
					if province_name<>"" then
					 Hireworkadds=Hireworkadds&province_name&" "
					end if
				end if
			   case "1"
				if trim(mystring(0))<>"" and trim(mystring(1))<>"" then
				province=trim(mystring(0))
				capital=trim(mystring(1))
				Hwassql="select province_city from job_provinceandcity where id="&province&""
				Hwasrs.open Hwassql,conn,1,1
					if not Hwasrs.eof then
					 province_name=Hwasrs("province_city")
					end if
				Hwasrs.close
				Hwassql="select province_city from job_provinceandcity where id="&capital&""
				Hwasrs.open Hwassql,conn,1,1
					if not Hwasrs.eof then
					 capital_name=Hwasrs("province_city")
					end if
				Hwasrs.close
					if province_name<>"" and capital_name<>"" then
					 Hireworkadds=Hireworkadds&province_name&capital_name&" "
					end if
				end if
			   case "2"
				if trim(mystring(0))<>"" and trim(mystring(1))<>"" and trim(mystring(2))<>"" then
				province=trim(mystring(0))
				capital=trim(mystring(1))
				city=trim(mystring(2))
				Hwassql="select province_city from job_provinceandcity where id="&province&""
				Hwasrs.open Hwassql,conn,1,1
					if not Hwasrs.eof then
					 province_name=Hwasrs("province_city")
					end if
				Hwasrs.close
				Hwassql="select province_city from job_provinceandcity where id="&capital&""
				Hwasrs.open Hwassql,conn,1,1
					if not Hwasrs.eof then
					 capital_name=Hwasrs("province_city")
					end if
				Hwasrs.close
				Hwassql="select province_city from job_provinceandcity where id="&city&""
				Hwasrs.open Hwassql,conn,1,1
					if not Hwasrs.eof then
					 city_name=Hwasrs("province_city")
					end if
				Hwasrs.close
					if province_name<>"" and capital_name<>"" and city_name<>"" then
					 Hireworkadds=Hireworkadds&province_name&capital_name&city_name&" "
					end if
				end if
			   end select
			  end if
			 next
	End Select
	else:Hireworkadds="不限"
end if
End Function
'分页代码
sub showpagelist(pagenum,maxperpage,linkurl,page)
if page="" or isnull(page) or page=0 then page=1
response.Write "<script type=""text/javascript"">"
response.Write "   function changepage(num){"
response.Write "      document.location.href="""&JoinChar(linkurl)&"page=""+num;"
response.Write "   }"
response.Write "</script>"
					response.Write "<div style=""clear:both;width:500px;width:98%; padding-right:20px;"">"
					if pagenum>=1 then
					   if pagenum=1 then
					      page=1
					      response.Write "首页 上一页 下一页 尾页 "
					   else
					      if page>=pagenum then
						     page=pagenum
							 response.Write "<a href='"&JoinChar(linkurl)&"page=1'>首页</a> <a href='"&JoinChar(linkurl)&"page="&page-1&"'>上一页</a> 下一页 尾页 "
					      elseif page=1 then
						     response.Write "首页 上一页 <a href='"&JoinChar(linkurl)&"page="&page+1&"'>下一页</a> <a href='"&JoinChar(linkurl)&"page="&pagenum&"'>尾页</a> "
						  else
						     response.Write "<a href='"&JoinChar(linkurl)&"page=1'>首页</a> <a href='"&JoinChar(linkurl)&"page="&page-1&"'>上一页</a> <a href='"&JoinChar(linkurl)&"page="&page+1&"'>下一页</a> <a href='"&JoinChar(linkurl)&"page="&pagenum&"'>尾页</a> "  
						  end if
					   end if
					      response.Write "每页 "&maxperpage&" 条 共 "&pagenum&" 页 跳转到第<select onchange='changepage(this.value);'>"
					      for iii=1 to pagenum
					         response.Write "<option value="&iii&" "
							 if int(request("page"))=iii then response.Write "selected"
							 response.Write "> "&iii&" </option>"
					      next
					         response.Write "</select>页"
					else
					end if
					response.Write "</div>"
end sub
'加入或
function JoinChar(join_url)
   if instr(join_url,"?")>0 then
      JoinChar=join_url&"&"
   else
      JoinChar=join_url&"?"
   end if
end function

'网上支付方式
sub pay_type()
   set payrs=conn.execute("select * from Iheeo_Pay")
   if payrs.eof then
      response.Write "暂无网上支付方式！"
   else
      do while not payrs.eof
	     response.Write "<input title=""div_OnlineBank"" style=""BORDER-RIGHT: 0px; BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 0px"" type=""radio"" value="""&payrs("PayKey")&""" name=""subRadio""/>"
		 select case payrs("PayKey")
		    case 1
			   pay_img="kq.gif"
		    case 2
			   pay_img="zfb.gif"
		    case 3
			   pay_img="tenpay.gif"
         end select
	     response.Write "<img src=""images/payType/"&pay_img&""" alt="""&trim(payrs("PayName"))&""" />"
	  payrs.movenext
	  if payrs.eof then exit do
	  loop
   end if 
   payrs.close
   set payrs=nothing
end sub

Function Encrypt(theNumber)
On Error Resume Next
Dim n, szEnc, t, HiN, LoN, i
n = CDbl((theNumber + 1570) ^ 2 - 7 * (theNumber + 1570) - 450)
If n < 0 Then szEnc = "R" Else szEnc = "J"
n = CStr(abs(n))
For i = 1 To Len(n) step 2
t = Mid(n, i, 2)
If Len(t) = 1 Then
szEnc = szEnc & t
Exit For
End If
HiN = (CInt(t) And 240) / 16
LoN = CInt(t) And 15
szEnc = szEnc & Chr(Asc("M") + HiN) & Chr(Asc("C") + LoN)
Next
Encrypt = szEnc
End Function 

Function Decrypt(theNumber)
On Error Resume Next
Dim e, n, sign, t, HiN, LoN, NewN, i
e = theNumber
If Left(e, 1) = "R" Then sign = -1 Else sign = 1
e = Mid(e, 2)
NewN = ""
For i = 1 To Len(e) step 2
t = Mid(e, i, 2)
If Asc(t) >= Asc("0") And Asc(t) <= Asc("9") Then
NewN = NewN & t
Exit For
End If
HiN = Mid(t, 1, 1)
LoN = Mid(t, 2, 1)
HiN = (Asc(HiN) - Asc("M")) * 16
LoN = Asc(LoN) - Asc("C")
t = CStr(HiN Or LoN)
If Len(t) = 1 Then t = "0" & t
NewN = NewN & t
Next
e = CDbl(NewN) * sign
Decrypt = CLng((7 + sqr(49 - 4 * (-450 - e))) / 2 - 1570)
End Function

'获取订单参数
function pay_config(pro_id,num)
   dim pay_sql,payrs
   pay_config=""
   pay_sql="select u_id,p_id,pro_num from u_order where [order_id]='"&pro_id&"'"
'   response.Write pro_id&pay_sql
'   response.end
   set payrs=conn.execute(pay_sql)
   if payrs.eof then
      response.Write "参数提交不正确!"
	  response.End
   else
      pay_config=payrs(num)
   end if
	  payrs.close
	  set payrs=nothing
end function
'商品,用户订单、购物车数据统计函数(用于分页)
function count_num(table,id)
   select case id
      case 0
	     where=""
	  case else
	     where=" where u_id="&id&""
   end select
   set countrs=conn.execute("select count(*) from "&table&""&where&"")
   count_num=countrs(0)
   countrs.close
   set countrs=nothing
end function
'获取商品参函数
function p_price(pid,filed)
   if not isnumeric(pid) then
      p_price="参数错误!"
   else
   set p_rs=conn.execute("select "&filed&" from Fk_Product where Fk_Product_Id="&pid&"")
   if p_rs.eof then
      p_price="商品已经删除或不存在!"
   else
      p_price=p_rs(0)
   end if
   p_rs.close
   set p_rs=nothing
   end if
end function
'獲取用户ID函数
function M_memberID(M_name)
   set p_rs=conn.execute("select id from u_members where m_uid='"&M_name&"'")
   if p_rs.eof then
      M_memberID="参数错误!"
   else
      M_memberID=p_rs(0)
   end if
   p_rs.close
   set p_rs=nothing
end function
'获取运费参数
function M_fee(num,SongKey)
   set p_rs=conn.execute("select * from Iheeo_Delivery where SongKey="&SongKey&"")
   if p_rs.eof then
      M_fee=0
   else
      M_fee=p_rs(num)
   end if
   p_rs.close
   set p_rs=nothing
end function

sub product_list(maxperpage,titlen,ftitlen,imgwidth,imgheight)
   page=int(request("page"))
			  linkurl="products.asp"
			  list_num=count_num("BJX_goods",0)
				    if list_num mod maxperpage=0 then
				       pagenum=list_num\maxperpage
					else
					   pagenum=list_num\maxperpage+1
					end if
			  if page="" or isnull(page) and int(page)<1 then
			     response.Redirect ""&linkurl&"?page=1"
				 response.end
			  end if
			  if page>pagenum then
			     response.Redirect ""&linkurl&"?page="&pagenum&""
				 response.end
			  end if
			  goodsql="select top "&maxperpage&" * from [BJX_goods]"
			  if page>1 then
			     goodsql=goodsql&" where bookid not in(select top "&(page-1)*maxperpage&" bookid from [BJX_goods] order by bookid desc)"
			  end if
			     goodsql=goodsql&" order by bookid desc"
			     set goodrs=conn.execute(goodsql)
				 response.Write "<ul>"
				 if goodrs.eof then
				    response.Write "暂无商品信息！"
				 else
				    do while not goodrs.eof
					   bookid = goodrs("bookid")
					   zhuang = goodrs("zhuang")
                       bookpic = goodrs("bookpic")
                       shichangjia = goodrs("shichangjia")
                       huiyuanjia = goodrs("huiyuanjia")
                       chengjiaocount = goodrs("chengjiaocount")
                       liulancount = goodrs("liulancount")
                       bookad = "【"&goodrs("bookad")&"】"
                       if len(trim(goodrs("bookname")))>titlen Or len(trim(goodrs("bookad")))>ftitlen then
                          bookname = left(trim(goodrs("bookname")),titlen)&" "&left(trim(goodrs("bookad")),ftitlen)&".."
                       else
                          bookname = goodrs("bookname")&" "&trim(goodrs("bookad"))
                       end if
					   'response.Write "<li><a href=""product.asp?bookid="&server.URLEncode(Encrypt(bookid))&""" title="""&goodrs("bookname")&" "&bookad&"""><img src='"&bookpic&"' width="&imgwidth&" height="&imgheight&"/></a><h6><a href=""product.asp?bookid="&server.URLEncode(Encrypt(bookid))&""" title="""&goodrs("bookname")&" "&bookad&""">"&bookname&"</a></h6><h7>市场价：￥ <s style=""color:#E14900;font-size:16px;font-weight:bolder;"">"&shichangjia&"</s></h7><h6>会员价：￥ <span style=""color:#E14900;font-size:16px;font-weight:bolder;"">"&huiyuanjia&"</span></h6></li>"
					goodrs.movenext
					if goodrs.eof then exit do
					loop
				 end if
				 response.Write "</ul>"
				 goodrs.close
				 set goodrs=nothing
				 call showpagelist(pagenum,maxperpage,linkurl,page)
end sub


Function getIP()
   Dim strIPAddr
   If Request.ServerVariables("HTTP_X_FORWARDED_FOR")="" OR InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"),"unknown")>0 Then
      strIPAddr=Request.ServerVariables("REMOTE_ADDR")
   ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"),",")>0 Then
      strIPAddr=Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"),1,InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"),",")-1)
   ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"),";")>0 Then
      strIPAddr=Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"),1,InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"),";")-1)
   Else
      strIPAddr=Request.ServerVariables("HTTP_X_FORWARDED_FOR")
   End If
   getIP=Trim(Mid(strIPAddr,1,30))
End Function
%>