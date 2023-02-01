<!--#Include File="../../inc/conn.asp"-->
<!--#Include File="../../class/Cls_DB.asp"-->
<!--#include file="sha1.asp"-->
<!--#include file="wechat.class.asp"-->
<!--#include file="cache.class.asp"-->
<%
dim signature	'微信加密签名
dim timestamp	'时间戳
dim nonce		'随机数
dim Token
dim Subscribe
dim wx_Random
dim signaturetmp

dim ToUserName	'开发者微信号
dim FromUserName'发送方帐号（一个OpenID）
dim CreateTime	'消息创建时间（整型）
dim MsgType		'text
dim strEvent	'事件类型
dim EventKey	'事件KEY值，与自定义菜单接口中KEY值对应
dim Content		'文本消息内容
dim MsgId		'消息id，64位整型
dim arrXml,returnstr,xmlhtml
dim cur_host,url,picurl
dim FKDB
Dim Conn,Rs,Sqlstr,SiteData,SiteDir,SiteDBDir
dim newsid,imgText_Id_List,os,ncount,arrN
dim receiveType,replyType,wx_NoneReply,wx_AppId,wx_AppSecret
Set FKDB=New Cls_DB
Call FKDB.DB_Open()
if not CheckFields("wx_NoneReply","weixin_config") then
	conn.execute("alter table weixin_config add column wx_NoneReply varchar(200) null")
end if
set rs=conn.execute("select top 1 wx_token,wx_Subscribe,wx_Random,wx_NoneReply,wx_AppId,wx_AppSecret from weixin_config")
if not rs.eof then
	Token=trim(rs(0)&" ")
	wx_Subscribe=trim(rs(1)&" ")
	wx_NoneReply=trim(rs(3)&" ")
	wx_Random=rs(2)
	wx_AppId=rs(4)
	wx_AppSecret=rs(5)
end if
rs.close
if Token="" then
	response.write "未设置Token"
	response.end
end if
on error resume next
dim wxObj
set wxObj = new WeixinClass
wxObj.SetToken=Token
if request("echostr")<>"" then
	wxObj.valid()
end if
cur_host=Request.ServerVariables("Server_Name")
arrXml=wxObj.getXml_toString()
'wxObj.WriteFile(arrXml)
'arrXml="oCulat7l7wx_XpkJRyKH09rmmDMo{$}gh_81e0484bc4a3{$}event{$}CLICK{$}1393040069{$}听歌"
if instr(arrXml,"{$}")>0 then
arrXml=split(arrXml,"{$}")
FromUserName=arrXml(0)
ToUserName=arrXml(1)
MsgType=arrXml(2)
select case MsgType
	case "event"
		strEvent=arrXml(3)
		'wxObj.WriteFile(strEvent)
		if strEvent="subscribe" then
			'wxObj.WriteFile(wx_Subscribe)
			receiveType=0
			call updateSubscibe(FromUserName,0)
			if instr(wx_Subscribe,"[wx_news=")>0 then
				'wxObj.WriteFile(wx_Subscribe)
				newsid=wxObj.strCut(wx_Subscribe,"[wx_news=","]",2)
				if isnumeric(newsid) then
					'wxObj.WriteFile(newsid)
					set rs=conn.execute("select id,imgText_Title,imgText_Pic,imgText_url,imgText_Summary,imgText_Id_List from weixin_imageText where imgText_status=0 and id="&newsid&"")
					if not rs.eof then
						imgText_Id_List=trim(rs("imgText_Id_List")&" ")
						if imgText_Id_List<>"" then
							if instr(imgText_Id_List,",")>0 then
								arrN=split(imgText_Id_List,",")
								ncount=ubound(arrN)+1
							else
								ncount=1
							end if
									if rs("imgText_url")<>"" then
										url=rs("imgText_url")
									else
										url="http://"&cur_host&"/admin/weixin/weixin_show.asp?id="&rs("id")
									end if
									if instr(rs("imgText_Pic"),"http://")>0 then
										picurl=rs("imgText_Pic")
									else
										picurl="http://"&cur_host&rs("imgText_Pic")
									end if
							xmlhtml="<ArticleCount>"&ncount+1&"</ArticleCount>"
							xmlhtml=xmlhtml&"<Articles>"
							xmlhtml=xmlhtml&"<item>" &_
								"<Title><![CDATA["&rs("imgText_Title")&"]]></Title>" &_
								"<Description><![CDATA["&rs("imgText_Summary")&"]]></Description>" &_
								"<PicUrl><![CDATA["&picurl&"]]></PicUrl>" &_
								"<Url><![CDATA["&url&"]]></Url>" &_
								"</item>"
							set os=conn.execute("select id,imgText_Title,imgText_Pic,imgText_url,imgText_Summary from weixin_imageText where imgText_status=0 and id in("&imgText_Id_List&") order by imgText_px desc")
							if not os.eof then
								do while not os.eof
									if os("imgText_url")<>"" then
										url=os("imgText_url")
									else
										url="http://"&cur_host&"/admin/weixin/weixin_show.asp?id="&os("id")
									end if
									if instr(os("imgText_Pic"),"http://")>0 then
										picurl=os("imgText_Pic")
									else
										picurl="http://"&cur_host&os("imgText_Pic")
									end if
									xmlhtml=xmlhtml&"<item>" &_
									"<Title><![CDATA["&os("imgText_Title")&"]]></Title>" &_
									"<Description><![CDATA["&os("imgText_Summary")&"]]></Description>" &_
									"<PicUrl><![CDATA["&picurl&"]]></PicUrl>" &_
									"<Url><![CDATA["&url&"]]></Url>" &_
									"</item>"
								os.movenext
								loop
								xmlhtml=xmlhtml&"</Articles>"
							end if
							os.close
						else
									if rs("imgText_url")<>"" then
										url=rs("imgText_url")
									else
										url="http://"&cur_host&"/admin/weixin/weixin_show.asp?id="&rs("id")
									end if			
									if instr(rs("imgText_Pic"),"http://")>0 then
										picurl=rs("imgText_Pic")
									else
										picurl="http://"&cur_host&rs("imgText_Pic")
									end if			
							xmlhtml="<ArticleCount>1</ArticleCount>"
							xmlhtml=xmlhtml&"<Articles>"
							xmlhtml=xmlhtml&"<item>" &_
											"<Title><![CDATA["&rs("imgText_Title")&"]]></Title>" &_
											"<Description><![CDATA["&rs("imgText_Summary")&"]]></Description>" &_
											"<PicUrl><![CDATA["&picurl&"]]></PicUrl>" &_
											"<Url><![CDATA["&url&"]]></Url>" &_
											"</item>"
							xmlhtml=xmlhtml&"</Articles>"
						end if
						returnstr=wxObj.transmitImage(FromUserName,ToUserName,xmlhtml)
					else
						xmlhtml="感谢关注！"
						returnstr=wxObj.transmitText(FromUserName,ToUserName,xmlhtml)
					end if
				else
					xmlhtml="感谢关注！"	
					returnstr=wxObj.transmitText(FromUserName,ToUserName,xmlhtml)			
				end if
			else
				xmlhtml=wx_Subscribe
				returnstr=wxObj.transmitText(FromUserName,ToUserName,xmlhtml)
			'setCache FromUserName&"_do","subscribe",600
			end if
			response.write returnstr
		elseif strEvent="unsubscribe" then
			'unsubscribe(FromUserName)
			receiveType=1
			call updateSubscibe(FromUserName,1)
			xmlhtml="呜呜...您怎么能抛弃企帮小Q呢"
			response.write returnstr
		elseif strEvent="CLICK" then
			receiveType=2
			EventKey=arrXml(5)
			call TextReply(EventKey)
		else
			receiveType=3
			xmlhtml="新事件:"&strEvent
			returnstr=wxObj.transmitText(FromUserName,ToUserName,xmlhtml)
			response.write returnstr
		end if
	case "text"
		receiveType=4
		Content=trim(arrXml(3))
		call TextReply(Content)
	case "image"
		receiveType=5
		Content="亲，不好意思，我还不能识别图片，只能等客服MM回复你了^_^"
		returnstr=wxObj.transmitText(FromUserName,ToUserName,Content)
		response.write returnstr
	case "voice"
		receiveType=6
		Content="亲，不好意思，我还不能识别语音，只能等客服MM回复你了^_^"
		returnstr=wxObj.transmitText(FromUserName,ToUserName,Content)
		response.write returnstr
	case "video"
		receiveType=7
		Content="亲，不好意思，我还不能识别视频，只能等客服MM回复你了^_^"
		returnstr=wxObj.transmitText(FromUserName,ToUserName,Content)
		response.write returnstr
	case "location"
		receiveType=8
		Content="亲，不好意思，我还不能识别地理，只能等客服MM处理了^_^"
		returnstr=wxObj.transmitText(FromUserName,ToUserName,Content)
		response.write returnstr
	case "link"
		receiveType=9
		Content="亲，您发的是一个链接哦，我可不会自动点击哦，只能等客服MM点了^_^"
		returnstr=wxObj.transmitText(FromUserName,ToUserName,Content)
		response.write returnstr
	case else
		receiveType=10
		Content="您发的是神马哦，我不认识哦，只能等客服MM处理了^_^"
		returnstr=wxObj.transmitText(FromUserName,ToUserName,Content)
		response.write returnstr
end select
if err then 
	WriteFile(err.description)
	err.clear
	response.end
end if
end if
set wxObj=nothing
sub updateSubscibe(openid,t)
	if t=0 then
	
		if not CheckFields("wxnickname","weixin_subscribeList") then
			conn.execute("alter table weixin_subscribeList add column wxnickname varchar(100) null")
			conn.execute("alter table weixin_subscribeList add column wxsex int default 0")
			conn.execute("alter table weixin_subscribeList add column wxlanguage varchar(100) null")
			conn.execute("alter table weixin_subscribeList add column wxcity varchar(50) null")
			conn.execute("alter table weixin_subscribeList add column wxprovince varchar(50) null")
			conn.execute("alter table weixin_subscribeList add column wxcountry varchar(50) null")
			conn.execute("alter table weixin_subscribeList add column wxheadimgurl varchar(200) null")
			conn.execute("alter table weixin_subscribeList add column wxremark varchar(255) null")
		end if
		dim nickname,sex,language,city,province,country,strheadimgurl,remark
		if wx_AppId<>"" then
			dim access_token,obj
			url="https://api.weixin.qq.com/cgi-bin/token?grant_type=client_credential&appid="&wx_AppId&"&secret="&wx_AppSecret
			returnstr=GetURL(url)
			Set obj = parseJSON(returnstr)
			access_token=obj.access_token
			'wxObj.WriteFile("access_token===>"&access_token)
			set obj=nothing
			url="https://api.weixin.qq.com/cgi-bin/user/info?access_token="&access_token&"&openid="&openid&"&lang=zh_CN"
			returnstr=GetURL(url)
			'wxObj.WriteFile("returnstr===>"&returnstr)
			Set obj = parseJSON(returnstr)
			nickname=obj.nickname
			subscribe_time=obj.subscribe_time
			sex=obj.sex
			language=obj.language
			city=obj.city
			province=obj.province
			country=obj.country
			strheadimgurl=obj.headimgurl
			remark=obj.remark
			set obj=nothing
			'wxObj.WriteFile("nickname===>"&nickname)
			conn.execute("insert into weixin_subscribeList(openID,wxnickname,wxsex,wxlanguage,wxcity,wxprovince,wxcountry,wxheadimgurl,wxremark) values('"&openid&"','"&nickname&"',"&sex&",'"&language&"','"&city&"','"&province&"','"&country&"','"&strheadimgurl&"','"&remark&"')")
		else
			conn.execute("insert into weixin_subscribeList(openID) values('"&openid&"')")
		end if
	else
		conn.execute("delete * from weixin_subscribeList where openID='"&openid&"'")
	end if
end sub 

sub TextReply(strContent)

		'response.write(strContent)
		dim srs,sqlxml,replyCounts,returnXml
		if wxObj.isQqFace(strContent) then
			returnXml=wxObj.transmitText(FromUserName,ToUserName,strContent)
			'wxObj.WriteFile(returnXml)
			response.write returnXml
		else
			sqlxml="select top 1 reply_qanswerText,reply_qanswerNews,reply_qanswerResource from Weixin_CustReply where status=0 and InStr(1,LCase(reply_qtitle),LCase('"&strContent&"'),0)<>0 order by px desc"
			set rs=conn.execute(sqlxml)
			'wxObj.WriteFile(sqlxml)
			if not rs.eof then
				if rs("reply_qanswerNews")<>"" then				
					if wx_Random=0 then
						sqlxml="select top 1 id,imgText_Title,imgText_Pic,imgText_url,imgText_Summary,imgText_Id_List from weixin_imageText where id in("&rs("reply_qanswerNews")&") and imgText_status=0 order by imgText_px desc"
					else
						Randomize()
						sqlxml="select top 1 id,imgText_Title,imgText_Pic,imgText_url,imgText_Summary,imgText_Id_List from weixin_imageText where id in("&rs("reply_qanswerNews")&") and imgText_status=0 order by rnd(-(id + " & Int((10000 * Rnd) + 1) & "))"
					end if
					set os=conn.execute(sqlxml)
					if not os.eof then
									if os("imgText_url")<>"" then
										url=os("imgText_url")
									else
										url="http://"&cur_host&"/admin/weixin/weixin_show.asp?id="&os("id")
									end if	
									if instr(os("imgText_Pic"),"http://")>0 then
										picurl=os("imgText_Pic")
									else
										picurl="http://"&cur_host&os("imgText_Pic")
									end if
						imgText_Id_List=trim(os("imgText_Id_List")&" ")
						if imgText_Id_List<>"" then
							if instr(imgText_Id_List,",")>0 then
								arrN=split(imgText_Id_List,",")
								ncount=ubound(arrN)+1
							else
								ncount=1
							end if	
							xmlhtml="<ArticleCount>"&ncount+1&"</ArticleCount>"
							xmlhtml=xmlhtml&"<Articles>"
							xmlhtml=xmlhtml&"<item>" &_
								"<Title><![CDATA["&os("imgText_Title")&"]]></Title>" &_
								"<Description><![CDATA["&os("imgText_Summary")&"]]></Description>" &_
								"<PicUrl><![CDATA["&picurl&"]]></PicUrl>" &_
								"<Url><![CDATA["&url&"]]></Url>" &_
								"</item>"
							set srs=conn.execute("select id,imgText_Title,imgText_Pic,imgText_url,imgText_Summary from weixin_imageText where imgText_status=0 and id in("&imgText_Id_List&") order by imgText_px desc")
							if not srs.eof then
								do while not srs.eof
									if srs("imgText_url")<>"" then
										url=srs("imgText_url")
									else
										url="http://"&cur_host&"/admin/weixin/weixin_show.asp?id="&srs("id")
									end if		
									if instr(srs("imgText_Pic"),"http://")>0 then
										picurl=srs("imgText_Pic")
									else
										picurl="http://"&cur_host&srs("imgText_Pic")
									end if
								xmlhtml=xmlhtml&"<item>" &_
									"<Title><![CDATA["&srs("imgText_Title")&"]]></Title>" &_
									"<Description><![CDATA["&srs("imgText_Summary")&"]]></Description>" &_
									"<PicUrl><![CDATA["&picurl&"]]></PicUrl>" &_
									"<Url><![CDATA["&url&"]]></Url>" &_
									"</item>"
								srs.movenext
								loop
								xmlhtml=xmlhtml&"</Articles>"
							end if
							srs.close
						else					
							xmlhtml="<ArticleCount>1</ArticleCount>"
							xmlhtml=xmlhtml&"<Articles>"
							xmlhtml=xmlhtml&"<item>" &_
											"<Title><![CDATA["&os("imgText_Title")&"]]></Title>" &_
											"<Description><![CDATA["&os("imgText_Summary")&"]]></Description>" &_
											"<PicUrl><![CDATA["&picurl&"]]></PicUrl>" &_
											"<Url><![CDATA["&url&"]]></Url>" &_
											"</item>"
							xmlhtml=xmlhtml&"</Articles>"
						end if
						returnXml=wxObj.transmitImage(FromUserName,ToUserName,xmlhtml)
					else
						xmlhtml="此条回复消息丢失了！"
						returnXml=wxObj.transmitText(FromUserName,ToUserName,xmlhtml)					
					end if
					os.close
					'wxObj.WriteFile(returnXml)
					response.write returnXml
					
				elseif rs("reply_qanswerResource")<>"" then
							
					sqlxml="select top 1 Sucai_Title,Sucai_file,Sucai_desc,Sucai_source from weixin_Sucai where id in("&rs("reply_qanswerResource")&") and Sucai_status=0 order by Sucai_px desc"
					set os=conn.execute(sqlxml)
					if not os.eof then
						if os("Sucai_source")=1 then
							dim host
							host=Request.ServerVariables("SERVER_NAME")
							host="http://"&host&os("Sucai_file")
						else
							host=os("Sucai_file")
						end if
						xmlhtml=xmlhtml&"<Music>" &_
											"<Title><![CDATA["&os("Sucai_Title")&"]]></Title>" &_
											"<Description><![CDATA["&os("Sucai_desc")&"]]></Description>" &_
											"<MusicUrl><![CDATA["&os("Sucai_file")&"]]></MusicUrl>" &_
											"<HQMusicUrl><![CDATA["&host&"]]></HQMusicUrl>" &_
										"</Music>"
						returnXml=wxObj.transmitMusic(FromUserName,ToUserName,xmlhtml)
					else
						xmlhtml="此条回复消息丢失了！"
						returnXml=wxObj.transmitText(FromUserName,ToUserName,xmlhtml)					
					end if
					os.close
					'wxObj.WriteFile(returnXml)
					response.write returnXml
					
				else
					xmlhtml=""
					if instr(rs("reply_qanswerText"),",")>0 then
						dim arrText
						arrText=split(rs("reply_qanswerText"),",")
						xmlhtml=arrText(0)
					else
						xmlhtml=rs("reply_qanswerText")
					end if
					returnXml=wxObj.transmitText(FromUserName,ToUserName,xmlhtml)
					'wxObj.WriteFile(returnXml)
					response.write returnXml
				end if
			else
			
				if instr(wx_NoneReply,"[wx_news=")>0 then
					'wxObj.WriteFile(wx_NoneReply)
					newsid=wxObj.strCut(wx_NoneReply,"[wx_news=","]",2)
					if isnumeric(newsid) then
						'wxObj.WriteFile(newsid)
						set rs=conn.execute("select id,imgText_Title,imgText_Pic,imgText_url,imgText_Summary,imgText_Id_List from weixin_imageText where imgText_status=0 and id="&newsid&"")
						if not rs.eof then
							imgText_Id_List=trim(rs("imgText_Id_List")&" ")
							if imgText_Id_List<>"" then
								if instr(imgText_Id_List,",")>0 then
									arrN=split(imgText_Id_List,",")
									ncount=ubound(arrN)+1
								else
									ncount=1
								end if
										if rs("imgText_url")<>"" then
											url=rs("imgText_url")
										else
											url="http://"&cur_host&"/admin/weixin/weixin_show.asp?id="&rs("id")
										end if
										if instr(rs("imgText_Pic"),"http://")>0 then
											picurl=rs("imgText_Pic")
										else
											picurl="http://"&cur_host&rs("imgText_Pic")
										end if
								xmlhtml="<ArticleCount>"&ncount+1&"</ArticleCount>"
								xmlhtml=xmlhtml&"<Articles>"
								xmlhtml=xmlhtml&"<item>" &_
									"<Title><![CDATA["&rs("imgText_Title")&"]]></Title>" &_
									"<Description><![CDATA["&rs("imgText_Summary")&"]]></Description>" &_
									"<PicUrl><![CDATA["&picurl&"]]></PicUrl>" &_
									"<Url><![CDATA["&url&"]]></Url>" &_
									"</item>"
								set os=conn.execute("select id,imgText_Title,imgText_Pic,imgText_url,imgText_Summary from weixin_imageText where imgText_status=0 and id in("&imgText_Id_List&") order by imgText_px desc")
								if not os.eof then
									do while not os.eof
										if os("imgText_url")<>"" then
											url=os("imgText_url")
										else
											url="http://"&cur_host&"/admin/weixin/weixin_show.asp?id="&os("id")
										end if
										if instr(os("imgText_Pic"),"http://")>0 then
											picurl=os("imgText_Pic")
										else
											picurl="http://"&cur_host&os("imgText_Pic")
										end if
										xmlhtml=xmlhtml&"<item>" &_
										"<Title><![CDATA["&os("imgText_Title")&"]]></Title>" &_
										"<Description><![CDATA["&os("imgText_Summary")&"]]></Description>" &_
										"<PicUrl><![CDATA["&picurl&"]]></PicUrl>" &_
										"<Url><![CDATA["&url&"]]></Url>" &_
										"</item>"
									os.movenext
									loop
									xmlhtml=xmlhtml&"</Articles>"
								end if
								os.close
							else
										if rs("imgText_url")<>"" then
											url=rs("imgText_url")
										else
											url="http://"&cur_host&"/admin/weixin/weixin_show.asp?id="&rs("id")
										end if			
										if instr(rs("imgText_Pic"),"http://")>0 then
											picurl=rs("imgText_Pic")
										else
											picurl="http://"&cur_host&rs("imgText_Pic")
										end if			
								xmlhtml="<ArticleCount>1</ArticleCount>"
								xmlhtml=xmlhtml&"<Articles>"
								xmlhtml=xmlhtml&"<item>" &_
												"<Title><![CDATA["&rs("imgText_Title")&"]]></Title>" &_
												"<Description><![CDATA["&rs("imgText_Summary")&"]]></Description>" &_
												"<PicUrl><![CDATA["&picurl&"]]></PicUrl>" &_
												"<Url><![CDATA["&url&"]]></Url>" &_
												"</item>"
								xmlhtml=xmlhtml&"</Articles>"
							end if
							returnstr=wxObj.transmitImage(FromUserName,ToUserName,xmlhtml)
						else
							xmlhtml="此条回复消息丢失了！"
							returnstr=wxObj.transmitText(FromUserName,ToUserName,xmlhtml)
						end if
					else
						xmlhtml="此条回复消息丢失了！"	
						returnstr=wxObj.transmitText(FromUserName,ToUserName,xmlhtml)			
					end if
				else
					xmlhtml=wx_Subscribe
					returnstr=wxObj.transmitText(FromUserName,ToUserName,xmlhtml)
				'setCache FromUserName&"_do","subscribe",600
				end if
				'xmlhtml="你的问题太高深了，我还没学会哦"
				'returnXml=wxObj.transmitText(FromUserName,ToUserName,xmlhtml)
				'wxObj.WriteFile(returnXml)
				response.write returnstr
			end if
			rs.close
		end if

end sub

function BindWeiXin(strOpenID,strDatalist)
	on error resume next
	dim rs,arrList
	set rs=conn.execute("select id from WeixinBindList where OpenID='"&strOpenID&"'")
	if not rs.eof then
		BindWeiXin=1	'已绑定
	else
		dim returnMsg
		arrList=split(strDatalist,",")
		returnMsg=trim(GetURL("http://"&arrList(0)&"/admin/shangwin-login.asp?name="&arrList(1)&"&pass="&arrList(2)&""))
		WriteFile("http://"&arrList(0)&"/admin/shangwin-login.asp?name="&arrList(1)&"&pass="&arrList(2)&"")
		if returnMsg="1" then
			dim os
			set os=conn.execute("select i.iisid from iis_table i inner join zhangqian_CMC z on i.iisid=z.siteid where i.salestate >=7 and (i.webstate >= 6 or order_style=3 or order_style=1) and i.state=2 and '|'+z.domain+'|' like '%|"&arrList(0)&"|%'")
			if not os.eof then
				iisid=os("iisid")
				conn.execute("insert into WeixinBindList(iisid,OpenID) values("&iisid&",'"&strOpenID&"')")
				BindWeiXin=0	'绑定成功
			else
				BindWeiXin=2	'未上线或已经失效客户
			end if
		else
			BindWeiXin=3	'不是企帮商赢客户
		end if
	end if
	rs.close
	set rs=nothing
	if err then
		WriteFile(err.description)
		err.clear
		BindWeiXin=4	'绑定出错
	end if
end function

Function GetURL(url)
    On Error Resume Next 
    Set objXML = Server.CreateObject("MSXML2.ServerXMLHTTP")  
	objXML.open "GET",url,false  
	objXML.send()  
	GetURL=objXML.responseText
	if err then
		err.clear
		GetURL=""
	end if
End Function

Function BytesToBstr(Body, Cset)
    Dim Objstream
    Set Objstream = Server.CreateObject("adodb.stream")
    objstream.Type = 1
    objstream.Mode = 3
    objstream.Open
    objstream.Write body
    objstream.Position = 0
    objstream.Type = 2
    objstream.Charset = "utf-8"
    BytesToBstr = objstream.ReadText
    objstream.Close
    Set objstream = Nothing
End Function

function chkBind(strOpenID)
	on error resume next
	dim rs
	set rs=conn.execute("select id from WeixinBindList where OpenID='"&strOpenID&"'")
	WriteFile("select id from WeixinBindList where OpenID='"&strOpenID&"'")
	if not rs.eof then
		chkBind=true
	else
		chkBind=false
	end if
	rs.close
	set rs=nothing
	if err then
		err.clear
		chkBind=false
	end if
end function

private function CheckFields(FieldsName,TableName)
	dim blnFlag,chkStrSql,chkStrRs
	blnFlag=False
	chkStrSql="select * from "&TableName
	Set chkStrRs=Conn.Execute(chkStrSql)
	for i = 0 to chkStrRs.Fields.Count - 1
		if lcase(chkStrRs.Fields(i).Name)=lcase(FieldsName) then
			blnFlag=True
			Exit For
		else
			blnFlag=False
		end if
	Next
	CheckFields=blnFlag
End Function

	'写入文件法调试
	public Function WriteFile(content)
		dim fso,fopen,filepath
		on error resume next
		filepath=server.mappath(".")&"\wx.txt"
		Set fso = Server.CreateObject("scripting.FileSystemObject")
		set fopen=fso.OpenTextFile(filepath, 8 ,true)
		content = content&vbcrlf&"************line seperate("&now()&")*****************"
		fopen.writeline(content)
		if err then
			response.write err.description
			err.clear
			response.end
		end if
		set fso=nothing
		set fopen=Nothing
		
	End Function

	
	sub setCache(cacheName,cacheValue,cacheTime)
		set myCache=New Cache ' 删除某以缓存 
		myCache.name=cacheName			'定义缓存名
		myCache.makeEmpty()
		if myCache.valid then			'判断是否可用(包括过期，与是否为空值)
		else
			myCache.add cacheValue,dateadd("s",cacheTime,now)	'写入缓存 xxx.add 内容,过期时间
		end if
		'myCache.makeEmpty()	'释放内存
		set myCache=nothing
	end sub

	function getCache(cacheName)
		set myCache=New Cache 
		myCache.name=cacheName			'定义缓存名
		if myCache.valid then
			getCache=myCache.value
		else
			getCache=""
		end if
		set myCache=nothing
	end function
%>
<script language="jscript" runat="server">  
	Array.prototype.get = function(x) { return this[x]; };  
	function parseJSON(strJSON) { return eval("(" + strJSON + ")"); }  
</script>
<!--#Include File="../../Code.asp"-->