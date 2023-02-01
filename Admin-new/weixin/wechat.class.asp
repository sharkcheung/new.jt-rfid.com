<%
class WeixinClass

	private api_url_prefix,auth_url,menu_create_url,menu_get_url,menu_delete_url,media_get_url,qrcode_create_url,qr_scene,qr_limit_scene,qrcode_img_url,user_get_url,user_info_url,group_get_url,group_create_url,group_update_url,group_member_update_url
	
	private strToken,strSignature,strTimestamp,strNonce
	public errCode 
	public errMsg 
	public token
	
	
	Public Property Let SetToken(ByVal strVarToken)
		strToken=strVarToken
	end Property
	
	'//---Class_Initialize()是类的初始化事件，类被调用，首先会触发该部分的执行，一般用来初始化默认值.
	Private Sub Class_Initialize() 
		api_url_prefix = "https://api.weixin.qq.com/cgi-bin"
		auth_url = "/token?grant_type=client_credential&"
		menu_create_url = "/menu/create?"
		menu_get_url = "/menu/get?"
		menu_delete_url = "/menu/delete?"
		media_get_url = "/media/get?"
		qrcode_create_url="/qrcode/create?"
		qr_scene = 0
		qr_limit_scene = 1
		qrcode_img_url="https://mp.weixin.qq.com/cgi-bin/showqrcode?ticket="
		user_get_url="/user/get?"
		user_info_url="/user/info?"
		group_get_url="/groups/get?"
		group_create_url="/groups/create?"
		group_update_url="/groups/update?"
		group_member_update_url="/groups/members/update?"
		errCode= 40001
		errMsg= "no access"
	End Sub 
	
	'**
	' * For weixin server validation 
	'**/	
	private function checkSignature()
	
        strSignature= request("signature")
        strTimestamp= request("timestamp")
        strNonce 	= request("nonce")
		dim tmpArr,strtmpStr
		tmpArr = array(strToken, strTimestamp, strNonce)
		arg_sort(tmpArr)
		strtmpStr = sha1(replace(join(arg_sort(tmpArr))," ",""))
	
WriteFile(strSignature&"-"&strTimestamp&"-"&strNonce&"-"&strVarToken&"-"&strtmpStr)
		if  strtmpStr = strSignature  then
			checkSignature= true
		else
			checkSignature= false
		end if
	end function
	

	public function valid()
		on error resume next
   		if (checkSignature()) then
   			die(request("echostr"))
   		else 
   			die("no access")
       	end if
		if err then
			err.clear
			WriteFile(err.description)
   			die("no access")
		end if
    end function
	
	'**********
	'*将xml转换成数组
	'**********
	public function getXml_toString()
		'on error resume next
		dim ToUserName,FromUserName,MsgType,Content,xml_dom,eve,createtime,EventKey,PicUrl,MediaId,Format,ThumbMediaId,Label,Location_X,Location_Y,Scale
		set xml_dom = Server.CreateObject("MSXML2.DOMDocument")'此处根据您的实际服务器情况改写
		xml_dom.load request
		ToUserName=xml_dom.getelementsbytagname("ToUserName").item(0).text
		FromUserName=xml_dom.getelementsbytagname("FromUserName").item(0).text
		MsgType=xml_dom.getelementsbytagname("MsgType").item(0).text
		if MsgType="event" then
			eve=xml_dom.getelementsbytagname("Event").item(0).text
			createtime=xml_dom.getelementsbytagname("CreateTime").item(0).text
			if eve="CLICK" then 
				EventKey=xml_dom.getelementsbytagname("EventKey").item(0).text		
'WriteFile(EventKey)
				getXml_toString=FromUserName&"{$}"&ToUserName&"{$}"&MsgType&"{$}"&eve&"{$}"&createtime&"{$}"&EventKey
			else
				getXml_toString=FromUserName&"{$}"&ToUserName&"{$}"&MsgType&"{$}"&eve&"{$}"&createtime&"{$}"
			end if	
		elseif MsgType="text" then
			Content=xml_dom.getelementsbytagname("Content").item(0).text
			getXml_toString=FromUserName&"{$}"&ToUserName&"{$}"&MsgType&"{$}"&Content
		elseif MsgType="image" then
			PicUrl=xml_dom.getelementsbytagname("PicUrl").item(0).text
			MediaId=xml_dom.getelementsbytagname("MediaId").item(0).text
			MsgId=xml_dom.getelementsbytagname("MsgId").item(0).text
			getXml_toString=FromUserName&"{$}"&ToUserName&"{$}"&MsgType&"{$}"&PicUrl&"{$}"&MediaId&"{$}"&MsgId
		elseif MsgType="voice" then
			MediaId=xml_dom.getelementsbytagname("MediaId").item(0).text
			Format=xml_dom.getelementsbytagname("Format").item(0).text
			MsgId=xml_dom.getelementsbytagname("MsgId").item(0).text
			getXml_toString=FromUserName&"{$}"&ToUserName&"{$}"&MsgType&"{$}"&Format&"{$}"&MediaId&"{$}"&MsgId
		elseif MsgType="video" then
			MediaId=xml_dom.getelementsbytagname("MediaId").item(0).text
			MsgId=xml_dom.getelementsbytagname("MsgId").item(0).text
			ThumbMediaId=xml_dom.getelementsbytagname("ThumbMediaId").item(0).text
			getXml_toString=FromUserName&"{$}"&ToUserName&"{$}"&MsgType&"{$}"&MediaId&"{$}"&MsgId&"{$}"&ThumbMediaId
		elseif MsgType="location" then
			Location_X=xml_dom.getelementsbytagname("Location_X").item(0).text
			Location_Y=xml_dom.getelementsbytagname("Location_Y").item(0).text
			Scale=xml_dom.getelementsbytagname("Scale").item(0).text
			Label=xml_dom.getelementsbytagname("Label").item(0).text
			MsgId=xml_dom.getelementsbytagname("MsgId").item(0).text
			getXml_toString=FromUserName&"{$}"&ToUserName&"{$}"&MsgType&"{$}"&Location_X&"{$}"&Location_Y&"{$}"&Scale&"{$}"&Label&"{$}"&MsgId
		elseif MsgType="link" then
			Title=xml_dom.getelementsbytagname("Title").item(0).text
			Description=xml_dom.getelementsbytagname("Description").item(0).text
			Url=xml_dom.getelementsbytagname("Url").item(0).text
			MsgId=xml_dom.getelementsbytagname("MsgId").item(0).text
			getXml_toString=FromUserName&"{$}"&ToUserName&"{$}"&MsgType&"{$}"&Title&"{$}"&Description&"{$}"&Url&"{$}"&MsgId
		end if

		set xml_dom=nothing
		if err then
			getXml_toString=err.description
			err.clear
		end if
	end function
	
	Function strCut(strContent,StartStr,EndStr,CutType)
		Dim strHtml,S1,S2
		strHtml = strContent
		On Error Resume Next
		Select Case CutType
		Case 1
			S1 = InStr(strHtml,StartStr)
			S2 = InStr(S1,strHtml,EndStr)+Len(EndStr)
		Case 2
			S1 = InStr(strHtml,StartStr)+Len(StartStr)
			S2 = InStr(S1,strHtml,EndStr)
		End Select
		If Err Then
			strCute = ""
			Err.Clear
			Exit Function
		Else
			strCut = Mid(strHtml,S1,S2-S1)
		End If
	End Function
	
	'写入文件法调试
	public Function WriteFile(content)
		filepath=server.mappath(".")&"\wx.txt"
		Set fso = Server.CreateObject("scripting.FileSystemObject")
		set fopen=fso.OpenTextFile(filepath, 8 ,true)
		content = content&vbcrlf&"************line seperate("&now()&")*****************"
		fopen.writeline(content)
		set fso=nothing
		set fopen=Nothing
	End Function
	
	public function die(str)
		response.write str
		response.end
	end function
	
	private function IIF(strExpression,v1,v2)
		if strExpression then
			IIF=v1
		else
			IIF=v2
		end if
	end function
	
	
	' /**
	 ' * 数据XML编码
	 ' * @param mixed arrData 数据
	 ' * @return string
	 ' */
	public function transmitText(strFromUserName,strToUserName,strContent) 
		dim xml
		xml="<xml>" &_
 	"<ToUserName><![CDATA["&strFromUserName&"]]></ToUserName>" &_
	"<FromUserName><![CDATA["&strToUserName&"]]></FromUserName>" &_
	"<CreateTime>"&getTimeStamp()&"</CreateTime>" &_
	"<MsgType><![CDATA[text]]></MsgType>" &_
	"<Content><![CDATA[" & strContent & "]]></Content>" &_
	"<FuncFlag>0<FuncFlag>" &_
	"</xml>"
		transmitText=xml
	end function	
	
	
	' /**
	 ' * 数据XML编码
	 ' * @param mixed arrData 数据
	 ' * @return string
	 ' */
	public function transmitImage(strFromUserName,strToUserName,strContent) 
		dim xml
		xml="<xml>" &_
 	"<ToUserName><![CDATA["&strFromUserName&"]]></ToUserName>" &_
	"<FromUserName><![CDATA["&strToUserName&"]]></FromUserName>" &_
	"<CreateTime>"&getTimeStamp()&"</CreateTime>" &_
	"<MsgType><![CDATA[news]]></MsgType>" &_
	 strContent &_
	"<FuncFlag>0<FuncFlag>" &_
	"</xml>"
		transmitImage=xml
	end function	
	
	
	' /**
	 ' * 数据XML编码
	 ' * @param mixed arrData 数据
	 ' * @return string
	 ' */
	public function transmitMusic(strFromUserName,strToUserName,strContent) 
		dim xml
		xml="<xml>" &_
 	"<ToUserName><![CDATA["&strFromUserName&"]]></ToUserName>" &_
	"<FromUserName><![CDATA["&strToUserName&"]]></FromUserName>" &_
	"<CreateTime>"&getTimeStamp()&"</CreateTime>" &_
	"<MsgType><![CDATA[music]]></MsgType>" &_
	 strContent &_
	"<FuncFlag>0<FuncFlag>" &_
	"</xml>"
		transmitMusic=xml
	end function
	
private Function getTimeStamp()

	getTimeStamp = DateDiff("s", "1970-1-1 8:00:00", Now())

End Function
	
'/**
'* 判断是否是QQ表情
'* 
'* @param content
'* @return
'*/
public function isQqFace(content)
	isQqFace = false
	on error resume next
	'// 判断QQ表情的正则表达式
	dim qqfaceRegex,regEx
	qqfaceRegex= "/::\)"
	'qqfaceRegex= "/::\\)|/::~|/::B|/::\\||/:8-\\)|/::<|/::$|/::X|/::Z|/::'\\(|/::-\\||/::@|/::P|/::D|/::O|/::\\(|/::\\+|/:--b|/::Q|/::T|/:,@P|/:,@-D|/::d|/:,@o|/::g|/:\\|-\\)|/::!|/::L|/::>|/::,@|/:,@f|/::-S|/:\\?|/:,@x|/:,@@|/::8|/:,@!|/:!!!|/:xx|/:bye|/:wipe|/:dig|/:handclap|/:&-\\(|/:B-\\)|/:<@|/:@>|/::-O|/:>-\\||/:P-\\(|/::'\\||/:X-\\)|/::\\*|/:@x|/:8\\*|/:pd|/:<W>|/:beer|/:basketb|/:oo|/:coffee|/:eat|/:pig|/:rose|/:fade|/:showlove|/:heart|/:break|/:cake|/:li|/:bome|/:kn|/:footb|/:ladybug|/:shit|/:moon|/:sun|/:gift|/:hug|/:strong|/:weak|/:share|/:v|/:@\\)|/:jj|/:@@|/:bad|/:lvu|/:no|/:ok|/:love|/:<L>|/:jump|/:shake|/:<O>|/:circle|/:kotow|/:turn|/:skip|/:oY|/:#-0|/:hiphot|/:kiss|/:<&|/:&>"
	Set regEx = New RegExp
	regEx.Pattern = "/::\)|/::~|/::B|/::\||/:8-\)|/::<|/::$|/::X|/::Z|/::'\(|/::-\||/::@|/::P|/::D|/::O|/::\(|/::\+|/:--b|/::Q|/::T|/:,@P|/:,@-D|/::d|/:,@o|/::g|/:\|-\)|/::!|/::L|/::>|/::,@|/:,@f|/::-S|/:\?|/:,@x|/:,@@|/::8|/:,@!|/:!!!|/:xx|/:bye|/:wipe|/:dig|/:handclap|/:&-\(|/:B-\)|/:<@|/:@>|/::-O|/:>-\||/:P-\(|/::'\||/:X-\)|/::\*|/:@x|/:8\*|/:pd|/:<W>|/:beer|/:basketb|/:oo|/:coffee|/:eat|/:pig|/:rose|/:fade|/:showlove|/:heart|/:break|/:cake|/:li|/:bome|/:kn|/:footb|/:ladybug|/:shit|/:moon|/:sun|/:gift|/:hug|/:strong|/:weak|/:share|/:v|/:@\)|/:jj|/:@@|/:bad|/:lvu|/:no|/:ok|/:love|/:<L>|/:jump|/:shake|/:<O>|/:circle|/:kotow|/:turn|/:skip|/:oY|/:#-0|/:hiphot|/:kiss|/:<&|/:&>"
	regEx.IgnoreCase = True
	regEx.Global = True
	if regEx.Test(content) then  '如果包含，则……
		isQqFace=true
	end if
	if err then
		die(err.description)
	end if
end function
	
	'对数组排序
	'sArray 排序前的数组
	'输出 排序后的数组
	function arg_sort(sArray)
		nCount = ubound(sArray)
		For i = nCount TO 0 Step -1
			minmax = sArray( 0 )
			minmaxSlot = 0
			For j = 1 To i
				mark = (sArray( j ) > minmax)
				If mark Then 
					minmax = sArray( j )
					minmaxSlot = j
				End If
			Next
			If minmaxSlot <> i Then 
				temp = sArray( minmaxSlot )
				sArray( minmaxSlot ) = sArray( i )
				sArray( i ) = temp
			End If
		Next
		arg_sort = sArray
	end function

	public Function Sort(Ary) 
		Dim KeepChecking,I,FirstValue,SecondValue 
		KeepChecking = TRUE 
		Do Until KeepChecking = FALSE 
			KeepChecking = FALSE 
			For I = 0 To UBound(Ary) 
				If I = UBound(Ary) Then Exit For 
				If Ary(I) > Ary(I+1) Then 
				FirstValue = Ary(I) 
				SecondValue = Ary(I+1) 
				Ary(I) = SecondValue 
				Ary(I+1) = FirstValue 
				KeepChecking = TRUE 
				End If 
			Next 
		Loop 
		Sort = Ary 
	End Function
	
end class
%>