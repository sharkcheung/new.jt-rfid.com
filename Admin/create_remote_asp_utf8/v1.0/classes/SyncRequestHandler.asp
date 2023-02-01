<!--#include file="../../../../inc/md5.asp"-->
<!--#include file="../util/tenpay_util.asp"-->
<%
'
'企帮客服版网站信息同步请求类
'============================================================================
'api说明：
'init(),初始化函数，默认给一些参数赋值，如cmdno,date等。
'getGateURL()/setGateURL(),获取/设置入口地址,不包含参数值
'getKey()/setKey(),获取/设置密钥
'getParameter()/setParameter(),获取/设置参数值
'getAllParameters(),获取所有参数
'getRequestURL(),获取带参数的请求URL
'doSend(),重定向到财付通支付
'getDebugInfo(),获取debug信息
'
'============================================================================
'

Class SyncRequestHandler
	'网关url地址
	Private gateUrl
	
	'密钥
	Private key
	
	'请求的参数
	Private parameters
	
	'debug信息
	Private debugInfo
	
	'初始构造函数
	Private Sub class_initialize()
		gateUrl = "http://service.tenpay.com/cgi-bin/v3.0/payservice.cgi"
		key = ""
		Set parameters = Server.CreateObject("Scripting.Dictionary")
		debugInfo = ""
	End Sub

	'初始化函数，默认给一些参数赋值，如cmdno,date等。
	Public Function init()
		parameters.RemoveAll
		parameters.Add "sign", ""
	End Function

	'获取入口地址,不包含参数值
	Public Function getGateURL()
		getGateURL = gateUrl
	End Function
	
	'设置入口地址,不包含参数值
	Public Function setGateURL(gateUrl_)
		gateUrl = gateUrl_
	End Function

	'获取密钥
	Public Function getKey()
		getKey = key
	End Function
	
	'设置密钥
	Public Function setKey(key_)
		key = key_
	End Function
	
	'获取参数值
	Public Function getParameter(parameter)
		getParameter = parameters.Item(parameter)
	End Function
	
	'设置参数值
	Public Function setParameter(parameter, parameterValue)
		If parameters.Exists(parameter) = True Then
			parameters.Remove(parameter)
		End If
		parameters.Add parameter, parameterValue	
	End Function

	'获取所有请求的参数,返回Scripting.Dictionary
	Public Function getAllParameters()
			
		'按键排序
		SortDictionary parameters, dictKey

		getAllParameters = parameters
	End Function

	'获取参数
	Public Function getParameters()

		Call createSign()
		
		Dim reqPars
		Dim k
		For Each k In parameters
			reqPars = reqPars & k & "=" & Server.URLEncode(parameters(k)) & "&" 
		Next
		
		'去掉最后一个&
		reqPars = Left(reqPars, Len(reqPars)-1)

		getParameters =  reqPars

	End Function
	
	'获取debug信息
	Public Function getDebugInfo()
		getDebugInfo = debugInfo
	End Function
	
	'重定向到财付通支付
	Public Function doSend()
		Response.Redirect(getRequestURL())
		Response.End
	End Function

	'创建签名
	Sub createSign()
		'按键排序
		SortDictionary parameters, dictKey

		Dim signPars
		Dim k
		For Each k In parameters
			Dim v
			v = parameters(k)
			'空字符串不参加签名
			If v <> "" And k <> "sign" Then
				signPars = signPars & k & "=" & Server.URLEncode(v) & "&"
			End If
		Next
		'密钥在最后面
		signPars = signPars & "key=" & key
		Dim sign
		sign= LCase(MD5(signPars,32))
		setParameter "sign", sign
		'response.write "-------"&sign&"-"

		'debuginfo
		debugInfo = signPars & " => sign:" & sign

	End Sub
	
	Function PostHttpPage(RefererUrl,PostUrl,PostData)
		Dim xmlHttp
		Dim RetStr
		Set xmlHttp = CreateObject("Msxml2.XMLHTTP")
		xmlHttp.Open "POST", PostUrl, false
		XmlHTTP.setRequestHeader "Content-Length",Len(PostData)
		xmlHttp.setRequestHeader "Content-Type", "application/x-www-form-urlencoded"
		xmlHttp.setRequestHeader "Referer", RefererUrl
		xmlHttp.Send PostData
		If Err.Number <> 0 Then
			Set xmlHttp=Nothing
			PostHttpPage = "$False$"
			Exit Function
		End If
		PostHttpPage=bytesToBSTR(xmlHttp.responseBody,"UTF-8")
		Set xmlHttp = nothing
	End Function

	private function BytesToBstr(Body,Cset)
		Dim Objstream
		Set Objstream = Server.CreateObject("adodb.stream")
		objstream.Type = 1
		objstream.Mode =3
		objstream.Open
		objstream.Write body
		objstream.Position = 0
		objstream.Type = 2
		objstream.Charset = Cset
		BytesToBstr = objstream.ReadText
		objstream.Close
		set objstream = nothing
	End Function
End Class

'数组排序
private Function sortArrayAZ(ByRef array)
	Dim min
	For i = 0 To UBound(array)-1
		min = i
		For j = i+1 To UBound(array)
			If StrComp(array(j), array(i)) < 0 Then
				min = j
			End if
		Next
		
		'swap
		temp = array(i)
		array(i) = array(min)
		array(min) = temp
		
	Next
		
End Function
%>