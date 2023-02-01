<!--#include file="../util/md5.asp"-->
<!--#include file="../util/tenpay_util.asp"-->
<%
'
'财付通中介担保支付应答类
'============================================================================
'api说明：
'getKey()/setKey(),获取/设置密钥
'getParameter()/setParameter(),获取/设置参数值
'getAllParameters(),获取所有参数
'isTenpaySign(),是否财付通签名,true:是 false:否
'doShow(),显示处理结果
'getDebugInfo(),获取debug信息
'
'============================================================================
'

Class MediPayResponseHandler

	'密钥
	Private key

	'应答的参数
	Private parameters

	'debug信息
	Private debugInfo

	'初始构造函数
	Private Sub class_initialize()
		key = ""
		Set parameters = Server.CreateObject("Scripting.Dictionary")
		debugInfo = ""
		
		parameters.RemoveAll
		
		Dim k
		Dim v
		
		'GET
		For Each k In Request.QueryString
			v = Request.QueryString(k)
			setParameter k,v
		Next
		
		'POST
		For Each k In Request.Form
			v = Request(k)
			setParameter k,v
		Next
		
	End Sub

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
	Public Function getAllParameter()
		'按键排序
		SortDictionary parameters, dictKey

		getAllParameter = parameters
	End Function

	'是否财付通签名
	'true:是 false:否
	Public Function isTenpaySign()
	
		ReDim signParameterArray(13)
		signParameterArray(0) = "attach"
		signParameterArray(1) = "buyer_id"
		signParameterArray(2) = "cft_tid"		
		signParameterArray(3) = "chnid"
		signParameterArray(4) = "cmdno"
		signParameterArray(5) = "mch_vno"
		signParameterArray(6) = "retcode"
		signParameterArray(7) = "seller"
		signParameterArray(8) = "status"
		signParameterArray(9) = "total_fee"
		signParameterArray(10) = "trade_price"
		signParameterArray(11) = "transport_fee"
		signParameterArray(12) = "version"
		
		//按字母a-z排序
		sortArrayAZ signParameterArray
	
		'按键排序
		SortDictionary parameters, dictKey

		Dim signPars
		Dim index
		For index = 0 To UBound(signParameterArray)
			Dim k
			Dim v
			k = signParameterArray(index)
			v = getParameter(k)
			'空字符串不参加签名
			If v <> "" Then
				signPars = signPars & k & "=" & v & "&"
			End If
		Next
		
		'密钥在最后面
		signPars = signPars & "key=" & key

		Dim sign
		sign= LCase(ASP_MD5(signPars))

		Dim tenpaySign
		tenpaySign = LCase( getParameter("sign"))

		'debugInfo
		debugInfo = signPars & " => sign:" & sign & " tenpaySign:" & tenpaySign

		isTenpaySign = (sign = tenpaySign)

	End Function

	'显示处理结果,输出meta值
	Public Function doShow()
		Dim strHtml
		strHtml = "<html><head>" &_
			"<meta name=""TENCENT_ONELINE_PAYMENT""" &_
				"content=""China TENCENT"">" &_
			"</head><body></body></html>"
		
		Response.Write(strHtml)

		Response.End

	End Function

	'获取debug信息
	Function getDebugInfo()
		getDebugInfo = debugInfo
	End Function
	
End Class




%>