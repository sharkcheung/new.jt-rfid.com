<!--#include file="../util/md5.asp"-->
<!--#include file="../util/tenpay_util.asp"-->
<%
'
'财付通中介担保支付请求类
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

Class MediPayRequestHandler
	
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

		'自定参数，原样返回
		parameters.Add "attach", ""
		
		'平台商帐号
		parameters.Add "chnid", ""
		
		'任务代码
		parameters.Add "cmdno", "12"
		
		'编码类型 1:gbk 2:utf-8
		parameters.Add "encode_type", "1"
		
		'交易说明，不能包含<>’”%特殊字符
		parameters.Add "mch_desc", ""
		
		'商品名称，不能包含<>’”%特殊字符
		parameters.Add "mch_name", ""
		
		'商品总价，单位为分。
		parameters.Add "mch_price", ""
		
		'回调通知URL
		parameters.Add "mch_returl", ""
		
		'交易类型：1、实物交易，2、虚拟交易
		parameters.Add "mch_type", ""
		
		'商家的定单号
		parameters.Add "mch_vno", ""
		
		'是否需要在财付通填定物流信息，1：需要，2：不需要。
		parameters.Add "need_buyerinfo", ""
		
		'卖家财付通帐号
		parameters.Add "seller", ""
		
		'支付后的商户支付结果展示页面
		parameters.Add "show_url", ""
		
		'物流公司或物流方式说明
		parameters.Add "transport_desc", ""
		
		'需买方另支付的物流费用
		parameters.Add "transport_fee", ""
		
		'版本号
		parameters.Add "version", "2"
		
		'摘要
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

	'获取带参数的请求URL
	Public Function getRequestURL()

		Call createSign()
		
		Dim reqPars
		Dim k
		For Each k In parameters
			reqPars = reqPars & k & "=" & Server.URLEncode(parameters(k)) & "&" 
		Next
		
		'去掉最后一个&
		reqPars = Left(reqPars, Len(reqPars)-1)

		getRequestURL = getGateURL & "?" & reqPars

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
	Private Sub createSign()
		'按键排序
		SortDictionary parameters, dictKey

		Dim signPars
		Dim k
		For Each k In parameters
			Dim v
			v = parameters(k)
			'空字符串不参加签名
			If v <> "" And k <> "sign" Then
				signPars = signPars & k & "=" & v & "&"
			End If
		Next
		
		'密钥在最后面
		signPars = signPars & "key=" & key

		Dim sign
		sign= LCase(ASP_MD5(signPars))

		setParameter "sign", sign

		'debuginfo
		debugInfo = signPars & " => sign:" & sign

	End Sub
	
End Class


%>