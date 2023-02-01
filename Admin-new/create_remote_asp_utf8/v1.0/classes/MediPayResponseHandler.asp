<!--#include file="../util/md5.asp"-->
<!--#include file="../util/tenpay_util.asp"-->
<%
'
'�Ƹ�ͨ�н鵣��֧��Ӧ����
'============================================================================
'api˵����
'getKey()/setKey(),��ȡ/������Կ
'getParameter()/setParameter(),��ȡ/���ò���ֵ
'getAllParameters(),��ȡ���в���
'isTenpaySign(),�Ƿ�Ƹ�ͨǩ��,true:�� false:��
'doShow(),��ʾ�������
'getDebugInfo(),��ȡdebug��Ϣ
'
'============================================================================
'

Class MediPayResponseHandler

	'��Կ
	Private key

	'Ӧ��Ĳ���
	Private parameters

	'debug��Ϣ
	Private debugInfo

	'��ʼ���캯��
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

	'��ȡ��Կ
	Public Function getKey()
		getKey = key
	End Function
	
	'������Կ
	Public Function setKey(key_)
		key = key_
	End Function
	
	'��ȡ����ֵ
	Public Function getParameter(parameter)
		getParameter = parameters.Item(parameter)
	End Function
	
	'���ò���ֵ
	Public Function setParameter(parameter, parameterValue)
		If parameters.Exists(parameter) = True Then
			parameters.Remove(parameter)
		End If
		parameters.Add parameter, parameterValue	
	End Function

	'��ȡ��������Ĳ���,����Scripting.Dictionary
	Public Function getAllParameter()
		'��������
		SortDictionary parameters, dictKey

		getAllParameter = parameters
	End Function

	'�Ƿ�Ƹ�ͨǩ��
	'true:�� false:��
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
		
		//����ĸa-z����
		sortArrayAZ signParameterArray
	
		'��������
		SortDictionary parameters, dictKey

		Dim signPars
		Dim index
		For index = 0 To UBound(signParameterArray)
			Dim k
			Dim v
			k = signParameterArray(index)
			v = getParameter(k)
			'���ַ������μ�ǩ��
			If v <> "" Then
				signPars = signPars & k & "=" & v & "&"
			End If
		Next
		
		'��Կ�������
		signPars = signPars & "key=" & key

		Dim sign
		sign= LCase(ASP_MD5(signPars))

		Dim tenpaySign
		tenpaySign = LCase( getParameter("sign"))

		'debugInfo
		debugInfo = signPars & " => sign:" & sign & " tenpaySign:" & tenpaySign

		isTenpaySign = (sign = tenpaySign)

	End Function

	'��ʾ�������,���metaֵ
	Public Function doShow()
		Dim strHtml
		strHtml = "<html><head>" &_
			"<meta name=""TENCENT_ONELINE_PAYMENT""" &_
				"content=""China TENCENT"">" &_
			"</head><body></body></html>"
		
		Response.Write(strHtml)

		Response.End

	End Function

	'��ȡdebug��Ϣ
	Function getDebugInfo()
		getDebugInfo = debugInfo
	End Function
	
End Class




%>