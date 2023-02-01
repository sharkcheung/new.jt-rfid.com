<%
	'类名：alipay_service
	'功能：支付宝外部服务接口控制
	'详细：该页面是请求参数核心处理文件，不需要修改
	'版本：3.0
	'修改日期：2010-07-26
	'说明：
	'以下代码只是为了方便商户测试而提供的样例代码，商户可以根据自己网站的需要，按照技术文档编写,并非一定要使用该代码。
	'该代码仅供学习和研究支付宝接口使用，只是提供一个参考
%>

<!--#include file="alipay_function.asp"-->

<%

dim gateway			'网关地址
dim mysign			'加密结果（签名结果）
dim sPara		'需要加密的已经过滤后的参数数组

'********************************************************************************

'构造函数
'从配置文件及入口文件中初始化变量
'inputPara 需要加密的参数数组
function alipay_service(inputPara)
	gateway = "https://www.alipay.com/cooperate/gateway.do?"
	sPara = para_filter(inputPara)
	sort_para = arg_sort(sPara)		'得到从字母a到z排序后的加密参数数组
	'获得签名结果
	mysign = build_mysign(sort_para,key,sign_type,input_charset)
end function

'********************************************************************************

'构造请求URL（GET方式请求）
'输出 请求url
function create_url()
	url = gateway
	sort_para = arg_sort(sPara)
	arg = create_linkstring_urlencode(sort_para)	'把数组所有元素，按照“参数=参数值”的模式用“&”字符拼接成字符串
	url = url & arg & "sign=" &mysign & "&sign_type=" & sign_type
	create_url = url
end function

'********************************************************************************

'构造Post表单提交HTML（POST方式请求）
'输出 表单提交HTML文本
function build_postform()
	sHtml = "<form id='alipaysubmit' name='alipaysubmit' action='"& gateway &"_input_charset="&input_charset&"' method='post'>"

	nCount = ubound(sPara)
	for i = 0 to nCount
		'把sArray的数组里的元素格式：变量名=值，分割开来
		pos = Instr(sPara(i),"=")			'获得=字符的位置
		nLen = Len(sPara(i))				'获得字符串长度
		itemName = left(sPara(i),pos-1)		'获得变量名
		itemValue = right(sPara(i),nLen-pos)'获得变量的值
		
		sHtml = sHtml & "<input type='hidden' name='"& itemName &"' value='"& itemValue &"'/>"
	next

	sHtml = sHtml & "<input type='hidden' name='sign' value='"& mysign &"'/>"
	sHtml = sHtml & "<input type='hidden' name='sign_type' value='"& sign_type &"'/></form>"

	sHtml = sHtml & "<input type=""button"" name=""v_action"" value=""支付宝确认付款"" onClick=""document.forms['alipaysubmit'].submit();"">"
	build_postform = sHtml
end function

%>