<%
'==========================================
'文 件 名：System.asp
'文件用途：系统信息文件
'版权所有：深圳企帮
'==========================================
Dim FkSystemName,FkSystemNameEn,FkSystemVersion
FkSystemName="企帮-商赢快车"
FkSystemNameEn="Qebang-shangwin"
FkSystemVersion=SysVersion

Dim arrTips		'提示信息语言数组
if lcase(left(SiteTemplate,3))="en_" then 

		arrTips=array("Failed to get the home page template !","Failed to get the template","Home","under construction","No Previous","No Next","Effective long-term","Day","Awaiting reply","Please do not send spam !","This item was not found !","Please enter the advisory title !","Advisory title can not exceed 50 characters !","Please enter the Advisory Content !","Consultation content can not exceed 500 characters !","Please enter your name !","Your name can not exceed 50 characters !","Please enter your contact !","Contact can not exceed 50 characters !","System parameter error","Consulting successfully submitted,We will reply as soon as possible !","Do not re-submitted to the Advisory","Please enter your search keywords !","Search keywords is not exceed 50 characters !")
		
elseif lcase(left(SiteTemplate,6))="japan_" then 

		arrTips=array("ホームページテンプレート取得失敗 !","テンプレートを失","トップページ","建設の中で","なし前","なし次","長期にわたり有効","デイ","返事を待っている","スパムを送信しないでください !","この機能の項を探し当てていません !","入力して見出しを問合せすることを下さい !","見出しを問合せして50のキャラクターより大きいことができません !","問い合わせの内容を入力して下さい !","問い合わせの内容は500のキャラクターより大きいことができません !","あなたの姓名を入力して下さい !","あなたの姓名は50のキャラクターより大きいことができません !","連絡方法を入力して下さい !","連絡方法は50のキャラクターより大きいことができません !","システムパラメタは誤った","問い合わせは成功に提出します,私達はできてできるだけ早く返答します !","繰り返し問い合わせに提出しないでください","キーワードを入力してください !","キーワードを検索して50のキャラクターより大きいことができません !")
		
else

		arrTips=array("首页模板获取失败！","模板获取失败","首页","建设中","无上一篇","无下一篇","长期有效","天","待回复","请勿发垃圾信息！","没有找到此功能项！","请输入咨询标题 !","咨询标题不能大于50个字符！","请输入咨询内容！","咨询内容不能大于500个字符！","请输入您的姓名！","您的姓名不能大于50个字符！","请输入联系方式！","联系方式不能大于50个字符！","系统参数错误！","咨询提交成功，我们会尽快回复！","请勿重复提交咨询！","请输入关键字！","搜索关键词不能大于50个字符")

end if
%>