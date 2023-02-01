//==========================================
//系统开发：深圳企帮
//http://www.qebang.cn/
//==========================================

//==========================================
//函数名：ShowBox
//用途：操作框弹?
//参数?
//==========================================


function ShowBox(DoUrl,strTitle){
	var layer1=layer.load(2); //又换了种风格，并且设定最长等待10秒 
	var sbwidth= arguments[2] || '700px';
	
	var sbheight= arguments[3] || 'auto';
	$.get(DoUrl,
		function(data){
			layer.closeAll('loading');
			//document.getElementById('BoxContent').innerHTML=data;
			layer.open({
				type: 1,
				title: strTitle,
				shadeClose: true,
				shade: 0.5,
				area: [sbwidth, sbheight],
				zIndex:88888,
				content: data
			});
			if($(".kinediter").length>0){	
				var editor;
				try
				{					
					KindEditor.create('.kinediter', {
						themeType		: 'simple',
						uploadJson		: '/admin/dkidtioenr/aps/upload_json.asp',
						fileManagerJson : '/admin/dkidtioenr/aps/file_manager_json.asp',
						allowFileManager: true,
						filterMode		: true,
						items			:[
											'source', '|', 'undo', 'redo', '|', 'preview', 'print', 'template', 'cut', 'copy', 'paste',
											'plainpaste', 'wordpaste', '|', 'justifyleft', 'justifycenter', 'justifyright',
											'justifyfull', 'insertorderedlist', 'insertunorderedlist', 'indent', 'outdent', 'subscript',
											'superscript', 'clearhtml', 'quickformat', 'selectall', '|', 'fullscreen', '/',
											'formatblock', 'fontname', 'fontsize', '|', 'forecolor', 'hilitecolor', 'bold',
											'italic', 'underline', 'strikethrough', 'lineheight', 'removeformat', '|', 'image', 'multiimage',
											'flash', 'media', 'insertfile', 'table', 'hr', 'emoticons', 'baidumap', 'pagebreak',
											'anchor', 'link', 'unlink'
										],
						htmlTags		:{
								font : ['color', 'size', 'face', '.background-color'],
								span : ['style'],
								div : ['class', 'align', 'style'],
								'table,td,th': ['class', 'border', 'cellspacing', 'cellpadding', 'width', 'height', 'align', 'style', 'colspan', 'rowspan', 'bgcolor', 'style'],
								a : ['class', 'href', 'target', 'name', 'style'],
								embed : ['src', 'width', 'height', 'type', 'loop', 'autostart', 'quality',
								'style', 'align', 'allowscriptaccess', '/'],
								img : ['src', 'width', 'height', 'border', 'alt', 'title', 'align', 'style', '/'],
								hr : ['class', '/'],
								br : ['/'],
								p:['class', 'style','width','height','align'],
								'ol,ul,li,blockquote,h1,h2,h3,h4,h5,h6,pre,tbody,tr,strong,b,sub,sup,em,i,u,strike,iframe': []
						},
						afterCreate		: function() {
							this.sync();
						},
						afterBlur		:function(){
							this.sync();
						}
					})
				}
				catch (e)
				{
					alert("编辑器创建失败！");	
				}
				//$(".kinediter").each(function(){
					//KE.init({
							//id : this.id,
							//imageUploadJson : '../../../admin/upload_json.asp',
							//autoSetDataMode:false,
							//shadowMode : false,
							//allowPreviewEmoticons : false
							////fileManagerJson : '../../../admin/file_manager_json.asp',
							////allowFileManager : true
					//});
					//KE.create(this.id);
				//});
			}
			if($(".xheditor").length>0){
				$('.xheditor').xheditor({upLinkUrl:"Upload.asp?Immediate=1",upLinkExt:"zip,rar,txt",upImgUrl:"Upload.asp?Immediate=1",upImgExt:"jpg,jpeg,gif,png",upFlashUrl:"Upload.asp?Immediate=1",upFlashExt:"swf",upMediaUrl:"Upload.asp?Immediate=1",upMediaExt:"avi"});
			}
			if($(".ueditor").length>0){
				
				$(".ueditor").each(function(){
					var initID;
					initID=($(this).attr("id"));
					console.info(initID);
					var ue = UE.getEditor(initID);
					//对编辑器的操作最好在编辑器ready之后再做
					ue.ready(function() {
					//设置编辑器的内容
					ue.setContent('hello');
					//获取html内容，返回: <p>hello</p>
					var html = ue.getContent();
					//获取纯文本内容，返回: hello
					var txt = ue.getContentTxt();
					});
					ue.sync();
				})
			}
		
			/* if($("#Fk_Article_Content").length>0){
				$('#Fk_Article_Content').xheditor({upLinkUrl:"Upload.asp?Immediate=1",upLinkExt:"zip,rar,txt",upImgUrl:"Upload.asp?Immediate=1",upImgExt:"jpg,jpeg,gif,png",upFlashUrl:"Upload.asp?Immediate=1",upFlashExt:"swf",upMediaUrl:"Upload.asp?Immediate=1",upMediaExt:"avi"});
			}
			if($("#Fk_Module_Content").length>0){
				$('#Fk_Module_Content').xheditor({upLinkUrl:"Upload.asp?Immediate=1",upLinkExt:"zip,rar,txt",upImgUrl:"Upload.asp?Immediate=1",upImgExt:"jpg,jpeg,gif,png",upFlashUrl:"Upload.asp?Immediate=1",upFlashExt:"swf",upMediaUrl:"Upload.asp?Immediate=1",upMediaExt:"avi"});
			}
			if($("#Fk_Product_Content").length>0){
				$('#Fk_Product_Content').xheditor({upLinkUrl:"Upload.asp?Immediate=1",upLinkExt:"zip,rar,txt",upImgUrl:"Upload.asp?Immediate=1",upImgExt:"jpg,jpeg,gif,png",upFlashUrl:"Upload.asp?Immediate=1",upFlashExt:"swf",upMediaUrl:"Upload.asp?Immediate=1",upMediaExt:"avi"});
			}
			if($("#Fk_Down_Content").length>0){
				$('#Fk_Down_Content').xheditor({upLinkUrl:"Upload.asp?Immediate=1",upLinkExt:"zip,rar,txt",upImgUrl:"Upload.asp?Immediate=1",upImgExt:"jpg,jpeg,gif,png",upFlashUrl:"Upload.asp?Immediate=1",upFlashExt:"swf",upMediaUrl:"Upload.asp?Immediate=1",upMediaExt:"avi"});
			}
			if($("#Fk_Info_Content").length>0){
				$('#Fk_Info_Content').xheditor({upLinkUrl:"Upload.asp?Immediate=1",upLinkExt:"zip,rar,txt",upImgUrl:"Upload.asp?Immediate=1",upImgExt:"jpg,jpeg,gif,png",upFlashUrl:"Upload.asp?Immediate=1",upFlashExt:"swf",upMediaUrl:"Upload.asp?Immediate=1",upMediaExt:"avi"});
			} */
			if($("#DelWord").length>0){
				$('#DelWord').text(unescape($('#DelWord').val()));
			}
			if($("#KeyWord").length>0){
				$('#KeyWord').text(unescape($('#KeyWord').val()));
			}
			PageReSize();
			$("#AlphaBox").height($(document).height());
			return true;
		}
	);
	$('select').hide();
}

//==========================================
//函数名：GetCheckbox
//用途：获取选中的checkbox
//参数?
//==========================================
function GetCheckbox(){
	var text="";   
	$("input[class=Checks]").each(function() {   
		if ($(this).attr("checked")) {
			if(text==''){
				text = $(this).val();   
			}else{
				text += ","+$(this).val();   
			}
		}   
	}); 
	return text;
}

//==========================================
//函数名：
//用途：按ESC关闭弹出窗口
//参数?
//==========================================
document.onkeydown=function(){ 
	if(window.event.keyCode==27){ 
		$("#Boxs").hide();
		$('select').show();
	} 
} 

//==========================================
//函数名：PageReSize
//用途：页面初&#59216;?
//参数?
//==========================================
function PageReSize(){
	var LeftWidth=189;
	var RightWidth=$("#Bodys").width()-189;
	var WindowsHeight=$(document).height()-90;
	var LeftHeight=$("#MainLeft").height();
	var RightHeight=$("#MainRight").height();
	if(RightWidth>812){
		$("#MainRight").width(RightWidth);
		$("#AllBox").width($("#Bodys").width());
	}else{
		$("#AllBox").width(1001);
	}
	if(LeftHeight<WindowsHeight||LeftHeight<RightHeight){
		if(RightHeight>WindowsHeight){
			$("#MainLeft").height(RightHeight);
		}else{
			$("#MainLeft").height(WindowsHeight);
		}
	}
}

//==========================================
//函数名：SetRContent
//用途：替换DIV内容;
//参数：DivId：&#59110;替换的DIV
//     Urls：获取内容的链接
//==========================================
function SetRContent(DivId,Urls){
	document.getElementById(DivId).innerHTML="<a href='javascript:void(0);' title='点击关闭' onclick=$('select').show();$('#Boxs').hide()><img src='Images/Loading.gif' /></a>";
	$.get(Urls,
		function(data){
			document.getElementById(DivId).innerHTML=data;
			PageReSize();
			return true;
		}
	);
	PageReSize();
}

//==========================================
//函数名：GetPinyin
//用途：获取拼音
//参数：InPutId：放&#57790;Input
//     Urls：获取内容的链接
//==========================================
function GetPinyin(InPutId,Urls){
	$.get(Urls,
		function(data){
			document.getElementById(InPutId).value=data;
			return true;
		}
	);
}

//==========================================
//函数名：OpenMenu
//用途：菜单开?
//参数：MenuId：菜单ID
//==========================================
function OpenMenu(MenuId){
	if($("#"+MenuId).css("display")=="none"){
		$("#"+MenuId).css("display","block");
	}else{
		$("#"+MenuId).css("display","none");
	}
}

//==========================================
//函数名：DelIt
//用途：通用删除
//参数：Cstr：提示&#57826;?
//     Urls：执行URL
//     F5Url：刷新URL
//     F5Div：刷新DIV
//==========================================
function DelIt(Cstr,Urls,F5Div,F5Url){
	//询问框
	layer.confirm(Cstr, {
		title:'操作提示：',
		btn: ['确定','取消'] //按钮
	}, function(){
		// layer.msg('的确很重要', {icon: 1});
		$.get(Urls,
			function(data){
				//alert(data);
				//提示层
				layer.msg(data);
				var Arrstr1 = new Array();
				var Arrstr2 = new Array();
				Arrstr1 = F5Div.split("|");
				Arrstr2 = F5Url.split("|");
				for(var i=0;i<Arrstr1.length;i++){
					SetRContent(Arrstr1[i],Arrstr2[i]);
				}
				return true;
			}
		);
	}, function(){
		return;
	});
	return;
}

//==========================================
//函数名：SendGet
//用途：表单提交获取信息
//参数：FormName：提交的FORM
//     ToUrl：提交向的链?
//     F5Div：刷新DIV
//==========================================
function SendGet(FormName,ToUrl,F5Div){
    var options = { 
		url:  ToUrl,
        beforeSubmit:function(formData, jqForm, options){
          return true; 
		},
        success:function(responseText, statusText){
          if(statusText=="success"){
			  document.getElementById(F5Div).value=responseText;
          }
          else{
           // alert(statusText);
			layer.msg(statusText);
          }
		}
    }; 
    $('#'+FormName+'').ajaxForm(options); 
}

//==========================================
//函数名：Sends
//用途：表单提交
//参数：FormName：提交的FORM
//     ToUrl：提交向的链?
//     SuGo：成功后&#57882;&#57600;链接?&#57600;?不转?
//     GoUrl：转向链?
//     SuAlert：成功后&#57882;弹出框提示，1弹出?不弹?
//     SuF5：成功后&#57882;刷新DIV?刷新?不刷?
//     F5Url：刷新URL
//     F5Div：刷新DIV
//==========================================
function Sends(FormName,ToUrl,SuGo,GoUrl,SuAlert,SuF5,F5Div,F5Url){
	var oldval=$("#button").val();
    var options = { 
		url:  ToUrl,
        beforeSubmit:function(formData, jqForm, options){
		  $("#button").val("正在提交...");
		  $("#button").attr("disabled","disabled");
          return true; 
		},
        success:function(responseText, statusText){
          if(statusText=="success"){
            if(responseText.search("成功")>0){
				layer.closeAll('page'); //关闭所有页面层
				if(SuAlert==1){
					//alert(responseText);
					layer.msg(responseText);
				}
				else{
					$("#Boxs").hide();
					var st=responseText.replace(/\|\|\|\|\|/gi,"\n");
					//alert(st);
					layer.msg(st);
				}
				if(SuGo==1){
					location.href=GoUrl;
				}
				if(SuF5==1){
					var Arrstr1 = new Array();
					var Arrstr2 = new Array();
					Arrstr1 = F5Div.split("|");
					Arrstr2 = F5Url.split("|");
					for(var i=0;i<Arrstr1.length;i++){
						SetRContent(Arrstr1[i],Arrstr2[i]);
					}
				}
            }
            else{
				var st=responseText.replace(/\|\|\|\|\|/gi,"\n");
				//alert(st);
				layer.msg(st);
		 		 $("#button").val(oldval);
		 		 $("#button").removeAttr("disabled");
           }
          }
          else{
            //alert(statusText);
			layer.msg(statusText);
          }
		  $("#button").val(oldval);
		  $("#button").removeAttr("disabled");
		}
    }; 
    $('#'+FormName+'').ajaxForm(options); 
}

//==========================================
//函数名：Sends_Div
//用途：表单提交更新DIV
//参数：FormName：提交的FORM
//     ToUrl：提交向的链?
//     F5Div：刷新DIV
//==========================================
function Sends_Div(FormName,ToUrl,F5Div){
	document.getElementById(F5Div).innerHTML="<a href='javascript:void(0);' title='点&#57629;关闭' onclick=$('#Boxs').hide()><img src='Images/Loading.gif' /></a>";
    var options = { 
		url:  ToUrl,
        beforeSubmit:function(formData, jqForm, options){
          return true; 
		},
        success:function(responseText, statusText){
          if(statusText=="success"){
            document.getElementById(F5Div).innerHTML=responseText;
          }
          else{
            //alert(statusText);
			layer.msg(statusText);

          }
		}
    }; 
    $('#'+FormName+'').ajaxForm(options); 
}

//==========================================
//函数名：ChangeSelect
//用途：&#57789;Select内&#57744;
//参数：Urls：执行URL
//     SId：操作的Select
//==========================================
function ChangeSelect(Urls,SId){
	$.get(Urls,
		function(data){
			if(data!=""){
				BuildSel(data,document.getElementById(SId));
			}
			return true;
		}
	);
}

//==========================================
//函数名：BuildSel
//用途：执&#58641;&#57789;Select内&#57744;
//参数：Urls：执行URL
//     SId：操作的Select
//==========================================
function BuildSel(Str,Sel){
	//先清空原来的数据.
	var Arrstr = new Array();
	Arrstr = Str.split(",,,,,");
	//开始构建新的Select.
	if(Str!=""){
		Sel.options.length=0;
		var arrst;
		for(var i=0;i<Arrstr.length;i++){
			if(Arrstr[i]!=""){
				Arrst=Arrstr[i].split("|||||");
				Sel.options[Sel.options.length]=new Option(Arrst[1],Arrst[0]);
			}
		}
	}
}


//==========================================
//函数名：ModuleTypeChange
//用途：模块类型选择内&#57744;变换
//参数：ModuleTypeId：模板类?
//==========================================
function ModuleTypeChange(ModuleTypeId){
	if(ModuleTypeId==0){
		$("#Fk_Module_PageCounts").css("display","none");
		$("#Fk_Module_Keywords").css("display","block");
		$("#Fk_Module_Descriptions").css("display","block");
		$("#Fk_Module_Dirs").css("display","none");
		$("#Fk_Module_FileNames").css("display","block");
		$("#Fk_Module_Templates").css("display","block");
		$("#Fk_Module_LowTemplates").css("display","none");
		$("#Fk_Module_Urls").css("display","none");
		$("#Fk_Module_PageCodes").css("display","none");
	}
	if(ModuleTypeId==1){
		$("#Fk_Module_PageCounts").css("display","block");
		$("#Fk_Module_Keywords").css("display","block");
		$("#Fk_Module_Descriptions").css("display","block");
		$("#Fk_Module_Dirs").css("display","block");
		$("#Fk_Module_FileNames").css("display","none");
		$("#Fk_Module_Templates").css("display","block");
		$("#Fk_Module_LowTemplates").css("display","block");
		$("#Fk_Module_Urls").css("display","none");
		$("#Fk_Module_PageCode").val("第一页|--|上一页|--|下一页|--|尾页|--|条/页|--|共|--|页/|--|条|--|当前第|--|页|--|第|--|页");
		$("#Fk_Module_PageCodes").css("display","block");
		
	}
	if(ModuleTypeId==2){
		$("#Fk_Module_PageCounts").css("display","block");
		$("#Fk_Module_Keywords").css("display","block");
		$("#Fk_Module_Descriptions").css("display","block");
		$("#Fk_Module_Dirs").css("display","block");
		$("#Fk_Module_FileNames").css("display","none");
		$("#Fk_Module_Templates").css("display","block");
		$("#Fk_Module_LowTemplates").css("display","block");
		$("#Fk_Module_Urls").css("display","none");
		$("#Fk_Module_PageCode").val("第一页|--|上一页|--|下一页|--|尾页|--|条/页|--|共|--|页/|--|条|--|当前第|--|页|--|第|--|页");
		$("#Fk_Module_PageCodes").css("display","block");
	}
	if(ModuleTypeId==3){
		$("#Fk_Module_PageCounts").css("display","none");
		$("#Fk_Module_Keywords").css("display","block");
		$("#Fk_Module_Descriptions").css("display","block");
		$("#Fk_Module_Dirs").css("display","none");
		$("#Fk_Module_FileNames").css("display","block");
		$("#Fk_Module_Templates").css("display","block");
		$("#Fk_Module_LowTemplates").css("display","none");
		$("#Fk_Module_Urls").css("display","none");
		$("#Fk_Module_PageCodes").css("display","none");
	}
	if(ModuleTypeId==4){
		$("#Fk_Module_PageCounts").css("display","block");
		$("#Fk_Module_Keywords").css("display","block");
		$("#Fk_Module_Descriptions").css("display","block");
		$("#Fk_Module_Dirs").css("display","none");
		$("#Fk_Module_FileNames").css("display","block");
		$("#Fk_Module_Templates").css("display","block");
		$("#Fk_Module_LowTemplates").css("display","none");
		$("#Fk_Module_Urls").css("display","none");
		$("#Fk_Module_PageCode").val("第一页|--|上一页|--|下一页|--|尾页|--|条/页|--|共|--|页/|--|条|--|当前第|--|页|--|第|--|页");
		$("#Fk_Module_PageCodes").css("display","block");
	}
	if(ModuleTypeId==5){
		$("#Fk_Module_PageCounts").css("display","none");
		$("#Fk_Module_Keywords").css("display","none");
		$("#Fk_Module_Descriptions").css("display","none");
		$("#Fk_Module_Dirs").css("display","none");
		$("#Fk_Module_FileNames").css("display","none");
		$("#Fk_Module_Templates").css("display","none");
		$("#Fk_Module_LowTemplates").css("display","none");
		$("#Fk_Module_Urls").css("display","block");
		$("#Fk_Module_PageCodes").css("display","none");
	}
	if(ModuleTypeId==6){
		$("#Fk_Module_PageCounts").css("display","none");
		$("#Fk_Module_Keywords").css("display","none");
		$("#Fk_Module_Descriptions").css("display","none");
		$("#Fk_Module_Dirs").css("display","none");
		$("#Fk_Module_FileNames").css("display","none");
		$("#Fk_Module_Templates").css("display","block");
		$("#Fk_Module_LowTemplates").css("display","none");
		$("#Fk_Module_Urls").css("display","none");
		$("#Fk_Module_PageCodes").css("display","none");
	}
	if(ModuleTypeId==7){
		$("#Fk_Module_PageCounts").css("display","block");
		$("#Fk_Module_Keywords").css("display","block");
		$("#Fk_Module_Descriptions").css("display","block");
		$("#Fk_Module_Dirs").css("display","block");
		$("#Fk_Module_FileNames").css("display","none");
		$("#Fk_Module_Templates").css("display","block");
		$("#Fk_Module_LowTemplates").css("display","block");
		$("#Fk_Module_Urls").css("display","none");
		$("#Fk_Module_PageCode").val("第一页|--|上一页|--|下一页|--|尾页|--|条/页|--|共|--|页/|--|条|--|当前第|--|页|--|第|--|页");
		$("#Fk_Module_PageCodes").css("display","block");
	}
}

//==========================================
//函数名：ColorPicker
//用途：颜色选择
//参数：ColorInput：&#59110;颜色的Input
//==========================================
function ColorPicker(ColorInput) { 
	var sColor=dlgHelper.ChooseColorDlg();
	if(sColor.toString(16)==0){
		ColorInput.value=""; 
	}else{
		ColorInput.value="#"+sColor.toString(16); 
	}
} 

//==========================================
//函数名：CheckAll
//用途：全?
//参数：form：表?
//==========================================
function CheckAll(form) {
	for (var i = 0; i < form.elements.length; i++) {
		var e = form.elements[i];
		if (e.name != 'chkall') 
			e.checked = form.chkall.checked;
	}
}

function SelectAll(strName) {
 var checkboxs=document.getElementsByName(strName);
 for (var i=0;i<checkboxs.length;i++) {
  var e=checkboxs[i];
  e.checked=!e.checked;
 }
}


function autoid(auto)
{
var xml = new ActiveXObject("MSXML2.XMLHTTP");
	hrefValue = window.location.hostname; //获取当前页面的地址
    xml.open("get","/admin/GetData.asp?act="+auto,false);
xml.send();
alert(xml.responseText);
   if (xml.responseText == 0)
   {
      //alert ("获取失败，请重试！");
		layer.msg('获取失败或该账号未设置开通，请重试！');
   }
   else
   {
     document.getElementById(auto).value=xml.responseText;
      ymPrompt.alert({title:'自动获取成功！',message:'获取成功！记得别忘记保存设置哦！',width:300,height:180})
   }
}



//自定义  弹出信息窗口
 function tan(txt)
   {
    ymPrompt.alert({message:txt,title:'提示：',showMask:false});
   }
// 自定义弹窗并刷新父窗口
 function tan2(txt)
   {
    ymPrompt.alert({message:txt,title:'提示：',showMask:false,handler:reloadtoppage});
   }
 function reloadtoppage()
  {
    window.parent.location.reload();
  }

//弹出错误提示并返回上一页
 function tan3(txt)
   {
    ymPrompt.alert({message:txt,title:'提示：',width:400,height:220,showMask:false,handler:goback});
   }
 function goback()
 {
     history.back(-1);
 }

function mid(mainStr,starnum,endnum){ 
if (mainStr.length>=0){ 
return mainStr.substr(starnum,endnum) 
}else{return null} 
//mainStr.length 
} 
 
 //关键词查排名

 function chakeywordspaiming(keywords,i){
	$("#chaciarea"+i).html("<img src=/admin/shangwin/images/load.gif>");
	$.ajax({
		type:"GET",
		url:"/admin/shangwin/getPaiming.asp",
		data:"d="+checkistopdomain()+"&k="+encodeURIComponent(keywords)+"&r="+Math.random(),
		dataType:"html",
		cache: false,
		timeout: 90000,
		error: function(XMLHttpRequest, textStatus, errorThrown){
			$("#chaciarea"+i).html("查询失败！");
			//$("#chaciarea"+i).html(XMLHttpRequest.status);
		},
		success:function(msg){
			//alert(msg);
			$("#chaciarea"+i).html(msg);
			//alert(checkistopdomain()=="");
			if (checkistopdomain()!=""){
				savekeywordspaiming(keywords,msg);
				if(mid(msg,6,1)!=0) {
					updatekeyword(checkistopdomain(),keywords);
				}
			}
			//savekeywordspaiming(keywords,msg);
		}
	})
 }
 
function checkistopdomain()
    {
        var domainname = document.domain.toLowerCase();
//alert(domainname.length);
        if (domainname.indexOf("www.")>=0)
        {
            domainname = domainname.substr(domainname.indexOf("www.") + 4, domainname.length - domainname.indexOf("www.") - 4);
        }
        if (domainname.indexOf(".com.cn")>-1 || domainname.indexOf(".com.hk")>-1 || domainname.indexOf(".net.cn")>-1 || domainname.indexOf(".gov.cn")>-1 || domainname.indexOf(".org.cn ")>-1)
        {

            if (domainname.split('.').length > 3)
            {
                return "";
            }
            else
            {
                return domainname;
            }
        }
        else
        {
            if (domainname.split('.').length > 2)
            {
                return "";
            }
            else
            {
                return domainname;
            }
        }
    }


 function savekeywordspaiming(keywords,paimingjieguo){
	// var xml = new ActiveXObject("MSXML2.XMLHTTP");
	// var saveurl="/admin/saveseopaiming.asp?paimingkeywords="+encodeURIComponent(keywords)+"&paimingjieguo="+paimingjieguo
    // xml.open("get",saveurl,true);
	// xml.send();
	$.ajax({
			url: '/admin/saveseopaiming.asp',
			type: 'GET',
			dataType: 'html',
			cache: false,
			timeout: 10000,
			data: "paimingkeywords="+encodeURIComponent(keywords)+"&paimingjieguo="+paimingjieguo, 
			error: function(){
			//alert('载入出错，请刷新重试！');
			},
			success: function(html){
			//alert(html);
		    }
			});
 }
 
//win系统站群获取关键词排名
function GetRank(u,k,i){
	$.ajax({
		type:"GET",
		url:"/admin/shangwin/seo/paiming/getRank.asp",
		data:"u="+u+"&k="+k+"&r="+Math.random(),
		dataType:"html",
		timeout: 90000,
		error: function(){
			//alert('载入出错，请刷新重试！');
		},
		beforeSend:function(){
			$("#chaciarea"+i).html("<img src=/admin/shangwin/seo/paiming/images/load.gif>");
		},
		success:function(msg){
			$("#chaciarea"+i).html(msg);
		}
	})
}

//更新关键词
function updatekeyword(u,k){
	$.ajax({
		type:"GET",
		url:"/admin/shangwin/updatePVR.asp",
		data:"u="+u+"&k="+encodeURIComponent(k)+"&r="+Math.random(),
		dataType:"html",
		timeout: 90000,
		error: function(){
			//alert('载入出错，请刷新重试！');
		},
		success:function(msg){
			//alert(msg);
		}
	})
}

//获取有效果访问量
function GetVisits(tjid,keyword,i,n){
	$.ajax({
		type:"GET",
		url:"/admin/shangwin/xmlhttp.asp",
		data:"tjid="+tjid+"&k="+encodeURIComponent(keyword)+"&r="+Math.random(),
		dataType:"html",
		cache: false,
		timeout: 90000,
		beforeSend:function(){
			$("#chavisits"+i).html("<img src=/admin/shangwin/images/load.gif>");
		},
		error: function(){
			$("#chavisits"+i).html("查询失败！");
		},
		success:function(msg){
			$("#chavisits"+i).html(msg);
			$("#tvisits").html(0);
			for(var j=0;j<=n;j++){
				var cnum=parseInt($("#chavisits"+j).html());
				if (isNaN(cnum)){cnum=0};
				$("#tvisits").html(cnum+parseInt($("#tvisits").html()));
			}
		}
	})
}

//获取所有有效果访问量
function GetAllVisits(tjid){
	$("#Allvisits").html("<img src=/admin/shangwin/images/load.gif>");
	$.ajax({
		type:"GET",
		url:"/admin/shangwin/GetAllVisits.asp",
		data:"tjid="+tjid+"&r="+Math.random(),
		dataType:"html",
		cache: false,
		timeout: 90000,
		error: function(){
			$("#chavisits"+i).html("查询失败！");
		},
		success:function(msg){
			$("#Allvisits").html(msg);
		}
	})
}

//提取关键词
function tiqu(t,ContentID,TagsID){
	var tq="";
	var c=$("#"+ContentID).val().replace(/<.*?>/g,"");
	var k=$("#Fk_Keywordlist").val();
	if (t==0){
		var arrK=new Array();
		if (k.indexOf("|")>-1)
		{
			arrK=k.split("|");
			for(var i=0 ;i<arrK.length;i++){
				if (c.indexOf(arrK[i])>-1) {
					tq=tq + ","+ arrK[i] ;
				}
			}
			if(tq.substring(0,1)==","){
				tq=tq.substr(1)
			}
		}
		else{
			if (c.indexOf(k)>-1) {
				tq= k ;
			}
		}
	}
	else{
		tq=c.replace(/(^\s+)|(\s+$)/g,"");
		tq=tq.replace(/&nbsp;/ig,"");
		tq=tq.replace(/\t/g,"");
		tq=tq.replace(" ","");
		tq=tq.substring(0,99);
	}
	$("#"+TagsID).val(tq);
}


function showback(){
	$(".forminfo").css("display","block");
	$(".childType").css("display","none");
	$(".childType tbody .showback").parent().css("display","none");
	$("#alertdiv").height($(".forminfo").height()+61);
}

function showChild(index){
	$(".forminfo").css("display","none");
	$(".childType").css("display","block");
	$(".childType tbody").css("display","none");showback
	$(".childType tbody .showback").parent().css("display","block");
	$(".childType #child"+index).css("display","block");
	$("#alertdiv").height($(".childType").height()+61);
}

function alertinfo(id){
//显示弹出层
var obj = document.getElementById(id); 
var W = screen.width;//取得屏幕分辨率宽度 
var H = screen.height;//取得屏幕分辨率高度 
var yScroll;//取滚动条高度 
if (self.pageYOffset) { 
yScroll = self.pageYOffset; 
} else if (document.documentElement && document.documentElement.scrollTop){ 
yScroll = document.documentElement.scrollTop; 
} else if (document.body) {
yScroll = document.body.scrollTop; 
} 
//obj.style.marginLeft= (W/2 - 200) + "px";
obj.style.top= (H/2 -90 - 225　+　yScroll) + "px";
obj.style.display="block"; var scrollstyle = scrolls();
scrollstyle.style.overflowX = "hidden"; 
scrollstyle.style.overflowY = "hidden"; 
} 
 
function closediv(id){
	//关闭弹出层 
	document.getElementById(id).style.display="none"; 
	var scrollstyle = scrolls(); 
	scrollstyle.style.overflowY = "auto"; 
	scrollstyle.style.overflowX = "hidden";
	$("#ChooseType").removeAttr("disabled");
} 
 
function scrolls(){
//取浏览器类型 
var temp_h1 = document.body.clientHeight; 
var temp_h2 = document.documentElement.clientHeight; 
var isXhtml = (temp_h2<=temp_h1&&temp_h2!=0)?true:false; 
var htmlbody = isXhtml?document.documentElement:document.body;
return htmlbody; 
}

function getCookieVal (offset)
{
var endstr = document.cookie.indexOf (";", offset);
if (endstr == -1)
endstr = document.cookie.length;
return unescape(document.cookie.substring(offset, endstr));
}
function GetCookie (name)
{
var arg = name + "=";
var alen = arg.length;
var clen = document.cookie.length;
var i = 0;
while (i < clen)
{
var j = i + alen;
if (document.cookie.substring(i, j) == arg)
return getCookieVal (j);
i = document.cookie.indexOf(" ", i) + 1;
if (i == 0)
break;
}
return null;
}
