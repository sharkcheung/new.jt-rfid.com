
		function set_title_color(color) {
			$('#Fk_Product_Title').css('color',color);
			$('#Fk_Product_Color').val(color);
		}
		$(document).ready(function(){
				
				var editor = window.KindEditor.editor({
					fileManagerJson : '/admin/dkidtioenr/aps/file_manager_json.asp',
					uploadJson		: '/admin/dkidtioenr/aps/upload_json.asp',
					allowFileManager : true
				});
				var oldsummary=$("#Fk_Product_Summary").val();
				$("#GetAbstract").click(function(){
					if($(this).attr("checked")){
						$("#Fk_Product_Summary").val("");
						$("#Fk_Product_Summary").attr("disabled",true);
					}
					else{
						$("#Fk_Product_Summary").attr("disabled",false);
						$("#Fk_Product_Summary").val(oldsummary);
					}
					
				})
				$('#chooseImg').click(function() {
				
					editor.loadPlugin('image', function() {
						editor.plugin.imageDialog({
							imageUrl : $('#Fk_Product_Pic').val(),
							clickFn : function(url, title, width, height, border, align) {
								$('#Fk_Product_Pic').val(url);
								$('#SlImg').attr("src",url);
								editor.hideDialog();
							}
						});
					});
				});
				$('#MutiImg').click(function() {
					if($('#imgslist li').length>=10){
						var dialog = window.KindEditor.dialog({
						width : 200,
						title : '提示',
						body : '<div style="margin:10px;"><strong>一个产品下最多允许10张图片</strong></div>',
						closeBtn : {
							name : '关闭',
							click : function(e) {
								dialog.remove();
							}
						}
						});
					}
					else{
					editor.loadPlugin('multiimage', function() {
						editor.plugin.multiImageDialog({
							clickFn : function(urlList) {
								var div = $('#imgslist');
								$.each(urlList, function(i, data) {
									if($('#imgslist li').length<10){
										if(i==0 && $('#imgslist li.current').length==0){
											div.append('<li class=\"current\"><input type=\"hidden\" name=\"SlidesImg[]\" value=\"'+unescape(data.url)+'\" /><table border=\"0\" cellpadding=\"0\" cellspacing=\"0\"><tr><td class=\"imgdiv\"><a href=\"javascript:;\" target=\"_self\"><img src=\"'+unescape(data.url)+'\" /></a></td></tr></table><p><a onclick=\"deleteCurrentPic(this)\" href=\"javascript:;\" target=\"_self\"><b class=\"icon icon_del\" title=\"删除此图\"></b></a><a onclick=\"DiaWindowOpen_GetImgURL(this)\" title=\"'+unescape(data.url)+'\" href=\"javascript:;\" target=\"_self\"><b class=\"icon icon_imglink\" title=\"查看图片路径\"></b></a></p></li>');
										}
										else{
											div.append('<li><input type=\"hidden\" name=\"SlidesImg[]\" value=\"'+unescape(data.url)+'\" /><table border=\"0\" cellpadding=\"0\" cellspacing=\"0\"><tr><td class=\"imgdiv\"><a href=\"javascript:;\" target=\"_self\"><img src=\"'+unescape(data.url)+'\" /></a></td></tr></table><p><a onclick=\"deleteCurrentPic(this)\" href=\"javascript:;\" target=\"_self\"><b class=\"icon icon_del\" title=\"删除此图\"></b></a><a onclick=\"DiaWindowOpen_GetImgURL(this)\" title=\"'+unescape(data.url)+'\" href=\"javascript:;\" target=\"_self\"><b class=\"icon icon_imglink\" title=\"查看图片路径\"></b></a></p></li>');
										}
									}
								});
								if($("#bigimg td").html()==""){
									$("#bigimg td").html('<span class=\"imgwrap\"><input type=\"hidden\" name=\"SlidesImg[]FirstImg\" value=\"'+unescape((urlList[0].url))+'\" /><img src=\"'+unescape((urlList[0].url))+'\" /></span>');
								}
								editor.hideDialog();
							}
						});
					});
					}
				});
				$('#MutiImgChoose').click(function() {
					editor.loadPlugin('filemanager', function() {
						editor.plugin.filemanagerDialog({
							viewType : 'VIEW',
							dirName : 'image',
							clickFn : function(url, title) {
								K('#url').val(url);
								editor.hideDialog();
							}
						});
					});
				});
				$('#Fk_Product_Pic').blur(function(){
					if($('#Fk_Product_Pic').val()==""){
						$('#SlImg').attr("src","http://image001.dgcloud01.qebang.cn/website/ext/images/image.jpg");
					}
					else{
						$('#SlImg').attr("src",$('#Fk_Product_Pic').val());
					}
				})
				$('.imgdiv a:first-child').live('click', function() {
					$(this).parents("ul").children("li").removeClass("current");
					$(this).parents("li").addClass("current");
					$("#bigimg td").html('<span class=\"imgwrap\"><input type=\"hidden\" name=\"SlidesImg[]FirstImg\" value=\"'+unescape($(this).children("img").attr("src"))+'\" /><img src=\"'+unescape($(this).children("img").attr("src"))+'\" /></span>');
				});
		})
		
		//$:删除当前图片(如果当前图片为选中状态则在删除后设置第一张图片为默认选中)
function deleteCurrentPic(obj){
	var li=obj.parentNode.parentNode;
	var ul=li.parentNode;
	ul.removeChild(li);
	if(li.className=="current"){
		var td=document.getElementById("bigimg").getElementsByTagName("td");
		td[0].innerHTML="";
	}
	setGoodsFirstImg("imgslist");
}

//$:获取图片路径
function DiaWindowOpen_GetImgURL(obj){
	var imgURL=obj.title;
	var text="<div class=\"padding10\">"
	text=text+"<div class=\"inbox\"><table border=\"0\" cellpadding=\"0\" cellspacing=\"0\" width=\"100%\" class=\"infoTable\">";
	text=text+"<tr><td style=\"width:90px;text-align:right;border-bottom:0;\">图片路径<!--图片路径-->:</td><td style='border-bottom:0;'><input type=\"text\" class=\"TxtClass\" id=\"ImgURL\" name=\"ImgURL\" value=\""+imgURL+"\" /></td></tr>";
	text=text+"</table></div>";
	text=text+"</div>";
	var dialog = window.KindEditor.dialog({
						width : 500,
						title : '图片路径',
						body : text,
						closeBtn : {
							name : '关闭',
							click : function(e) {
								dialog.remove();
							}
						}
					});
	//var a=OpenWBS.CreateDiaWindow(600,200,"GetImgURL",Lang_Js_ImageURL,text,true);
}


//$:点击图片则作为默认选中
function selectFirstPic(parentNodeID,ElementName){
	var ul=document.getElementById(parentNodeID);
	var img=ul.getElementsByTagName("img");
	var pli;
	for(var i=0; i<img.length; i++){
		img[i].onclick=function(){
			pli=this.parentNode.parentNode.parentNode.parentNode.parentNode.parentNode;
			var tImg=ul.getElementsByTagName("li");
			for(var j=0; j<tImg.length; j++){
				tImg[j].className="";
			}
			pli.className="current";
			var td=document.getElementById("bigimg").getElementsByTagName("td");
			td[0].className="";
			td[0].innerHTML="<span class=\"imgwrap\"><input type=\"hidden\" name=\"SlidesImg[]FirstImg\" value=\""+this.src+"\" /><img src=\""+this.src+"\" /></span>";
		}
	}
}

//$:如果列表图片中没有选中的则设置第一张图片为默认选中
function setGoodsFirstImg(parentNodeID,ElementName){
	var lis=document.getElementById(parentNodeID).getElementsByTagName("li");
	var inputs=document.getElementById(parentNodeID).getElementsByTagName("input");
	if(inputs.length>0){
		var BigPic=document.getElementById("bigimg").getElementsByTagName("img");
		if(BigPic.length<1){
			var td=document.getElementById("bigimg").getElementsByTagName("td");
			lis[0].className="current";
			td[0].className="";
			td[0].innerHTML="<span class=\"imgwrap\"><input type=\"hidden\" name=\"SlidesImg[]FirstImg\" value=\""+inputs[0].value+"\" /><img src=\""+inputs[0].value+"\" /></span>";
		}
	}
}
//标签内容切换 href="javascript:TabSwitch('J_TabBar','TabBar_',1)"
function TabSwitch(TagsId,idpre,id){
	var divs,TabsMenu=document.getElementById(TagsId);
	var $li=TabsMenu.getElementsByTagName("li");
	for(var i=1;i<$li.length+1;i++){
		divs=document.getElementById(idpre+i);
		if(divs!=null){
			if(i==id){
				$li[i-1].className="current";divs.style.display="block";
			}else{
				$li[i-1].className="";divs.style.display="none";
			}
		}
	}
}
