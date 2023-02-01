/*=================================================
程序名：JMenuTab(所谓的滑动门)
作者：xling
Blog:http://xling.blueidea.com
日期：2007/05/23

2007/05/25
把24日加入的自定义事件：onTabChange完善了一下，加入两个参数。
onTabChange(oldTab,self.activedTab);
oldTab:上次点击的那个tab
newTab:本次点击的tab
tab有三个属性：
index:
label:
tabPage:那addTab方法中的第二个参数。

加入方法:setSkin(pSkinName)
pSkinName是CSS文件中的。
示例：
#JMenuTabBlue {
	font-family: Verdana, Arial, Helvetica, sans-serif;
	font-size: 12px;
	padding: 2px;
	border-right-width: 1px;
	border-bottom-width: 1px;
	border-left-width: 1px;
}
#JMenuTabBlue .oInnerline {
	background-color: #FFFFFF;
}
...
...
要想使用这个skin,要先引用这个css文件，然后：
setSkin("JMenuTabBlue");
具体见示例文件：Demo2.htm
===================================================*/
function JMenuTab(pWidth,pHeight,pBody){
	var self = this;
	
	//________________________________________________
	var width = pWidth || "99%";
	var height = pHeight;
	
	this.titleHeight = 24;
	//________________________________________________
	var oOutline = null;
	var oTitleOutline = null;
	var oPageOutline = null;
	var oTitleArea = null;
	var oPageArea = null;
	
	var tabArray = new Array();
	this.activedTab = null;
	//________________________________________________
	this.onTabChange = new Function();
	//________________________________________________
	
	var $ = function(pObjId){
		return document.getElementById(pObjId);	
	}
	
	//________________________________________________
	
	var body = $(pBody) || document.body;
	
	//________________________________________________
	
	var htmlObject = function(pTagName){
		return document.createElement(pTagName);
	}
	
	//________________________________________________
	
	var isRate = function(pRateString){
		if(!isNaN(pRateString)) return false;
		if(pRateString.substr(pRateString.length-1,1) != "%")
			return false;
		if(isNaN(pRateString.substring(0,pRateString.length - 1)))
			return false;
		return true;
	}	
	
	//________________________________________________
	
	var createOutline = function(){
		
		var width_ = isRate(width) ? width : (!isNaN(width) ? width + "px" : "100%");
		
		oOutline = htmlObject("DIV");
		body.appendChild(oOutline);
		oOutline.style.width = width_;
	}
	
	//________________________________________________
	
	/*这个方法是为了解决外观问题，比如：
	setClassId("JMenuTab");
	在CSS里就要这样写：
	#JMenuTab {...}
	#JMenuTab .oTitleHeight{...}
	*/
	this.setSkin = function(pSkin){
		oOutline.id = pSkin;
	}
	//________________________________________________	
	
	var createTitleOutline = function(){
		oTitleOutline = htmlObject("DIV");
		oOutline.appendChild(oTitleOutline);
		oTitleOutline.className = "oTitleOutline";
		
		var vTable = htmlObject("TABLE");
		oTitleOutline.appendChild(vTable);
		vTable.width = "100%";
		vTable.border = 0;
		vTable.cellSpacing = 0;
		vTable.cellPadding = 0;
		
		var vTBody = htmlObject("TBODY");
		vTable.appendChild(vTBody);
		
		var vTr1 = htmlObject("TR");
		vTBody.appendChild(vTr1);
		
		var vTdTopLeft = htmlObject("TD");
		vTr1.appendChild(vTdTopLeft);
		vTdTopLeft.height = self.titleHeight;
		vTdTopLeft.className = "oTopLeft";
		
		oTitleArea = htmlObject("TD");/////////////////////////////////
		vTr1.appendChild(oTitleArea);
		oTitleArea.className = "oTitleArea";
		
		var vTdTopRight = htmlObject("TD");
		vTr1.appendChild(vTdTopRight);
		vTdTopRight.className = "oTopRight";
	}
	
	//________________________________________________
	this.setTitleHeight = function(pHeight){
		//设置标题区域的高度
	}
	
	//________________________________________________
	
	var tabBtn_click = function(){
		self.setActiveTab(this.index);
	}
	
	var tabBtn_mouseover = function(){
		if(this.className =="oTabBtnActive")
			return;
		
		this.className = "oTabBtnHover";
	}
	
	var tabBtn_mouseout = function(){
		if(this.className =="oTabBtnActive")
			return;
		this.className = "oTabBtn";
	}	
	//________________________________________________
	
	var createTabBtn = function(pLabel,pTabPage){
		var vTabBtn = htmlObject("DIV");
		oTitleArea.appendChild(vTabBtn);
		vTabBtn.className = "oTabBtn";
		//////////////////////////////////
		vTabBtn.index = tabArray.length;
		vTabBtn.label = pLabel;
		vTabBtn.tabPage = pTabPage;
		//////////////////////////////////
		vTabBtn.onclick = tabBtn_click;
		vTabBtn.onmouseover = tabBtn_mouseover;
		vTabBtn.onmouseout = tabBtn_mouseout;
		
		tabArray.push(vTabBtn);
		
		var vTabBtnL = htmlObject("DIV");
		vTabBtn.appendChild(vTabBtnL);
		vTabBtnL.className = "oTabBtnLeft";
		
		vTabBtnC = htmlObject("DIV");
		vTabBtn.appendChild(vTabBtnC);
		vTabBtnC.className = "oTabBtnCenter";
		vTabBtnC.innerHTML = pLabel;
		
		vTabBtnR = htmlObject("DIV");
		vTabBtn.appendChild(vTabBtnR);
		vTabBtnR.className = "oTabBtnRight";
	}
	
	
	var createPageOutline = function(){
		oPageOutline = htmlObject("DIV");
		oOutline.appendChild(oPageOutline);
		oPageOutline.className = "oPageOutline";
		
		var vTable = htmlObject("TABLE");
		oPageOutline.appendChild(vTable);
		vTable.width = "100%";
		vTable.border = 0;
		vTable.cellSpacing = 0;
		vTable.cellPadding = 0;
		vTable.style.borderCollapse = "collapse";
		vTable.style.tableLayout="fixed";
		
		var vTBody = htmlObject("TBODY");
		vTable.appendChild(vTBody);
		
		var vTr1 = htmlObject("TR");
		vTBody.appendChild(vTr1);
		
		var vTdBottomLeft = htmlObject("TD");
		vTr1.appendChild(vTdBottomLeft);
		vTdBottomLeft.className = "oBottomLeft";
		vTdBottomLeft.rowSpan = 2;
		
		oPageArea = htmlObject("TD");///////////////////////////////////////
		vTr1.appendChild(oPageArea);
		oPageArea.className = "oPageArea";
		if(oPageArea.filters)
			oPageArea.style.cssText = "FILTER: progid:DXImageTransform.Microsoft.Wipe(GradientSize=1.0,wipeStyle=0, motion='forward');";
		oPageArea.height = 10;
		
		var vTdBottomRight = htmlObject("TD");
		vTr1.appendChild(vTdBottomRight);
		vTdBottomRight.className = "oBottomRight";
		vTdBottomRight.rowSpan = 2;
		
		var vTr2 = htmlObject("TR");
		vTBody.appendChild(vTr2);
		
		var vTdBottomCenter = htmlObject("TD");
		vTr2.appendChild(vTdBottomCenter);
		vTdBottomCenter.className = "oBottomCenter";
	}
	
	//________________________________________________
	
	this.addTab = function (pLabel,pPageBodyId){
		createTabBtn(pLabel,pPageBodyId);
		if($(pPageBodyId)){
			oPageArea.appendChild($(pPageBodyId));
			$(pPageBodyId).style.display = "none";
		}
	}
		
	//________________________________________________
	
	this.setActiveTab = function(pIndex){
		if(oPageArea.filters)
			oPageArea.filters[0].apply();
		
		if(self.activedTab != null){
			self.activedTab.className = "oTabBtn";
			if($(self.activedTab.tabPage))
				$(self.activedTab.tabPage).style.display = "none";
		}
		
		var oldTab = self.activedTab;
		self.activedTab = tabArray[pIndex];
		self.onTabChange(oldTab,self.activedTab);//自定义事件,两个参数分别是先前的活动页签和现在活动的页签的index。
		self.activedTab.className = "oTabBtnActive";
		if($(self.activedTab.tabPage))
			$(self.activedTab.tabPage).style.display = "";
		
		if(oPageArea.filters)
			oPageArea.filters[0].play(duration=1);
	};
	
	//________________________________________________
	
	
	this.create = function(){
		createOutline();
		createTitleOutline();
		createPageOutline();
	}
}