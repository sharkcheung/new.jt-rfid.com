
//不知道谁写的做什么的。。。。。。。

function showumvcon(n)
{
    for(i=1;i<=2;i++)
    {
	document.getElementById("umvcon" + i).style.display='none';
	document.getElementById("umvtitle" + i).src="/images/button" + i + ".gif";
    }
    document.getElementById("umvcon" + n).style.display='block';
    document.getElementById("umvtitle" + n).src="/images/_button" + n + ".gif";
}

function showzgx(n_n)
{
    for(i_i=1;i_i<=3;i_i++)
    {
	document.getElementById("zgx" + i_i).style.display='none';
	document.getElementById("zgxm" + i_i).src="images/top" + i_i + ".jpg";
    }
    document.getElementById("zgx" + n_n).style.display='block';
    document.getElementById("zgxm" + n_n).src="images/_top" + n_n + ".jpg";
}
	
function g(o){
    return document.getElementById(o);
}
function HoverLi(n,m,q,p){
    for(var i=1;i<=m;i++)
    {
	g(q +i).className='normaltab';
	g(p+i).className='undis';
    }
    g(p+n).className='dis';
    g(q+n).className='hovertab';
}


function g2(o2){
    return document.getElementById(o2);
}
function HoverLi2(n2,m2,q2,p2){
    for(var i2=1;i2<=m2;i2++)
    {
	g2(q2+i2).className='normaltab2';
	g2(p2+i2).className='undis2';
    }
    g2(p2+n2).className='dis2';
    g2(q2+n2).className='hovertab2';
}	




//排行榜
$(function()
{
    $('div.top10_ul dl').mouseenter(function()
    {
	//关闭以前的
	$(this).parent().find('dd').hide();
	$(this).parent().find('dt').show();
	//打开当前的
	var dd = $(this).find('dd').show();
	var dt = $(this).find('dt');
	if (dd.height() > dt.height())
	{
	    $(this).find('dt').hide();
	}
    });
	
    //auction
//    $.get('/ajax/auction.php?act=current_price', {
//	auctionid : 1
//    }, function(data) {
//	$('#aution_block p.prices_block ').html('￥'+data);
//    });

    get_live_status();//2010-4-30 lll 直播预告
	
});


//处理公告滚动
$(function()
{
    var announceAllowAutoSlide = true;
    setInterval(function()
    {
	if (announceAllowAutoSlide) announceSlideDown();
    },3000);
	
    //鼠标进入公告区域后停止自动滚动
    $('.shell').mouseenter(function()
    {
	announceAllowAutoSlide = false;
    }).mouseleave(function()
    {
	announceAllowAutoSlide = true;
    });
	
    //箭头效果
    $('#announce_down').mouseover(function()
    {
	var img = $(this).find('img');
	img.data('old_src',img.attr('src'));
	img.attr('src','/images/yellow2.jpg');
    }).mouseout(function()
    {
	var img = $(this).find('img');
	img.attr('src',img.data('old_src'));
    }).click(announceSlideDown);
    $('#announce_up').mouseover(function()
    {
	var img = $(this).find('img');
	img.data('old_src',img.attr('src'));
	img.attr('src','/images/yellow.jpg');
    }).mouseout(function()
    {
	var img = $(this).find('img');
	img.attr('src',img.data('old_src'));
    }).click(announceSlideUp);
});


//公告向下滚动
function announceSlideDown()
{
    $('#aunnounce_ul').stop().animate({
	top:'-18px'
    },500,function()
{
	var li = $('#aunnounce_ul li:first-child');
	if (li.length > 0)
	{
	    $(li.get(0).cloneNode(1)).appendTo('#aunnounce_ul');
	    li.remove();
	    $('#aunnounce_ul').css('top','0');
	}
    });
}

//公告向上滚动
function announceSlideUp()
{
    var li = $('#aunnounce_ul li:last-child');
    $('#aunnounce_ul').stop().css('top','-18px').prepend(li.get(0).cloneNode(1));
    li.remove();
    $('#aunnounce_ul').animate({
	top:'0px'
    },500);
}




//开始slideshow


$(function()
{
    if(typeof(slidedata) =='undefined') return;
    var zindex = 60000;
    for(i=0;i<slidedata.length;i++)
    {
	$('<img src="'+slidedata[i].img+'" />').css({
	    zIndex:--zindex,
	    left:i*620
	    }).attr('index',i).appendTo('#slideshow_photo');
	$('<div class="slideshow-bt" index="'+i+'">'
	    +'<div class="slideshow-img"><img src="'+slidedata[i].btimg+'" /></div>'
	    +'<table class="slideshow-words"><tr><td>'
	    +slidedata[i].txt
	    +'</td></tr></table></div>').appendTo('#slideshow_footbar');
    }
    $('#slideshow_footbar .slideshow-bt').eq(0).addClass('bt-on');
    $('#slideshow_footbar .slideshow-bt').mouseenter(function(e)
    {
	slideTo(this);
    });
	
    $('#slideshow_photo img').add('#slideshow_footbar .slideshow-bt').click(function()
    {
	var index = parseInt($(this).attr('index'));
	var data = slidedata[index];
	if (data.target == '_blank')
	{
	    window.open(data.url);
	}
	else
	{
	    window.location = data.url;
	}
    });
	
    var indexAllowAutoSlide = true;
    $('#slideshow_wrapper').mouseenter(function()
    {
	indexAllowAutoSlide = false;
    }).mouseleave(function()
    {
	indexAllowAutoSlide = true;
    });
	
    setInterval(function()
    {
	if (indexAllowAutoSlide) slideDown();
    },3000);
	
});

function slideDown()
{
    if ($('#slideshow_footbar .slideshow-bt.bt-on').length <= 0) return;
    var nxt = $('#slideshow_footbar .slideshow-bt.bt-on').get(0).nextSibling;
    if (nxt == null)
    {
	nxt = $('#slideshow_footbar .slideshow-bt').get(0);
    }
    slideTo(nxt);
}

function slideTo(o)
{
    var x = $(o).get(0).offsetLeft;
    var imgx = -1*parseInt($('#slideshow_photo img[index='+$(o).attr('index')+']').css('left'));
    $('#slideshow_photo').stop().animate({
	left:imgx
    },500);
    $('#slideshow_footbar .bt-on').css({
	color:'#aaaaaa'
    }).removeClass('bt-on');
    $(o).css({
	color:'#ffffff'
    }).addClass('bt-on');
    $('#slide_mask').stop().animate({
	left:x
    },500);
}
$(function()
{
	var order =$('#djq').html();
	$.post("/ajax/get_topic_num.php?order="+order,{ }, function(data){
		var numa = data.up;
		var numb = data.down;
		$('#topicnuma').html(numa);
		$('#topicnumb').html(numb);
	},'json');
	
	
});
/*llll 2010-7-7
//聚焦彩条
$(function()
{
	var order =$('#djq').html();
	$.getJSON("/ajax/get_topic_num.php?order="+order,{ }, function(data){
		var numa = data.up;
		var numb = data.down;
		if(numa > numb){
			var pp=numb/numa;
			var width = Math.floor(pp  * 225);
			$('.line_color_green_list_b').animate({width:width+'px'},500);
		}else{
			var pp=numa/numb;
			var width = Math.floor(pp  * 225);
			$('.line_color_yellow_list_a').animate({width:width+'px'},500);
		}
		$('#numa').html(numa);
		$('#numb').html(numb);
	});
});
*/


//投票
$(function()
{
    //彩条上色
    $('[voteid]').each(function()
    {
	var t = 0;
	$(this).find('[barid]').each(function()
	{
	    t++;
	    $(this).addClass('color'+t);
	});
    });
    //调整箭头高度
    var arrow_top = Math.ceil($('.vote_wrapper_wrapper').height()/2) - 10;
    $('a.vote_left_arrow').css('top',arrow_top);
    $('a.vote_right_arrow').css('top',arrow_top);
	
    updateBar();
	
    $('.vote_block').eq(-1).attr('lastvote','yes');
    checkArrowAvail($('.vote_block').eq(0).attr('votecurrent','yes').attr('firstvote','yes'));
    $('a.vote_right_arrow').click(function()
    {
	var v = $('[votecurrent=yes]').next('.vote_block');
	if (!v || v.length <= 0) return;
	$('[votecurrent=yes]').removeAttr('votecurrent');
	v.attr('votecurrent','yes');
	checkArrowAvail(v);
	var x = -1*parseInt(v.get(0).offsetLeft);
	$('#vote_wrapper').stop().animate({
	    left:x
	},500,function()
{
			
	    });
    }).hover(function()
    {
	$(this).data('hover',true);
	if (!$(this).data('disabled'))
	    $(this).css('background-image','url(/images/vote_arrow_right_hover.gif)');
    },function()
    {
	$(this).data('hover',false);
	if (!$(this).data('disabled'))
	    $(this).css('background-image','url(/images/vote_arrow_right.gif)');
    });
	
    $('a.vote_left_arrow').click(function()
    {
	var v = $('[votecurrent=yes]').prev('.vote_block');
	if (!v || v.length <= 0) return;
	$('[votecurrent=yes]').removeAttr('votecurrent');
	v.attr('votecurrent','yes');
	checkArrowAvail(v);
	var x = -1*parseInt(v.get(0).offsetLeft);
	$('#vote_wrapper').stop().animate({
	    left:x
	},500,function()
{
	    });
    }).hover(function()
    {
	$(this).data('hover',true);
	if (!$(this).data('disabled'))
	    $(this).css('background-image','url(/images/vote_arrow_left_hover.gif)');
    },function()
    {
	$(this).data('hover',false);
	if (!$(this).data('disabled'))
	    $(this).css('background-image','url(/images/vote_arrow_left.gif)');
    });
});


function updateBar()
{
    updateBarBefor(); //2010-6-4

    $('[voteid]').each(function()
    {
	var total = 0;
	$(this).find('[optionid]').each(function()
	{
	    var vid = parseInt($(this).attr('optionid'));
	    var v = parseInt($(this).html());
	    total+= v;
	});
	if (total == 0) total = 1;
	$(this).find('[optionid]').each(function()
	{
	    var vid = parseInt($(this).attr('optionid'));
	    var v = parseInt($(this).html());
	    var percent = Math.floor(v*100/total);
	    $('[barid='+vid+']').animate({
		width:percent+'%'
	    },500);
	});
    });
}

function updateBarBefor()
{
    $('[voteid]').each(function()
    {
	var _voteid = $(this).attr('voteid');
	$.post("/ajax/home_page.php?action=color_bar&voteid=" + _voteid , {
	    content:"n"
	},
	function(rt){
	    if(not_empty(rt)) {
	    var arr = rt.split(",");
	    for(i=0;i<arr.length;i++)	$("[optionid="+_voteid+"000"+i+"]").html(arr[i]);
	    }
	}
	);
    });
}

function checkArrowAvail(v)
{
    if (!$('a.vote_left_arrow').data('hover')) $('a.vote_left_arrow').css('background-image','url(/images/vote_arrow_left.gif)');
    if (!$('a.vote_right_arrow').data('hover')) $('a.vote_right_arrow').css('background-image','url(/images/vote_arrow_right.gif)');
    $('a.vote_left_arrow').data('disabled',false);
    $('a.vote_right_arrow').data('disabled',false);
    if (v.attr('lastvote') == 'yes')
    {
	$('a.vote_right_arrow').data('disabled',true).css('background-image','url(/images/vote_arrow_right_no.gif)');
    }
	
    if (v.attr('firstvote') == 'yes')
    {
	$('a.vote_left_arrow').data('disabled',true).css('background-image','url(/images/vote_arrow_left_no.gif)');
    }
}


/*2010-7-8
function vote(id, type)
{
    //测试用
    //s = Math.floor(Math.random()*1000);
    //测试用end
    if(get_vote_status("survey", type) == false){
	alert("感谢您的参与！投票已成功，请勿重复投票！");
	return false;
    }

    //do php request
    $.post("/ajax/rating.php?vote_commont=y&up_down=1", {
	id:id,
	type:11
    },
    function(data){
	if(data ==0 ){
	    alert("投票失败！");
	}else if(data ==1 ){
	    //成功不用提示 alert("投票成功！");
	    $('[optionid='+id+']').html(parseInt($('[optionid='+id+']').html())+1);
	}else if(data ==2 ){
	    alert("感谢您的参与！投票已成功，请勿重复投票！");
	}else{
	    alert(data);
	}
	updateBar();
    }
    );
	
//$.post('/xxx.php',{ optionid: optionid },function(s)
//{
//	if (parseInt(s) > 0)
//	{
//		$('[optionid='+id+']').html(s);
//	}
//	else alert(s);
//});
}
*/
function vote(id, type)
{
	common_vote("survey", type, "感谢您的参与！投票已成功，请勿重复投票！", "index", 1, id, "survey");
}
function common_vote_callback(channel,ud,id,item_type){
	$('[optionid='+id+']').html(parseInt($('[optionid='+id+']').html())+1);
	updateBar();
}

//lll
function get_live_status(){ /* 首页直播设置 2010-4-12 */
    get_live_status_do();
    interval = window.setInterval(function(){
	get_live_status_do();
    }, 60000);
//window.clearInterval(interval);	
}
function get_live_status_do(){
    $.post("/ajax/home_page.php?action=index_live", {
	key:'v'
    },
    function(data){	//alert(data);
	if(data !="" ){
	    var pos=data.indexOf(':');
	    var program_id=data.substring(0,pos);
	    var home_living_img='/'+data.substr(pos+1); //alert(home_living_img);

	    $('#jiemu_yugao_0').hide();//隐藏第一个;
	    $('#jiemu_yugao_3').hide();
	    $('#jiemu_yugao_4').hide();
	    $('#index_live_pic').show(); //显示直播div
	    $('#index_live_url').attr('href','/live/'+program_id);//直播链接
	    $('#index_live_img_src').attr('src',home_living_img);//图片
	}else{
	    $('#jiemu_yugao_0').show();
	    $('#jiemu_yugao_3').show();
	    $('#jiemu_yugao_4').show();
	    $('#index_live_pic').hide();
	}
    });
}


$(function()
{
    $('.dd_con').mouseenter(function()
    {
	var a = $(this).find('a');
	var gap = a.width() - $(this).width();
	if (gap > 0)
	{
	    gap = -1*gap;
	    a.stop().animate({
		left: gap+'px'
	    },-1*gap*500/20,'linear');
	}
    }).mouseleave(function()
    {
	$(this).find('a').stop().animate({
	    left:'0px'
	},500);
    });
});

//处理广告高度
function deal_adv_height(){
    $('.advertisement_block').each(function(){
	if($(this).height() < 20) {
	    $(this).css("margin",'0 auto 0 auto');
	}
    });
    $('.imgs_gg_index2').each(function(){
	if($(this).height() < 20) {
	    $('.div_gg_120px_block_img').css("margin",'0 auto 0 auto').css('height','0px');
	}
    });
}
$(function()
{
	setTimeout("deal_adv_height()", 3000);
});

//-------------2010-7-15--------------------------------------------- 
function get_live_status(){ /* 首页直播设置 2010-4-12 */
	var d = new Date(); //alert(d.getHours());
	if(d.getHours() <12) return; // 到19点才计时

	get_live_status_do();
	interval = window.setInterval(function(){ get_live_status_do(); }, 60000);
	//window.clearInterval(interval);	
}
function get_live_status_do(){
	$.post("/ajax/home_page.php?action=index_live", { key:'v' }, 
	function(data){	//alert(data);
		if(data !="" ){
			var pos=data.indexOf('@');
			var program_id=data.substring(0,pos);
			var tmp=data.substr(pos+1);

			var ps=tmp.indexOf('@');
			var img_url=tmp.substring(0,ps); 
			var title=tmp.substr(ps+1);

			$('#lmhd_live').show(); //显示直播div
			$('#lmhd_default').hide(); //隐藏在路上首图
			//$('#live_title').text(title);
			$('#lmhd_live_url').attr('href','/live/'+program_id);//直播链接
			$('#lmhd_live_src').attr('src','/'+img_url);//图片链接
		}else{
			$('#lmhd_live').hide();
			$('#lmhd_default').show();
		}
	});
}