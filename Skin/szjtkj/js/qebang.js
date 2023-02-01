//menu js
$(document).ready(function(e) {
	var menuli = $(".nav > ul > li");
	var liwidth = menuli.width();
	var now = $(".nav ul a.hover").parent("li").index();	
	var distance=0;
    menuli.hover(function(){
    	var index=$(this).index();
    	console.log(index);
    	menumove(index);
    	$(this).children("ul").stop(true,true).delay(200).slideDown("fast");
	  	$(this).children("a").addClass("over");
			$(this).children("ul").mouseenter(function(){
	    	$(this).parent("li").children("a").addClass("hover");
			});
    }, function(){
    	menumove();
        $(this).children("ul").stop(true,true).slideUp("fast");
        $(this).children("a").removeClass("over");
        $(this).children("ul").mouseleave(function(){
			    $(this).parent("li").children("a").removeClass("hover");
			  });
    })
    function menumove(index){
    	var index = arguments.length > 0 ? index : now;
    	$(".menulihover").css("left",index*liwidth + distance);   
    	
    }
   menumove();
   $(".menulihover").css("display","block");
})


$(function(){
	$(".menu").on('click',function(){
    	$(this).find(".menu_icon").toggleClass("on");    	
    	if($(this).find(".menu_icon").hasClass("on")){
    		$("header nav").addClass('active');
    	}else{
    		$("header nav").removeClass('active');
    	}
  })
	
//	$(".nav > ul > li > ul").mouseenter(function(){
//  $(this).parent("li").children("a").addClass("hover");
//});
//$(".nav > ul > li > ul").mouseleave(function(){
//  $(this).parent("li").children("a").removeClass("hover");
//});
	
})

$(function(){
	$(".search_icon").click(function(){
		$('.search_form').toggle();
	})	
})


$(".backtop").click(function(){
	$('html,body').animate({scrollTop:0},700)
})
$(window).scroll(function(){
	var scrolTop=$(window).scrollTop();
	if(scrolTop < 500){
		$(".backtop").css("height","0");
	}else{ 
		$(".backtop").css("height","58px");
	}
})
$(window).scroll();
