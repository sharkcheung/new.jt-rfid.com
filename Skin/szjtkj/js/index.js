$(document).ready(function(e) {
	$('.banner.owl-carousel').owlCarousel({
	    loop:true,
	    margin:10,
	    items:1,
	    autoplay:true,
	    nav:false,
	    autoplayTimeout:3000,
    	autoplayHoverPause:true,
	});

	$('.product_list').owlCarousel({
	    loop:true,
	    nav:true,    
	    autoplay:true,
	    navText:['←','→'],
	    responsive:{
		    	0:{
	            items:1,
	        },	       	
	       	768:{
	            items:2,
	       	}
	    },    
	});
		
	$('.customer_list').owlCarousel({
    loop:true,
    nav:true,    
    autoplay:true,
    navText:['↑','↓'],
    responsive:{
	    	0:{
            items:1,
        },	       	
       	768:{
            items:2,
            margin:10,
       	},
       	1200:{
            items:3,
            margin:15,
       	}
    },    
	});
		
	$('.news_leftList').owlCarousel({
    loop:true,
    nav:false,    
    autoplay:true,
    margin:10,
    responsive:{
	    	0:{
            items:1,
        },	       	
       	520:{
            items:2,
       	},
       	1200:{
            items:3,
       	}
    },    
	});
		
	jQuery(".news dd").slide({mainCell:".news_rightList",autoPage:true,effect:"topLoop",autoPlay:true,vis:3});
})



$(function () {
    var list = new Swiper('.solution_bottom .center', {
        slidesPerView: 9,
        observer: true,//修改swiper自己或子元素时，自动初始化swiper
        observeParents: true,
        autoplayDisableOnInteraction: false,
        slideActiveClass : 'on',
    })    
    var lans = new Swiper('.solution_content', {
        paginationClickable: true,
        observer: true,//修改swiper自己或子元素时，自动初始化swiper
        observeParents: true,
        autoplayDisableOnInteraction: false,
        touchRatio: '0',
        autoplay: 3000,
        onSlideChangeEnd: function (swiper) {
          $('.solution .solution_bottom .swiper-slide').removeClass('on');
          $('.solution .solution_bottom .swiper-slide').eq(swiper.activeIndex).addClass('on');
        }      
    })   
    $('.solution .solution_bottom .swiper-slide').on('click', function () {
        var index = $(this).index();
        $('.solution .solution_bottom .swiper-slide').removeClass('on');
        $(this).addClass('on');
        lans.slideTo(index);   //slideTo  Swiper切换到指定slide
    })
})