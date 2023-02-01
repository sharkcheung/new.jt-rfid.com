$(function(){
	var mySwiper = new Swiper('.menu .swiper-container',{
		pagination: '.menu .propagination',
		paginationClickable: true,
		slidesPerView:4,
		calculateHeight:true,
	})  
	var mySwiper = new Swiper('.banner .swiper-container',{
			pagination: '.banner .pagination',
			loop:true,
			grabCursor: true,
			paginationClickable: true,
			autoplayDisableOnInteraction:false,
			calculateHeight:true,
			autoplay:2000,
		  })
})

