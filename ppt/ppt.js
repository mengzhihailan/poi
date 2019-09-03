var pptSwiper = new Swiper ('.pptSwiper', {
	autoplay: true,
	loop: false,
	effect : 'fade',
	navigation: {
        nextEl: '.gallery-top-box .swiper-button-next',
        prevEl: '.gallery-top-box .swiper-button-prev',
    },
	pagination: {
		el: '.swiper-pagination',
	}
});

pptSwiper.on('slideChangeTransitionEnd', function () {
	var i = pptSwiper.activeIndex;
	$('.menu-img').removeClass('active');
	$('.menu-img').eq(i).addClass('active');
});

$('.menu-img').on('click', function () {
	pptSwiper.slideTo($(this).index(), 1000, true);
	$('.menu-img').removeClass('active');
	$(this).addClass('active');
	$('slideshow').addClass("pauseed");
	$("#play_icon").attr("src", "ppt/pause.png");
	pptSwiper.autoplay.stop();
	$(this).animate({scrollTop: 0}, 500)
})

$(".slideshow").click(function(){
	if($(this).hasClass("pauseed")){
		$(this).removeClass("pauseed");
		$("#play_icon").attr("src", "ppt/play.png");
		pptSwiper.autoplay.start();
	}else{
		$(this).addClass("pauseed");
		$("#play_icon").attr("src", "ppt/pause.png");
		pptSwiper.autoplay.stop();
	}
})