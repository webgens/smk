// 액션
$(window).scroll(function () {
	var currentScroll = $(document).scrollTop();
	var winHei = $(window).height();

	if (currentScroll > winHei) {
		$('.footer-wrap').addClass('point');
	} else if (currentScroll < winHei) {
		$('.footer-wrap').removeClass('point');
	}
});

// 브라우저 전체 영역 or 브라우저 전체 영역 - .header height
$(window).on('load resize', function () {
	var wHei = $(window).height();
	var wWid = $(window).width();

	$('.search-dim').css('height', wHei + 'px');
	$('.footer-wrap').css('height', wHei - 120 + 'px');
	$('.wrap-container').css('padding-bottom', wHei - 120 + 'px');
});

// .wrap-gnb <클릭 이벤트>
$('.wrap-gnb>ul>li>a').on('click', function () {
	var _this = $(this);

	/* lsh add */
	var menuName = _this.data("gnb");
	if (menuName == "brands") {
		get_GNB_BrandList('');
	}

	_this.closest('li').addClass('current').siblings().removeClass('current');
});

// 최근본 상품 및 장바구니 레이어 조작
var panelCtlr = function () {
	var recently = false,
		basket = false;

	$('.etc-menu-icon1').on('mouseenter', function () {
		$('.recently-viewed').show();
	});
	$('.etc-menu-icon1').on('mouseleave', function () {
		if (!recently) {
			$('.recently-viewed').hide();
			recently = false;
		}
	});
	$('.recently-viewed').on('mouseenter', function () {
		recently = true;
		$('.recently-viewed').show();
	});
	$('.recently-viewed').on('mouseleave', function () {
		$('.recently-viewed').hide();
		setTimeout(function () {
			recently = false;
		}, 1000)
	});

	$('.etc-menu-icon3').on('mouseenter', function () {
		get_GNB_CartList();
		$('.items-in-basket').show();
	});
	$('.etc-menu-icon3').on('mouseleave', function () {
		if (!basket) {
			$('.items-in-basket').hide();
			basket = false;
		}
	});
	$('.items-in-basket').on('mouseenter', function () {
		basket = true;
		get_GNB_CartList();
		$('.items-in-basket').show();
	});
	$('.items-in-basket').on('mouseleave', function () {
		$('.items-in-basket').hide();
		setTimeout(function () {
			basket = false;
		}, 1000)
	});
}();

$('.btn-hide').on('click', function () {
	// $('li.current>div').hide();
	// $('li.current>div').hide();
	$('li.current').removeClass('current');
});

// .search-ty1 <검색란 호출>
$('.etc-menu-icon4').on('click', function () {
	var $this = $(this);

	$this.closest('header').addClass('on');
	$this.closest('header').find('.search-dim').addClass('on')
	$this.closest('body').addClass('noscroll');
	if ($('.gnb-top-banner-wrap').css('display') === 'block') {
		$this.parent().next().addClass('pos');
		$this.parent().siblings('.search-panel').addClass('pos');
	} else {
		$this.parent().next().removeClass('pos');
		$this.parent().siblings('.search-panel').removeClass('pos');
	}

});
$('.search-ty1 .inn>button').on('click', function () {
	var $this = $(this);

	$this.closest('header').removeClass('on');
	$this.closest('header').find('.search-dim').removeClass('on');
	$this.closest('body').removeClass('noscroll');

	$this.parents('.search-ty1').removeClass('pos');
	$this.parents('.search-ty1').siblings('.search-panel').removeClass('pos');
});
$('.search-dim').on('click', function () {
	var $this = $(this);

	$this.closest('header').removeClass('on');
	$this.removeClass('on');
	$this.closest('body').removeClass('noscroll');

	$this.prev().find('.search-ty1').removeClass('pos');
	$this.prev().find('.search-panel').removeClass('pos');

});


// 스크롤바 Custom
/*
$('.scrollbar').slimScroll({
	width: '100%',
	height: '100%',
	size: '4px',
	opacity: '.4',
	borderRadius: '0',
	color: '#282828',
	alwaysVisible: false,
	disableFadeOut: false,
	railVisible: true,
	scrollTop: 0,
	railColor: '#dedede',
	railOpacity: 1,
	railBorderRadius: '0',
});
*/
// gnb sky banner
$('.gnb-top-banner-wrap .btn-banner-close').on('click', function () {
	$('.search-panel').animate({
		top: '120px'
	}, 0, function () {
		$('.search-ty1').animate({
			top: '0'
		}, 200);
	});
	$('.wrap-list .area-top123-list').animate({
		top: '120px'
	}, 300);
	$('.gnb-top-banner-wrap').animate({
		height: "toggle"
	}, 300, function () {
		$('.gnb-top-banner-wrap').remove();
	});
/*
	$('.wrap-list .area-top123-list').animate({
		top: '120px'
	}, 300);
	$('.gnb-top-banner-wrap').animate({
		height: "toggle"
	}, 300, function () {
		$('.gnb-top-banner-wrap').remove();
	});
	$('.search-ty1').animate({
		top: '0'
	}, 200);
	$('.search-panel').animate({
		top: '120px'
	}, 0);
*/
});

// 이벤트 배너 유무) .header fixed 0
$(window).on('load, scroll', function () {
	var scrollPos = $(document).scrollTop();
	var _topEventBanner = $('.gnb-top-banner-wrap');

	// 이벤트 베너가 없을 때
	if (_topEventBanner.length === 0) {
		$('.header').addClass('fixed');
		$('.wrap-container').addClass('fixed');


		// 이벤트 베너가 있을 때
	} else {
		if (scrollPos >= 70) {
			$('.header').stop().addClass('fixed');
			$('.wrap-container').addClass('fixed');
		} else if (scrollPos < 70) {
			$('.header').stop().removeClass('fixed');
			$('.wrap-container').removeClass('fixed');
		}
	}
});

//list 페이지 bg fixed top 이동
$(document).ready(function () {
	var _topEventBanner = $('.gnb-top-banner-wrap');

	if (_topEventBanner.length === 0) {
		$('.wrap-list .area-top123-list').css('top', '120px');
	} else {
		$('.wrap-list .area-top123-list').css('top', '190px');
	}
});

// Pop Up 호출 시 전체 스크롤 제거
if ($('.area-dim').css('display') === 'block') {
	$('body').addClass('noscroll');
}

//popup size
$(document).ready(function () {
	$('.area-pop').each(function () {
		var _this = $(this);
		if (_this.closest('wrap-pop').css('display') == 'block') {
			var _popHeight = _this.height();
			var _windowHeight = $(window).height();
			var _maxHeight = _windowHeight - 100;

			if (_popHeight > _maxHeight) {
				_this.css('height', _maxHeight);
			} else {
				_this.css('height', 'auto');
			}

			_this.closest('body').addClass('ofh');
		}
	});
});



/* GNB FILTER CLICK */
$(document).ready(function () {
	$('.alphabet-filter>ul>li').on('click', function () {
		$('.alphabet-filter>ul>li').find("a").removeClass('selected');
		$(this).find("a").addClass('selected');
		var prefix = $(this).find("a").data("prefix");
		get_GNB_BrandList(prefix);
	});
});
