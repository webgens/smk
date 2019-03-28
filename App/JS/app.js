window.cssTransition = function (target, duration) {
	return target.css({
		'-webkit-transition-duration': duration + 'ms',
		'-moz-transition-duration': duration + 'ms',
		'-o-transition-duration': duration + 'ms',
		'-ms-transition-duration': duration + 'ms',
		'transition-duration': duration + 'ms'
	});
};

if (!String.prototype.queryStringToObject) {
	(function () {
		return String.prototype.queryStringToObject = function () {
			var d, data, item, j, key, len, val, value;
			value = {};
			data = this.replace('?', '').split('&');
			for (j = 0, len = data.length; j < len; j++) {
				d = data[j];
				item = d.split('=');
				key = item[0];
				val = item[1];
				value[key] = val;
			}
			return value;
		};
	})();
};

// window.addEventListener('load', function() {
//     FastClick.attach(document.body);
// }, false);

$.fn.collapse = function (opts) {
	var _ = this,
		setting,
		module;

	options = $.extend(true, {}, collapseSetting, opts);
	setting = {
		firstOpen: options.firstOpen,
		accordion: options.accordion,
		useAnimate: options.useAnimate,
		tabMode: options.tabMode,
		duration: options.duration,
		item: _.find(options.selector.item),
		switch: _.find(options.selector.switch),
		toggler: _.find(options.selector.toggler),
		panel: _.find(options.selector.panel),
		closer: _.find(options.selector.closer),
		panelInner: _.find(options.selector.panelInner)
	}
	module = {
		init: function () {
			if (typeof (setting.firstOpen) == 'number') {
				module.open($(setting.panel[setting.firstOpen]), $(setting.panelInner[setting.firstOpen]).innerHeight(), 0);
				$(setting.item[0]).addClass('unfolded');
			}
			$(setting.toggler).on('click', module.event);
			$(setting.closer).on('click', function () {
				var
					$this = $(this),
					$item = $this.closest($(setting.item)),
					$panelInner = $item.find(setting.panelInner),
					i = $item.index(),
					h = $panelInner.innerHeight();

				module.close($(setting.panel), setting.duration);
				$item.removeClass('unfolded');
			});
		},
		event: function () {
			var
				$this = $(this),
				$item = $this.closest($(setting.item)),
				$panelInner = $item.find(setting.panelInner),
				i = $item.index(),
				h = $panelInner.innerHeight();

			if ($item.hasClass('unfolded')) {
				if (setting.tabMode) return;
				module.close($(setting.panel[i]), setting.duration);
				$item.removeClass('unfolded');
			} else {
				if (setting.accordion) {
					module.close($(setting.panel), setting.duration);
					$item.siblings().removeClass('unfolded');
				}
				module.open($(setting.panel[i]), h, setting.duration);
				$item.addClass('unfolded');
			}
		},
		open: function (target, h, d) {
			target.css('height', h);
			module.animate(target, d);
		},
		close: function (target, d) {
			target.css('height', 0);
			module.animate(target, d);
		},
		animate: function (target, d) {
			if (!setting.useAnimate) {
				cssTransition(target, 0);
			} else {
				cssTransition(target, d);
			}
		}
	}
	return module.init();
};
var collapseSetting = {
	firstOpen: 0, // number
	accordion: false, // boolean
	tabMode: false, // boolean
	useAnimate: true, // boolean
	duration: 500, // number
	selector: {
		item: '.collapse-item',
		switch: '.collapse-switch',
		toggler: '.collapse-toggler',
		closer: '.collapse-closer',
		panel: '.collapse-panel',
		panelInner: '.collapse-panel-inner'
	}
};
var FormCheckbox = {
	build: function (el) {
		if ($(el).find('input').is(':disabled')) {
			$(el).addClass('is-disabled');
		}
		if ($(el).find('input').prop('readonly')) {
			$(el).addClass('is-readonly');
		}
		if ($(el).find('input').is(':checked')) {
			$(el).addClass('is-checked')
		}
	},
	change: function (el) {
		if ($(el).is(':checked')) {
			$(el).parent().addClass('is-checked');
		} else {
			$(el).parent().removeClass('is-checked');
		}
	},
	focusin: function (el) {
		if ($(el).is(':focus')) {
			$(el).parent().addClass('is-focus');
		} else {
			$(el).parent().removeClass('is-focus');
		}
	},
	allchk: function (el) {
		var $target = $('[data-allparts=' + $(el).data('allchk') + ']');
		if ($(el).is(':checked')) {
			$target.prop('checked', true).trigger('change');
		} else {
			$target.prop('checked', false).trigger('change');
		}
	},
	allparts: function (el) {
		var
			$total = $('[data-allparts=' + $(el).data('allparts') + ']').length,
			$now = $('[data-allparts=' + $(el).data('allparts') + ']:checked').length;
		if ($total <= $now) {
			$('[data-allchk=' + $(el).data('allparts') + ']').prop('checked', true).parent().addClass('is-checked');
		} else {
			$('[data-allchk=' + $(el).data('allparts') + ']').prop('checked', false).parent().removeClass('is-checked');
		}
	}
};
var FormRadio = {
	build: function (el) {
		if ($(el).find('input').is(':disabled')) {
			$(el).addClass('is-disabled');
		}
		if ($(el).find('input').prop('readonly')) {
			$(el).addClass('is-readonly');
		}
		if ($(el).find('input').is(':checked')) {
			$(el).addClass('is-checked');
		}
	},
	change: function (el) {
		var groupName = $(el).attr('name');
		$('[name=' + groupName + ']').parent().removeClass('is-checked');
		$('[name=' + groupName + ']:checked').parent().addClass('is-checked');
	},
	focusin: function (el) {
		if ($(el).is(':focus')) {
			$(el).parent().addClass('is-focus');
		} else {
			$(el).parent().removeClass('is-focus');
		}
	}
};
var FormSelect = {
	build: function (el) {
		$('.value', el).text($('option:selected', el).text());
		if ($('select', el).is(':disabled')) {
			$(el).addClass('is-disabled');
		}
		if ($('select', el).prop('readonly')) {
			$(el).addClass('is-readonly');
		}
	},
	change: function (el) {
		$(el).parent().find('.value').text($('option:selected', el).text());
	},
	focusin: function (el) {
		if ($(el).is(':focus')) {
			$(el).parent().addClass('is-focus');
		} else {
			$(el).parent().removeClass('is-focus');
		}
	}
};

var FormInput = {
	build: function (el) {
		if ($(el).find('input').is(':disabled')) {
			$(el).addClass('is-disabled');
		}
		if ($(el).find('input').prop('readonly')) {
			$(el).addClass('is-readonly');
		}
		if ($(el).find('input').prop('type') == 'search') {
			$(el).append(
				$('<button></button>').attr('type', 'button').addClass('reset').append(
					$('<span></span>').addClass('blind').text('값 지우기')
				).on('click', function () {
					FormInput.valueReset(this);
				})
			)
		}
	},
	focusin: function (el) {
		if ($(el).is(':focus')) {
			$(el).parent().addClass('is-focus');
		} else {
			$(el).parent().removeClass('is-focus');
		}
	},
	valueReset: function (el) {
		$(el).prev().val('');
	}
};
var FormEmail = {
	change: function (el, mom) {
		if ($('option:selected', el).val() == 'self') {
			$('.email-address', mom).val('');
		} else {
			$('.email-address', mom).val($('option:selected', el).val());
		}
	}
};


var Tabs = {
	selector: {
		container: '.tab',
		panel: '.tab-panel',
		list: '.tab-selector li',
		item: '.tab-selector a'
	},
	build: function (i, el) {
		if (Tabs.getQueryStringUse()) {
			var
				pathname = location.search.queryStringToObject(),
				getLoc = $(Tabs.selector.container).data('use');

			if (pathname[getLoc] == undefined) {
				$(el).find(Tabs.selector.panel).eq(0).addClass('active');
				$(el).find(Tabs.selector.list + ':first-child').addClass('active');
			} else {
				$(el).find(Tabs.selector.panel).eq(parseInt(pathname[getLoc])).addClass('active');
				$(el).find(Tabs.selector.list).eq(parseInt(pathname[getLoc])).addClass('active');
			}
		} else {
			$(el).find(Tabs.selector.panel).eq(0).addClass('active');
			$(el).find(Tabs.selector.list + ':first-child').addClass('active');
		}
	},
	getQueryStringUse: function () {
		if ($(Tabs.selector.container).data('use') == undefined) {
			return false;
		} else {
			return true;
		}
	},
	getData: function (el) {
		$wrap = '#' + $(el).closest(Tabs.selector.container).attr('id');
		$target = $($wrap).find($('#' + $(el).data('target')));
		Tabs.openPanel($target, $wrap);
		Tabs.classChange(el, $wrap);
	},
	classChange: function (el, wr) {
		$(wr).find(Tabs.selector.list).removeClass('active');
		$(el).parent(Tabs.selector.list).addClass('active');
	},
	openPanel: function (el, wr) {
		$(wr).find(Tabs.selector.panel).removeClass('active');
		$(el).addClass('active');
	}
};


var commonUi = function () {
	var $win = $(window),
		selector,
		module

	selector = {
		header: '#header',
		footer: '#footer',
		container: '#container',
		content: '#container > .content',
		gnb: '#gnb',
		nav: '.gnb-list', // *** 수정 *** 190110 : gnb 액션 수정
		sort: '.wrap-sort',
		bnb: '#bnb',
		footerToggler: '#footer .btn-more',
		footerDetail: '#footer .company-detail'
	}
	module = {
		init: function () {

			// 슈마커 정보확인 열고 닫기
			$(selector.footerToggler).on('click', function () {
				if ($(selector.footer).hasClass('expanded')) {
					module.footer.close();
				} else {
					module.footer.open();
				}
			});

			// 상품목록 랭킹 슬라이드 init
			var rankingSlider = new Swiper('.ranking-slider', {
				slidesPerView: 'auto', // *** 수정 *** 190110 : 간격조절 수정
				spaceBetween: 20,
				// centeredSlides: true, // *** 수정 *** 190110 : 간격조절 수정
				observer: true,
				observeParents: true,
				on: {
					observerUpdate: true
				}

			});

			// 상품목록 분류 슬라이드 init
			var itemGroup = new Swiper('.item-group', {
				slidesPerView: 3.5,
				spaceBetween: 10,
				observer: true,
				observeParents: true
			});

			//브랜드목록 분류 슬라이드
			var brandGroup = new Swiper('.brand-group', {
				slidesPerView: 3.5,
				spaceBetween: 10,
				observer: true,
				observeParents: true
			});

			// 상품목록 이벤트 슬라이드 init
			var evtSlide = new Swiper('.evt-slider', {
				slidesPerView: 'auto',
				spaceBetween: 5,
				centeredSlides: true,
				observer: true,
				observeParents: true,
				pagination: {
					el: '.swiper-pagination',
					clickable: true
				},
			});

			// 상품상세 이미지 슬라이드 init
			var detailImgSlide = new Swiper('#detailImg', {
				observer: true,
				observeParents: true,
				scrollbar: {
					el: '.swiper-scrollbar',
					hide: false,
					draggable: true
				}
			});

			// 상품상세 컬러 슬라이드 init
			var colorList = new Swiper('.color-list', {
				slidesPerView: 4,
				spaceBetween: 10,
				observer: true,
				observeParents: true
			});

			// 상품상세 확대보기 슬라이드 init
			/*
			var zoomControl = new Swiper('.zoom-control', {
				slidesPerView: 1,
				zoom: {
					maxRatio: 3,
					minRatio: 1,
					containerClass: 'swiper-zoom-container',
				},
				centeredSlides: true,
				observer: true,
				observeParents: true,
				pagination: {
					el: '.swiper-pagination',
					clickable: true
				},
				navigation: {
					nextEl: '.swiper-button-next',
					prevEl: '.swiper-button-prev',
				},
			});
			*/

			// 상품상세 상품후기 이미지 init
			/*
			var reviewImg = new Swiper('.review-img', {
				centeredSlides: true,
				observer: true,
				observeParents: true,
				pagination: {
					el: '.swiper-pagination',
					clickable: true
				},
				navigation: {
					nextEl: '.swiper-button-next',
					prevEl: '.swiper-button-prev',
				},
			});
			*/

			//sale swiper
			var saleSwiper = new Swiper('.sale-swiper', {
				slidesPerView: 1,
				spaceBetween: 5,
				centeredSlides: true,
				observer: true,
				observeParents: true,
				pagination: {
					el: '.swiper-pagination',
					clickable: true
				},
			});

			//main-swiper
			/*
			var mainSwiper = new Swiper('.main-swiper', {
				slidesPerView: 1,
				spaceBetween: 5,
				centeredSlides: true,
              autoplay: {
                delay: 2000,  
              },
              autoplayDisableOnInteraction: true,
				observer: true,
				observeParents: true,
				pagination: {
					el: '.swiper-pagination',
					clickable: true
				},
			});
			*/
			//main>md-swiper    
			var mdSwiper = new Swiper('.md-swiper', {
				slidesPerView: 1.5,
				spaceBetween: 10,
			});
			//main>style-swiper        
			var styleSwiper = new Swiper('.style-swiper', {
				slidesPerView: 'auto',  // *** 수정 *** 190110 : 간격조절 수정
				spaceBetween: 10,
			});
			//main>review-swiper        
			var reviewSwiper = new Swiper('.review-swiper', {
				slidesPerView: 2,
				slidesPerGroup: 2,
				spaceBetween: 10,
				pagination: {
					el: '.swiper-pagination',
					clickable: true
				},
			});
			//main>style-swiper        
			var streetSwiper = new Swiper('.street-swiper', {
				slidesPerView: 1,
				spaceBetween: 10,

				pagination: {
					el: '.swiper-pagination',
					clickable: true
				}
			});
			/* ====================brandHall Nike Swipers================= */
			//NIKE-swiper 
			var nikeSwiper1 = new Swiper('.nike-swiper', {
				slidesPerView: 1,
				spaceBetween: 10,

				pagination: {
					el: '.swiper-pagination',
					clickable: true,
					renderBullet: function (index, className) {
						return '<span class="' + className + '">' + '0' + (index + 1) + '</span>';
					}
				}
			});

			var nikeSwiper2 = new Swiper('.nike-swiper2', {
				slidesPerView: 1.7,
				spaceBetween: 10,
				centeredSlides: true,
			});

			/* ====================//brandHall Nike Swipers================= */

			// 상품상세 관련이벤트 init
			var swiper = new Swiper('.more-event', {
				slidesPerView: 'auto',
				spaceBetween: 5,
				centeredSlides: true,
				observer: true,
				observeParents: true,
				scrollbar: {
					el: '.swiper-scrollbar',
					hide: false,
					draggable: true
				},
			});

			// TIME SALE POP *** 수정 *** 190104 : Time Sale Slider -->
			/*
            var swiper = new Swiper('.timeSale-slide .swiper-container', {
                slidesPerView: 1,
                loop: true,
                direction: 'horizontal',
                autoplay: {
                    delay: 2500,
                    disableOnInteraction: false,
                },
                pagination: {
                    el: '.swiper-pagination',
                    clickable: true,
                },
                observer: true,
                observeParents: true
            });
			*/
			// 메인 Layer Pop Banner *** 수정 *** 190104 : 메인 Layer Pop Banner -->
			/*
            var bannerSlide = new Swiper('.pop-banner-slide .swiper-container', {
                slidesPerView: 1,
                loop: true,
                loopedSlides: 10,
                autoplay: {
                    delay: 2500,
                    disableOnInteraction: false,
                },
                pagination: {
                    el: '.swiper-pagination',
                    clickable: true,
                },
                observer: true,
                observeParents: true
            });
			*/

			// BRADS 메인 슬라이드 *** 수정 *** 190114 : BRANDS 메인 슬라이드 추가
			var brandsMain = new Swiper('.brand-main .swiper-container', {
				slidesPerView: 1,
				loop: true,
				autoplay: {
					delay: 3000,
					disableOnInteraction: false,
				},
				pagination: {
					el: '.swiper-pagination',
					clickable: true,
				},
				observer: true,
				observeParents: true
			});

			// *** 수정 *** 190118 : 푸터 최근 본 상품 아이콘 슬라이드 추가
			var relatedItem = new Swiper('.relatedItem .swiper-container', {
				slidesPerView: 1,
				loop: true,
				autoplay: {
					delay: 2000,
					disableOnInteraction: false,
				},
				observer: true,
				observeParents: true
			});

			$(window).on('load resize', function () {
				var widHei = $(window).height();
				var titPop = $('.tit-pop').outerHeight();

				//브랜드 리스트 슬라이드 사이즈 Custom
				var brandList = $('.brand-line .brand-category .brand-group');
				var imgML = (brandList.find('.img').width() - brandList.find('.img img').width()) / 2;
				brandList.find('.txt').css('padding-top', '0');
				brandList.find('.img img').css('margin-left', imgML);

				// 상품상세 이미지 슬라이드 사이즈 Custom
				var detailImgList = $('#detailImg');
				var slideLength = detailImgList.find('.swiper-slide').length;
				var slideWid = detailImgList.find('.swiper-slide').width();

				$(detailImgList).find('.swiper-slide').css('height', slideWid + 'px');

				// 상품목록 분류 슬라이드 사이즈 Custom
				var itemList = $('.color-list');
				var slideWid = itemList.find('.swiper-slide').width();

				$(itemList).find('.swiper-slide').css({
					'height': slideWid
				});
			});

			// TODO : 클릭 시 해당 메뉴명에 active 효과주고 화면 플리킹 연결
			// TODO : 상품 상세 필터 코딩

			// gnb sticky
			/*
			$win.on('load scroll', function(){
					if(module.gnb.child() >= 2){
					if($win.scrollTop() >= module.gnb.getTop2()){
						module.gnb.fix();
						module.nav.none();
					} else {
						module.gnb.release();
						module.nav.block();
					}
				}
				else{
					if($win.scrollTop() >= module.gnb.getTop()){
						module.gnb.fix();
					} else {
						module.gnb.release();
					}
				}
			
				if ($win.scrollTop() >= module.gnb.getTop()) {
					module.gnb.fix();
				} else {
					module.gnb.release();
				}
			});
			*/
		},
		gnb: {
			fix: function () {
				$(selector.gnb).addClass('fixed')
			},
			release: function () {
				$(selector.gnb).removeClass('fixed')
			},
			getTop: function () {
				return $(selector.header).outerHeight()
			},
			getTop2: function () {
				return $(selector.header).outerHeight() + $(selector.nav).outerHeight();
			},
			child: function () {
				return $(selector.gnb).children().length;
			}
		},
		nav: {
			none: function () {
				$(selector.nav).addClass('none');
			},
			block: function () {
				$(selector.nav).removeClass('none');
			}
		},
		content: {
			add: function () {
				$(selector.content).addClass('add');
			},
			remove: function () {
				$(selector.content).removeClass('add');
			}
		},
		footer: {
			open: function () {
				$(selector.footer).addClass('expanded');
				$(selector.container).addClass('expanded');
				module.footer.animation(module.footer.getHeight(), 300);
				$('html, body').animate({ scrollTop: $(document).height() }, '300'); // *** 수정 *** 190117 :  '슈마커 정보확인' 버튼 클릭시 숨겨져서 보임 - 추가내용
			},
			close: function () {
				$(selector.footer).removeClass('expanded');
				$(selector.container).removeClass('expanded');
				module.footer.animation(0, 300)
			},
			animation: function (height, duration) {
				$(selector.footerDetail).css('height', height);
				setTimeout(function () {
					cssTransition($(selector.footerDetail), duration);
				}, 1);
				// setTimeout(function(){    // *** 수정 *** 190117 :  '슈마커 정보확인' 버튼 클릭시 숨겨져서 보임 - 주석처리
				// 	$(selector.container).css('margin-bottom', '-' + module.footer.getContH() + 'px');
				// 	$(selector.content).css('padding-bottom', + module.footer.getContH() + 'px');
				// 	cssTransition($(selector.container), duration);
				// 	cssTransition($(selector.content), duration);
				// }, 301)
			},
			getHeight: function () {
				return $('.inner', selector.footerDetail).outerHeight()
			},
			getContH: function () {
				return $(selector.footer).outerHeight()
			}
		}
	};
	module.init()
};

var formInit = function () {
	$('.checkbox').each(function (i, el) {
		FormCheckbox.build(el);
		$(el).find('input').on('change', function () {
			FormCheckbox.change(this);
			if ($(this).data('allchk') != undefined) {
				FormCheckbox.allchk(this);
			} else if ($(this).data('allparts') != undefined) {
				FormCheckbox.allparts(this);
			}
		});
		$(el).find('input').on('focus blur click', function () {
			FormCheckbox.focusin(this);
		});
	});
	$('.radio, .radio2').each(function (i, el) {
		FormRadio.build(el);
		$(el).find('input').on('change', function () {
			FormRadio.change(this);
		});
		$(el).find('input').on('focus blur click', function () {
			FormRadio.focusin(this);
		});
	});
	$('.select').each(function (i, el) {
		FormSelect.build(el);
		$(el).find('select').on('change', function () {
			FormSelect.change(this);
		});
		$(el).find('select').on('focus blur click', function () {
			FormSelect.focusin(this);
		});
	});
	$('.input').each(function (i, el) {
		FormInput.build(el);
		$(el).find('input').on('focus blur click touchend', function () {
			FormInput.focusin(this);
		});
	});
	$('.ty-email').each(function (i, el) {
		FormEmail.change(this, el);
		$('.domain-helper', el).on('change', function () {
			FormEmail.change(this, el);
		});
	});
};
var tabBuild = function () {
	$(Tabs.selector.container).each(function (i, el) {
		Tabs.build(i, el)
	});
	$(Tabs.selector.item).on('click', function () {
		Tabs.getData(this)
	});
};

//popup
var popCall = function () {
	$('.area-pop>div').each(function () {

		var _this = $(this);

		_this.closest('body').addClass('posFixed');

		var _windowHeight = $(window).height();
		var _popHeight = _this.height();
		var _contentHeight = _this.find('.contents').outerHeight();
		var _btnHeight = _this.find('.btns').height();
		var _maxHeight = _windowHeight - _btnHeight;

		if (_this.is('.alert')) {
			if (_popHeight > _maxHeight) {
				_this.css('height', _maxHeight);
				_this.addClass('scroll');
			}
		}
		else if (_this.is('.full')) {
			if (_this.find('.container-pop').children().length === 1) {
				_this.addClass('scroll');
			}
		}
			// *** 수정*** 191011 : 찜, 최근 본 상품 등 팝업 스타일 추가
		else if (_this.is('.top-exposed')) {
			_this.css('height', _windowHeight - 82 + 'px');
			_this.closest('.wrap-pop').removeClass('hidden');

			_this.find('.btn-hide-pop').on('click', function () {
				_this.removeClass('vertical');
				_this.closest('body').removeClass('posFixed');
				_this.closest('.wrap-pop').addClass('hidden');
			});
		}

			//  *** 수정 *** 190115 :타임세일 팝업 호출 전
		else if (_this.is('.ly-timeSale')) {
			_this.closest('body').removeClass('posFixed');
			_this.closest('.wrap-pop').css('display', 'none');

			_this.find('.btn-hide').on('click', function () {
				_this.closest('.wrap-pop').next().removeClass('hide');
				_this.closest('.wrap-pop').css('display', 'none');
				_this.closest('body').removeClass('posFixed');
			});
		}
	});
};

//  *** 수정 *** 190115 :타임세일 아이콘 추가, 타임세일 호출 이벤트
$('.ico-timesale>button').on('click', function () {
	var __this = $(this);
	__this.parent().addClass('hide');
	__this.closest('body').addClass('posFixed');

	__this.parent().prev().css('display', 'block');
});

// sorting num
var sortNum = function () {
	var sort = $('.wrap-sort');

	if (sort.find('>div').length > 2) {
		sort.find('>div').css('width', 33.333333 + '%');
	} else {
		sort.find('>div').css('width', 50 + '%');
	}
};

//가격대 슬라이더
/*
$( function() {
	$( '.range-bar' ).slider({
		range: true,
		min: 0,
		max: 100,
		values: [ 0, 50 ],
		slide: function( event, ui ) {
			$( '#amount' ).val(ui.values[ 0 ]+'만원' + ' - ' + ui.values[ 1 ] +'만원');
		}
	});
	$( '#amount' ).val($( '.range-bar' ).slider( 'values', 0 )+'만원' +
		' ~ ' + $( '.range-bar' ).slider( 'values', 1 ) +'만원');
} );
*/

//브랜드 리스트 찜하기 버튼 클릭
$('.wrap-brand-list .brand-bg .called').on('click', function () {
	($(this).hasClass('on')) ? $(this).removeClass('on') : $(this).addClass('on')
});


//timedeal
function timedeal() {
	var now = new Date();
	var dday = new Date(2018, 10, 12, 18, 00, 00);

	var days = (dday - now) / 1000 / 60 / 60 / 24;
	var daysRound = Math.floor(days);
	var hours = (dday - now) / 1000 / 60 / 60 - (24 * daysRound);
	var hoursRound = Math.floor(hours);
	var minutes = (dday - now) / 1000 / 60 - (24 * 60 * daysRound) - (60 * hoursRound);
	var minutesRound = Math.floor(minutes);
	var seconds = (dday - now) / 1000 - (24 * 60 * 60 * daysRound) - (60 * 60 * hoursRound) - (60 * minutesRound);
	var secondsRound = Math.floor(seconds);
	var miliseconds = ((dday - now) - (24 * 60 * 60 * 1000 * daysRound) - (60 * 60 * 1000 * hoursRound) - (60 * 1000 * minutesRound) - (1000 * secondsRound)) / 10;
	var milisecondsRound = Math.round(miliseconds);

	if (minutesRound < 10) minutesRound = '0' + minutesRound;
	if (secondsRound < 10) secondsRound = '0' + secondsRound;
	if (milisecondsRound < 10) milisecondsRound = '0' + milisecondsRound;

	var todayresult = hoursRound + ':' + minutesRound + ':' + secondsRound + ':' + milisecondsRound;

	document.getElementById('remaintime').innerHTML = todayresult;
}

//timesale
function timesale() {
	var now = new Date();
	var dday = new Date(2018, 11, 15, 18, 00, 00);

	var days = (dday - now) / 1000 / 60 / 60 / 24;
	var daysRound = Math.floor(days);
	var hours = (dday - now) / 1000 / 60 / 60 - (24 * daysRound);
	var hoursRound = Math.floor(hours);
	var minutes = (dday - now) / 1000 / 60 - (24 * 60 * daysRound) - (60 * hoursRound);
	var minutesRound = Math.floor(minutes);
	var seconds = (dday - now) / 1000 - (24 * 60 * 60 * daysRound) - (60 * 60 * hoursRound) - (60 * minutesRound);
	var secondsRound = Math.floor(seconds);

	if (minutesRound < 10) minutesRound = '0' + minutesRound;
	if (secondsRound < 10) secondsRound = '0' + secondsRound;

	var timeresult = hoursRound + ':' + minutesRound + ':' + secondsRound

	document.getElementById('timesale').innerHTML = timeresult;
}

$(document).ready(function () {
	commonUi();
	formInit();
	tabBuild();
	//popCall();
	sortNum();

	/*
    if ($('#remaintime').get(0) != undefined) {
        setInterval(timedeal, 1);
    }
    if ($('#timesale').get(0) != undefined) {
        setInterval(timesale, 1);
    }
	*/
});


// 상품 상세 상품후기, 상품정보, 상품문의, 주문/배송, 교환/반품/AS 아코디언
var detailAccord = function () {
	var selector,
		module;

	selector = {
		bgImg: '.bg-img',
		starBg: '.satisfy',
		reviewBg: '.sect',
		button: '.btn-accord-ty1',
		toggler: '.accordion-selector',
		panel: '.accordion-panel'
	}
	module = {
		init: function () {
			$(selector.button).on('click', function () {
				module.accordion(this);
			});
			$(window).trigger('scroll');
			$(selector.button).trigger('click');
		},
		accordion: function (el) {
			var target = $(el).data('target');

			$(selector.toggler).addClass('folded');
			$(selector.panel).slideUp(400);

			$(selector.button).removeClass('on');

			/*
			$(selector.starBg).animate({ 'backgroundColor': 'transparent' }, 300);
			$(selector.reviewBg).animate({ 'backgroundColor': 'transparent' }, 300);
			$(selector.bgImg).removeClass('change');
			*/

			if ($(selector.panel, '#' + target).css('display') === 'none') {
				$(selector.panel, '#' + target).slideDown();
				$(selector.panel, '#' + target).prev().find('button').addClass('on');

				/*
				$(selector.bgImg).addClass('change');
				$(selector.starBg).animate({ 'backgroundColor': '#016a4c' }, 500);
				$(selector.reviewBg).animate({ 'backgroundColor': '#014f39' }, 500);
				*/
			}
		}
	}
	module.init();
}();

// 상품상세 상품옵션 창 슬라이드
/*
var optionSelectShow = function(){
	$('.bnb-ty2 .btn-buy').on('click', function(){
		$('.area-select-option').addClass('is-block');
	});
	$('.area-select-option .btn-hide-select').on('click', function(){
		$('.area-select-option').removeClass('is-block');
	});
}();
*/

// 상품상세 상품옵션 선택 아코디언
var detailSelectOption = function () {
	var selector,
		module;

	selector = {
		button: '.clickEvt',
		toggler: '.selector',
		panel: '.option'
	};

	module = {
		init: function () {
			$(selector.button).on('click', function () {
				module.accordion(this);
			});
			$(window).trigger('scroll');
		},
		accordion: function (el) {
			var target = $(el).data('target');

			$(selector.panel).slideUp(400);
			$(selector.toggler).removeClass('is-focus');

			if ($(selector.panel, '#' + target).css('display') === 'none') {
				$(selector.panel, '#' + target).slideDown();
				$(selector.toggler, '#' + target).addClass('is-focus');
			}
		}
	};
	module.init();
}();

//top버튼
$('.move-top').find('button').on('click', function () {
	$('html, body').animate({ scrollTop: 0 }, 400);
});

// 팝업 아코디언
var popAccord = function () {
	var selector,
		module;

	selector = {
		button: '.pop-accordion-selector>button',
		toggler: '.pop-accordion-selector',
		panel: '.pop-accordion-panel'
	}
	module = {
		init: function () {
			$(selector.button).on('click', function () {
				module.accordion(this);
			});
			$(window).trigger('scroll');
			$(selector.button).trigger('click');
		},
		accordion: function (el) {
			var target = $(el).data('target');

			$(selector.toggler).removeClass('current');
			$(selector.panel).slideUp();

			if ($(selector.panel, '#' + target).css('display') === 'none') {
				$(selector.toggler, '#' + target).addClass('current');
				$(selector.panel, '#' + target).slideDown();
			}
		}
	}
	module.init();
}();

// gnb 슬라이딩 *** 수정 *** 190110 : gnb 액션 수정
var gnbSlide = function () {
	var gnbList = $('.gnb-list li');

	// street306, only 메뉴 정렬
	if (gnbList.length < 5) {
		gnbList.css({ 'width': 100 / 4 + '%' }).addClass('length4');
		gnbList.find('>a').css({ 'width': 'auto', 'padding': '0 5' }); // *** 수정 190114 : 간격 수정 추가***
	}

	// gnb 메뉴 스크롤 이동
	$('.gnb-list li').each(function () {
		var ceil = Math.ceil($(this).find('a').outerWidth());
		$(this).find('a').outerWidth(ceil); // 소수점 올림

		$(this).on('click', function () {
			var prevWidth = $(this).prev().outerWidth();
			var prevAllWidth = 0;

			$(this).prevAll().each(function () {
				prevAllWidth += $(this).outerWidth();
			});

			$('.gnb-list').scrollLeft(prevAllWidth - prevWidth);
			$(this).addClass('current').siblings().removeClass('current');
		});
	});

	// 선택한 메뉴를 앞으로 스크롤 이동
	if (gnbList.hasClass('current')) {
		var loadCurrent = $('.gnb-list li.current');
		var loadPrevWidth = loadCurrent.prev().outerWidth();
		var loadPrevAllWidth = 0;

		loadCurrent.prevAll().each(function () {
			loadPrevAllWidth += $(this).outerWidth();
		});

		$('.gnb-list').scrollLeft(loadPrevAllWidth - loadPrevWidth);
	}
}();

//street306 정렬
/*
$(window).on('load', function () {
        $('#grid').masonry({
          columnWidth: '.card',
          itemSelector: '.card',
          gutter:5,
          horizontalOrder: true,
          percentPosition: true
    });
});
*/

/*** 마이페이지 ***/
// 마이페이지 아코디언
var mypageAccodion = function () {
	var selector,
		module;

	selector = {
		parent: '.accord-mypage',
		button: '.clickEvt',
		toggler: '.ly-title',
		panel: '.ly-content',
		// 주문취소/반품/교환 내 아코디언 안에 아코디언
		parent_sub: '.accord-sub-mypage',
		button_sub: '.clickEvt_sub',
		toggler_sub: '.ly-title_sub',
		panel_sub: '.ly-content_sub'
	};

	module = {
		init: function () {
			$(selector.button).on('click', function () {
				module.accordion(this);
			});
			$(selector.button_sub).on('click', function () {
				module.accordion_sub(this);
			});
			$(window).trigger('scroll');
			$(selector.parent).eq(0).find($(selector.panel)).show();
			$(selector.parent).eq(0).find($(selector.toggler)).addClass('is-on');
			$(selector.parent_sub).eq(0).find($(selector.panel_sub)).show();
			$(selector.parent_sub).eq(0).find($(selector.toggler_sub)).addClass('is-on');
		},
		accordion: function (el) {
			var target = $(el).data('target');

			$(selector.panel).slideUp(400);
			$(selector.toggler).removeClass('is-on');

			if ($(selector.panel, '#' + target).css('display') === 'none') {
				$(selector.panel, '#' + target).slideDown(400);
				$(selector.toggler, '#' + target).addClass('is-on');
			}
		},
		accordion_sub: function (el) {
			var target = $(el).data('target');

			$(selector.panel_sub).slideUp(300);
			$(selector.toggler_sub).removeClass('is-on');

			if ($(selector.panel_sub, '#' + target).css('display') === 'none') {
				$(selector.panel_sub, '#' + target).slideDown(300);
				$(selector.toggler_sub, '#' + target).addClass('is-on');
			}
		}
	};
	module.init();
}();

// DatePicker
/*
var datePiker = function(){
	var dateFormat = "yy-mm-dd",
		from = $( ".date-from" ).datepicker({
			changeMonth: true,
			changeYear: true,
			numberOfMonths: 1,
			showOn: "button",
			buttonImage: "../../assets/images/ico/btn-calendar.png",
			buttonImageOnly: true,
			buttonText: "Select date",
			monthNames: ['1월', '2월', '3월', '4월', '5월', '6월', '7월', '8월', '9월', '10월', '11월', '12월'],
			monthNamesShort: ['1월', '2월', '3월', '4월', '5월', '6월', '7월', '8월', '9월', '10월', '11월', '12월'],
			dayNames: ['일', '월', '화', '수', '목', '금', '토'],
			dayNamesShort: ['일', '월', '화', '수', '목', '금', '토'],
			dayNamesMin: ['일', '월', '화', '수', '목', '금', '토'],
			showMonthAfterYear: true
		})
			.on( "change", function() {
				to.datepicker( "option", "minDate", getDate( this ) );
			}),
		to = $( ".date-to" ).datepicker({
			changeMonth: true,
			changeYear: true,
			numberOfMonths:1,
			showOn: "button",
			buttonImage: "../../assets/images/ico/btn-calendar.png",
			buttonImageOnly: true,
			buttonText: "Select date",
			monthNames: ['1월', '2월', '3월', '4월', '5월', '6월', '7월', '8월', '9월', '10월', '11월', '12월'],
			monthNamesShort: ['1월', '2월', '3월', '4월', '5월', '6월', '7월', '8월', '9월', '10월', '11월', '12월'],
			dayNames: ['일', '월', '화', '수', '목', '금', '토'],
			dayNamesShort: ['일', '월', '화', '수', '목', '금', '토'],
			dayNamesMin: ['일', '월', '화', '수', '목', '금', '토'],
			showMonthAfterYear: true
		})
			.on( "change", function() {
				from.datepicker( "option", "maxDate", getDate( this ) );
			});

	function getDate(element) {
		var date;
		try {
			date = $.datepicker.parseDate( dateFormat, element.value );
		} catch( error ) {
			date = null;
		}

		return date;
	}
}();
*/

// A/S 신청내역 더보기
$('.btn-more-txt').on('click', function () {
	$(this).toggleClass('open');
	$(this).prev('a').find('.paragraph').toggleClass('open');
});

// 포인트 내역보기
var pointCheck = function () {
	var selector,
		module;

	selector = {
		button: '.btn-more',
		toggler: '.cont',
		panel: '.check'
	};

	module = {
		init: function () {
			$(selector.button).on('click', function () {
				module.accordion(this);
			});
			$(window).trigger('scroll');
		},
		accordion: function (el) {
			var target = $(el).data('target');

			$(selector.panel).slideUp(400);
			$(selector.toggler).removeClass('is-on');

			if ($(selector.panel, '#' + target).css('display') === 'none') {
				$(selector.panel, '#' + target).slideDown(400);
				$(selector.toggler, '#' + target).addClass('is-on');
			}
		}
	};
	module.init();
}();


// 상품후기 별점 선택
//별점 관련
$('.star-grade').each(function () {
	var _this = $(this);
	var _thisSpan = _this.children('span');

	_thisSpan.click(function () {
		var _spanIndex = $(this).index();
		var _starNum = (_spanIndex + 1) / 2;
		$(this).closest('.post-body').children('.star-num').text(_starNum.toFixed(1));

		$(this).parent().children('span').removeClass('on');
		$(this).addClass('on').prevAll('span').addClass('on');
		return false;
	});
});

// 상품문의 내용 더 보기 *** 수정 *** 190115 : 댓글 폼 추가
$('.btn-toggle>button').on('click', function () {
	$(this).parent().prev().find('.tit').toggleClass('all');
});