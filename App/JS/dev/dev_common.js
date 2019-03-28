/* GNB 장바구니 건수 가져오기 */
$(function () {
	get_GNB_CartCount();
});

/* GNB 장바구니 건수 가져오기 */
function get_GNB_CartCount() {
	$.ajax({
		type: "post",
		url: "/Common/Ajax/GNB_CartCount.asp",
		async: false,
		data: "",
		dataType: "text",
		success: function (data) {
			$("#GNB_CartCount").html(data);
		},
		error: function (data) {
			alert(data.responseText)//alert("처리 도중 오류가 발생하였습니다.");
		}
	});
}

/* LOGIN 페이지로 이동 */
function move_Login(loc) {
	var progID = $("input[name='ProgID']", "form[name='botLoginForm']").val();
	if (progID.toLowerCase().indexOf('login.asp') > -1) {
		$("input[name='ProgID']", "form[name='botLoginForm']").val("/");
	}
	if (loc == "S") {
		document.botLoginForm.action = "/ASP/Member/SubLogin.asp";
	}
	else if (loc == "O") {
		document.botLoginForm.action = "/ASP/Order/Login.asp";
	}
	else {
		document.botLoginForm.action = "/ASP/Member/Login.asp";
	}
	document.botLoginForm.submit();
}


/* 본문 페이지 가져오기 */
function common_getPage(pid, pType) {
	var ajaxUrl = "";
	var ajaxData = "";
	if (pType == "ShoesGiftAdd") {
		ajaxUrl = "/ASP/Mypage/Ajax/ShoesGiftAdd.asp";
		ajaxData = "";
	} else if (pType == "MyReviewList") {		//마이페이지 > 리뷰 리스트
		ajaxUrl = "/ASP/Mypage/Ajax/MyReviewList.asp";
		ajaxData = "";
	} else if (pType == "MyQnaList") {		//마이페이지 > 상품Q&A, 1:1문의 리스트
		ajaxUrl = "/ASP/Mypage/Ajax/MyQnaList.asp";
		ajaxData = "";
	} else if (pType == "MyAddrList") {			//마이페이지 > 배송지관리 리스트
		ajaxUrl = "/ASP/Mypage/Ajax/MyAddrList.asp";
		ajaxData = "";
	} else if (pType == "MyAddr") {				//마이페이지 > 배송지관리 입력/수정
		ajaxUrl = "/ASP/Mypage/Ajax/MyAddr.asp";
		var addrType = $("input[name='addrType']", "form[name='MyAddrListForm']").val();
		if (addrType == "insert") {				//배송지관리 입력
			ajaxData = "addrType=insert";
		} else if (addrType == "modify") {		//배송지관리 수정
			ajaxData = "addrType=modify&idx=" + $("input[name='idx']:checked", "form[name='MyAddrListForm']").val();
		}
	} else if (pType == "MyInfoModifyPwChk") {	//마이페이지 > 회원정보 수정(비밀번호 확인)
		ajaxUrl = "/ASP/Mypage/Ajax/MyInfoModifyPwChk.asp";
		ajaxData = "";
	} else if (pType == "MyInfoModify") {		//마이페이지 > 회원정보 수정
		ajaxUrl = "/ASP/Mypage/Ajax/MyInfoModify.asp";
		ajaxData = "";
	} else if (pType == "Withdraw") {			//마이페이지 > 회원탈퇴
		ajaxUrl = "/ASP/Mypage/Ajax/Withdraw.asp";
		ajaxData = "";
	} else if (pType == "MySnsList") {			//SNS계정정보
		ajaxUrl = "/ASP/Mypage/Ajax/MySnsList.asp";
		ajaxData = "";
	} else if (pType == "JoinChgMem") {			//회원전환
		ajaxUrl = "/ASP/Member/Ajax/snsSDupInfoView.asp";
		ajaxData = "";
	}

	$.ajax({
		type: "post",
		url: ajaxUrl,
		async: false,
		data: ajaxData,
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				$("#" + pid).html(cont);
			} else if (result == "LOGIN") {
				common_msgPopOpen("", cont, "/ASP/Member/Login.asp");
				return;
			} else {
				common_msgPopOpen("", cont);
				return;
			}
		},
		error: function (data) {
			document.write(data.responseText);
			common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.");
		}
	});
}


/* 회원가입 및 기타 이용약관 띄우기 */
function PolicyView(sType) {
	APP_PopupGoUrl("/ASP/Member/Policy.asp?SType=" + sType);
}

/* 이용약관 닫기 */
function PolicyClose() {
	$('#PolicyPopup').hide();
	$('body').removeClass('noscroll');
}

function regularLogin() {
	common_msgPopOpen("", "정회원만 이용 가능한 메뉴입니다.");
	return;
}

/* 팝업 공통메시지 레이어 열기 */
function common_msgPopOpen(title, msg, script, focus, pStyle) {
	//Title:팝업타이틀, Msg:팝업내용, Script:사용자함수 또는 스크립트(확인버튼 클릭 시 처리), Focus:팝업close시 focus처리(form.Name Or Name), pStyle:팝업스타일(N:일반, C:confirm형)
	//common_msgPopOpen('테스트타이틀', '테스트메시지<br>테스트메시지', 'location.href=\'/ASP/Member/Login.asp\'', 'form1.userID', 'N');
	//common_msgPopOpen('테스트타이틀','테스트메시지<br>테스트메시지','aa();');
	ajaxData = "Title=" + encodeURIComponent(title) + "&Msg=" + encodeURIComponent(msg) + "&Script=" + script + "&Focus=" + focus + "&pStyle=" + pStyle;

	$.ajax({
		type: "post",
		url: "/Common/Ajax/messagePopup.asp",
		async: false,
		data: ajaxData,
		dataType: "html",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				//$("#msgPopup").html(cont);
				$("#msgPopup").html(cont);
				$("#msgPopup").show();

				//Pop Up 높이 값
				var _this = $("#msgPopup .area-pop");
				var _windowHeight = $(window).height();
				var _maxHeight = _windowHeight - 100;

				// Pop Up 호출 시 전체 스크롤 제거
				//$('body').addClass('noscroll');
				$("body").css("overflow", "hidden");
			}
			else {
				alert(cont);
			}
		},
		error: function (data) {
			//alert(data.responseText)
			alert("처리 도중 오류가 발생하였습니다.");
		}
	});
}

/* 팝업 공통메시지 레이어 닫기 */
function common_msgPopClose(formF) {
	$("#msgPopup").hide();
	//$('body').removeClass('noscroll');
	$("body").css("overflow", "auto");
	if (formF != "") {
		//document.getElementsByName(formF)[0].focus();
		if (formF.indexOf(".") > 0) {
			var formArr = formF.split(".");
			var formF1 = formArr[0];
			var formF2 = formArr[1];
			$("input[name='" + formF2 + "']", "form[name='" + formF1 + "']").focus();
		} else {
			$("input[name='" + formF + "']").focus();
		}
	}
}


/* 팝업 레이어 오픈 */
function common_PopOpen(pid, pType) {
	var ajaxUrl = "";
	var ajaxData = "";
	if (pType == "ShoesGiftAdd") {
		ajaxUrl = "/ASP/Mypage/Ajax/ShoesGiftAdd.asp";
		ajaxData = "";
	} else if (pType == "MyReviewWrite") {					//마이페이지 상품후기 관리
		ajaxUrl = "/ASP/Mypage/Ajax/ReviewWrite.asp";
		ajaxData = $("form[name='MyReviewListForm']").serialize();
	} else if (pType == "MyMtmQnaWrite") {					//마이페이지 > 1:1문의 입력폼
		ajaxUrl = "/ASP/Mypage/Ajax/MyMtmQnaWrite.asp";
		ajaxData = "";
	} else if (pType == "MyAddr") {							//마이페이지 배송지주소록 관리
		ajaxUrl = "/ASP/Mypage/Ajax/MyAddr.asp";
		ajaxData = $("form[name='MyAddrListForm']").serialize();
	} else if (pType == "MyInfoModify") {					//마이페이지 회원정보 관리
		ajaxUrl = "/ASP/Mypage/Ajax/MyInfoModifyPwChk.asp";
		ajaxData = "";
	} else if (pType == "RefundAccountAdd") {				//마이페이지 > 회원정보 수정 > 계좌정보 입력/수정
		ajaxUrl = "/ASP/Mypage/Ajax/RefundAccountAdd.asp";
		ajaxData = $("#RefundAccountForm").serialize();
	} else if (pType == "JoinChgMem") {						//정회원 전환
		ajaxUrl = "/ASP/Member/Ajax/snsSDupInfoView.asp";
		ajaxData = "";
	} else if (pType == "NoticeView") {						//고객센터 > 공지사항 뷰
		ajaxUrl = "/ASP/Customer/Ajax/NoticeView.asp";
		ajaxData = "idx=" + $("input[name='Idx']", "form[name='NoticeListForm']").val();
	} else if (pType == "StoreView") {						//고객센터 > 전국매장안내 뷰
		ajaxUrl = "/ASP/Customer/Ajax/StoreView.asp";
		ajaxData = "";
	} else if (pType == "FooterNoticeView") {				//Footer 공지사항 뷰
		ajaxUrl = "/ASP/Customer/Ajax/NoticeView.asp";
		ajaxData = "idx=" + $("input[name='Idx']", "form[name='FooterNoticeViewForm']").val();
	}

	$.ajax({
		type: "post",
		url: ajaxUrl,
		async: false,
		data: ajaxData,
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				$("#" + pid).html(cont);
				$("#" + pid).show();

				//Pop Up 높이 값
				//var _this = $("#" + pid + " .area-pop");
				var _windowHeight = $(window).height();
				var _maxHeight = _windowHeight - 100;
				//_this.css('max-height', _maxHeight);
				//_this.closest('body').addClass('ofh');

				/*
				var popTop = (($("#" + pid).height() - $("#" + pid + " .area-pop").height()) / 2);
				alert(popTop)
				$("#" + pid + " .area-pop").css({ 'top-margin': popTop });
				*/

				// Pop Up 호출 시 전체 스크롤 제거
				//$('body').addClass('noscroll');
				$("body").css("overflow", "hidden");

				//openPop(pid);
			}
			else if (result == "LOGIN") {
				PareReload();
				return;
			}
			else {
				common_msgPopOpen("", "처리 중 오류가 발생하였습니다.[01]");
			}
		},
		error: function (data) {
			//alert(data.responseText)
			common_msgPopOpen("", "처리 중 오류가 발생하였습니다.[02]");
		}
	});
}

/* 팝업 레이어 닫기 */
function common_PopClose(pid) {
	$("#" + pid).hide();
	//$('body').removeClass('noscroll');
	$("body").css("overflow", "auto");
}

/* 메인페이지 BEST SELLER / NEW ARRIVALS 상품 */
function get_BestNArrivalsProductList(sCode0, sCode1) {
	$.ajax({
		type		 : "post",
		url			 : "/Common/Ajax/Main_BestNArrivalsProductList.asp",
		async		 : false,
		data		 : "SCode0=" + sCode0 + "&SCode1=" + sCode1,
		dataType	 : "text",
		success		 : function (data) {
						var splitData	 = data.split("|||||");
						var result		 = splitData[0];
						var cont		 = splitData[1];


						if (result == "OK") {
							$("#BestNArrivalsProductList").html(cont);
							return;
						}
						else {
							alert(cont);
							return;
						}
		},
		error		 : function (data) {
						alert(data.responseText)//alert("처리 도중 오류가 발생하였습니다.");
		}
	});
}

/* 메인페이지 BEST BRANDS 상품 */
function get_BestBrandsProductList(idx) {
	$.ajax({
		type		 : "post",
		url			 : "/Common/Ajax/Main_BestBrandsProductList.asp",
		async		 : false,
		data		 : "Idx=" + idx,
		dataType	 : "text",
		success		 : function (data) {
						var splitData	 = data.split("|||||");
						var result		 = splitData[0];
						var cont		 = splitData[1];


						if (result == "OK") {
							$("#BestBrandsProductList").html(cont);
							return;
						}
						else {
							alert(cont);
							return;
						}
		},
		error		 : function (data) {
						alert("처리 도중 오류가 발생하였습니다.");
		}
	});
}

/* 메인페이지 지금뜨는 상품 */
function get_NowBestProductList(idx) {
	$.ajax({
		type		 : "post",
		url			 : "/Common/Ajax/Main_NowBestProductList.asp",
		async		 : false,
		data		 : "Idx=" + idx,
		dataType	 : "text",
		success		 : function (data) {
						var splitData	 = data.split("|||||");
						var result		 = splitData[0];
						var cont		 = splitData[1];


						if (result == "OK") {
							$("#NowBestProductList").html(cont);
							return;
						}
						else {
							alert(cont);
							return;
						}
		},
		error		 : function (data) {
						alert("처리 도중 오류가 발생하였습니다.");
		}
	});
}

/* 메인페이지 STYLE PEOPLE */
function get_StylePeopleProductList(idx) {
	$.ajax({
		type		 : "post",
		url			 : "/Common/Ajax/Main_StylePeopleProductList.asp",
		async		 : false,
		data		 : "Idx=" + idx,
		dataType	 : "text",
		success		 : function (data) {
						var splitData	 = data.split("|||||");
						var result		 = splitData[0];
						var cont		 = splitData[1];


						if (result == "OK") {
							$("#StylePeopleProductList").html(cont);
							return;
						}
						else {
							alert(cont);
							return;
						}
		},
		error		 : function (data) {
						alert("처리 도중 오류가 발생하였습니다.");
		}
	});
}

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

/* 상품리스트 스마트 검색 초기화 */
function init_SmartSearch() {
	$("input:checkbox[name='SBrandCode']", "form[name='form1']").prop("checked", false);
	$("input:checkbox[name='SSizeCD']", "form[name='form1']").prop("checked", false);
	$(".RangeBar").slider({ values: [0, 100] });
	$("#amount").val("0 ~ 100만");
	$("#SPrice", "form[name='form1']").val("0");
	$("#EPrice", "form[name='form1']").val("100");
	$("input:checkbox[name='SColorCode']", "form[name='form1']").prop("checked", false);
	$("input:checkbox[name='SPickupFlag']", "form[name='form1']").prop("checked", false);
	$("input:checkbox[name='SFreeFlag']", "form[name='form1']").prop("checked", false);
	$("input:checkbox[name='SReserveFlag']", "form[name='form1']").prop("checked", false);
	$("#SPickupFlagChecked").removeClass("is-checked");
	$("#SFreeFlagChecked").removeClass("is-checked");
	$("#SReserveFlagChecked").removeClass("is-checked");
}

/* 브랜드 스마트 검색 초기화 */
function init_BrandSmartSearch() {
	$("input:checkbox[name='SCode1']", "form[name='form1']").prop("checked", false);
	$("input:checkbox[name='SSizeCD']", "form[name='form1']").prop("checked", false);
	$(".RangeBar").slider({ values: [0, 100] });
	$("#amount").val("0 ~ 100만");
	$("#SPrice", "form[name='form1']").val("0");
	$("#EPrice", "form[name='form1']").val("100");
	$("input:checkbox[name='SColorCode']", "form[name='form1']").prop("checked", false);
	$("input:checkbox[name='SPickupFlag']", "form[name='form1']").prop("checked", false);
	$("input:checkbox[name='SFreeFlag']", "form[name='form1']").prop("checked", false);
	$("input:checkbox[name='SReserveFlag']", "form[name='form1']").prop("checked", false);
	$("#SPickupFlagChecked").removeClass("is-checked");
	$("#SFreeFlagChecked").removeClass("is-checked");
	$("#SReserveFlagChecked").removeClass("is-checked");
}

/* 찜처리 */
function set_MyWishList(productCode, onFlag) {
	var url = "";
	if (onFlag == "Y") {
		url = "/ASP/Mypage/Ajax/Product_Pick_Delete.asp";
	}
	else {
		url = "/ASP/Mypage/Ajax/Product_Pick_Insert.asp"
	}

	var retVal = "";
	$.ajax({
		type: "post",
		url: url,
		async: false,
		data: "ProductCode=" + productCode,
		dataType: "text",
		success: function (data) {
			retVal = data;
		},
		error: function (data) {
			retVal = "FAIL|||||처리 도중 오류가 발생하였습니다.";
		}
	});
	return retVal;
}

/* 브랜드 찜처리 */
function set_MyBrandPick(brandCode, onFlag) {
	var url = "";
	if (onFlag == "Y") {
		url = "/ASP/Mypage/Ajax/Brand_Pick_Delete.asp";
	}
	else {
		url = "/ASP/Mypage/Ajax/Brand_Pick_Insert.asp"
	}

	var retVal = "";
	$.ajax({
		type: "post",
		url: url,
		async: false,
		data: "BrandCode=" + brandCode,
		dataType: "text",
		success: function (data) {
			retVal = data;
		},
		error: function (data) {
			retVal = "FAIL|||||처리 도중 오류가 발생하였습니다.";
		}
	});
	return retVal;
}


//달력
jQuery(function ($) {
	$.datepicker.regional['ko'] = {
		closeText: '닫기',
		prevText: '이전달',
		nextText: '다음달',
		currentText: '오늘',
		monthNames: ['1월(JAN)', '2월(FEB)', '3월(MAR)', '4월(APR)', '5월(MAY)', '6월(JUN)',
			'7월(JUL)', '8월(AUG)', '9월(SEP)', '10월(OCT)', '11월(NOV)', '12월(DEC)'],
		monthNamesShort: ['1월', '2월', '3월', '4월', '5월', '6월',
			'7월', '8월', '9월', '10월', '11월', '12월'],
		dayNames: ['일', '월', '화', '수', '목', '금', '토'],
		dayNamesShort: ['일', '월', '화', '수', '목', '금', '토'],
		dayNamesMin: ['일', '월', '화', '수', '목', '금', '토'],
		weekHeader: 'Wk',
		dateFormat: 'yy-mm-dd',
		firstDay: 0,
		isRTL: false,
		showMonthAfterYear: true,
		yearSuffix: '년',
		changeYear: true,	/* 년 선택박스 사용 */
		changeMonth: true,	/* 월 선택박스 사용 */
		showOtherMonths: true,    /* 이전/다음 달 일수 보이기 */
		selectOtherMonths: true    /* 이전/다음 달 일 선택하기 */
	};
	$.datepicker.setDefaults($.datepicker.regional['ko']);
});




/* 쿠폰다운로드 */
function coupon_ProductCoupon(cIdx) {
	$.ajax({
		type		 : "post",
		url			 : "/ASP/Product/Ajax/CouponDownLoad.asp",
		async		 : false,
		data		 : "Idx=" + cIdx,
		dataType	 : "text",
		success		 : function (data) {
						var splitData	 = data.split("|||||");
						var result		 = splitData[0];
						var cont		 = splitData[1];

						if (result == "OK") {
							openAlertLayer("alert", "쿠폰이 지급 되었습니다.<br />마이페이지 MY쿠폰에서 확인 가능합니다.", "closePop('alertPop', '')", "");
							return;
						}
						else if (result == "LOGIN") {
							openAlertLayer("confirm", "로그인 후 이용 가능합니다.<br />로그인 하시겠습니까?", "closePop('confirmPop', '');", "closePop('confirmPop', '');login_Top();");
							return;
						}
						else {
							openAlertLayer("alert", cont, "closePop('alertPop', '')", "");
							return;
						}
		},
		error		 : function (data) {
						openAlertLayer("alert", "처리 도중 오류가 발생하였습니다.", "closePop('alertPop', '')", "");
		}
	});
}

//대분류 보여주기
function GetCategory1() {

	close_ProductLatest();

	var _this = $(".top-exposed");
	_this.closest('body').addClass('posFixed');

	var _windowHeight = $(window).height();
	var _popHeight = _this.height();
	var _contentHeight = _this.find('.contents').outerHeight();
	var _btnHeight = _this.find('.btns').height();
	var _maxHeight = _windowHeight - _btnHeight;

	$.ajax({
		type: "post",
		url: "/Common/Ajax/Get_Category1List.asp",
		async: false,
		data: "",
		dataType: "text",
		success: function (data) {

			$("#Category1Cont").html(data);

			_this.addClass('vertical');
			_this.css('height', _windowHeight - 82 + 'px');
			_this.closest('.wrap-pop').removeClass('hidden');

			$("#Category1PopView").show();

			var categoryAccodion = function () {
				var selector,
					module;

				selector = {
					parent: '.accord-mypage',
					button: '.clickAct',
					toggler: '.ly-title',
					panel: '.ly-content',
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

		},
		error: function (data) {
			alert(data.responseText);
			common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
		}
	});
}

function GetSubCategory(sCode1, sCode2, sCode3) {
	$.ajax({
		type: "post",
		url: "/Common/Ajax/Get_Category1List.asp",
		async: false,
		data: "SCode1=" + sCode1 + "&SCode2=" + sCode2 + "&SCode3=" + sCode3,
		dataType: "text",
		success: function (data) {
			$("#Category1Cont").html(data);
		},
		error: function (data) {
			alert(data.responseText);
			common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
		}
	});
}

function GetCategory1Close() {
	var _this = $(".top-exposed");

	_this.removeClass('vertical');
	_this.closest('body').removeClass('posFixed');
	_this.closest('.wrap-pop').addClass('hidden');

	$("#Category1PopView").hide();
}

//최근상품보기
function open_ProductLatest() {

	GetCategory1Close();

	var _this = $("#ProductLatest");
	_this.closest('body').addClass('posFixed');

	var _windowHeight = $(window).height();
	var _popHeight = _this.height();
	var _contentHeight = _this.find('.contents').outerHeight();
	var _btnHeight = _this.find('.btns').height();
	var _maxHeight = _windowHeight - _btnHeight;

	$.ajax({
		type: "post",
		url: "/Common/Ajax/Footer_Product_Latest.asp",
		async: false,
		data: "",
		dataType: "text",
		success: function (data) {

			$("#ProductLatestCont").html(data);

			_this.addClass('vertical');
			_this.css('height', _windowHeight - 82 + 'px');
			_this.closest('.wrap-pop').removeClass('hidden');

			$("#ProductLatestView").show();
		},
		error: function (data) {
			alert(data.responseText);
			common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
		}
	});
}

function close_ProductLatest() {
	var _this = $("#ProductLatest");

	_this.removeClass('vertical');
	_this.closest('body').removeClass('posFixed');
	_this.closest('.wrap-pop').addClass('hidden');

	$("#ProductLatestView").hide();
}


//타임세일
function open_TimeSale() {
	$.ajax({
		type: "post",
		url: "/ASP/Product/Ajax/TimeSale.asp",
		async: false,
		data: "",
		dataType: "text",
		success: function (data) {
			arrData = data.split("|||||");
			Result = arrData[0];
			Data = arrData[1];

			if (Result == "OK") {
				$("#TimeSaleContent").html(Data);
				$("#TiemSalePop").show();
			}
			else {
				common_msgPopOpen("SHOEMARKER", "진행중인 타임세일이 없습니다.", "", "msgPopup", "N");
			}
		},
		error: function (data) {
			alert(data.responseText);
			common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
		}
	});
}

function close_TimeSale() {
	$("#TiemSalePop").hide();
}



/* 1:1문의 입력폼 */
function popMtmQnaAdd() {
	common_PopOpen('DimDepth1', 'MyMtmQnaWrite');
}

/* 1:1문의 입력폼 다중 selectbox */
var arrCate1_1 = ["문의", "칭찬", "불만", "제안", "기타"];															//온라인

var arrCate1_1_1 = ["상품문의", "주문/결제", "배송관련", "교환/환불", "A/S관련", "적립금", "회원정보관련", "기타"];	//온라인 > 문의
var arrCate1_1_1_1 = ["재고문의", "사이즈문의", "기타"];															//온라인 > 문의 > 상품문의
var arrCate1_1_1_3 = ["출고문의", "배송지연", "오배송", "택배예약", "기타"]											//온라인 > 문의 > 배송관련
var arrCate1_1_1_4 = ["접수문의", "진행상황문의", "기타"];															//온라인 > 문의 > 교환/환불
var arrCate1_1_1_5 = ["접수문의", "수선품질", "수선응대시 불만", "수선기간 불만", "수선비용 불만", "기타"];			//온라인 > 문의 > A/S관련
var arrCate1_1_1_6 = ["사용문의", "오류문의", "기타"];																//온라인 > 문의 > 적립금

var arrCate1_1_2 = ["직원칭찬", "서비스이용", "기타"];																//온라인 > 칭찬
var arrCate1_1_3 = ["직원불친절", "프로모션", "교환/환불", "답변불만", "배송불만", "기타"];							//온라인 > 불만
var arrCate1_1_4 = ["제휴관련"];																					//온라인 > 제안


function selCate2() {
	var cate1 = $("select[name=category1]", "form[name=MtmQnaWriteForm]").val();
	if (cate1 == "온라인") {
		$("#span_cate2").show();
		var cate2Value = "<option value=''>온라인 상담유형을 선택하세요.</option>";
		for (i = 0; i < arrCate1_1.length; i++) {
			cate2Value += "<option value='" + arrCate1_1[i] + "'>" + arrCate1_1[i] + "</option>";
		}
		$("select[name=category2]", "form[name=MtmQnaWriteForm]").html(cate2Value);
	} else {
		$("select[name=category2]", "form[name=MtmQnaWriteForm]").html("");
		$("select[name=category3]", "form[name=MtmQnaWriteForm]").html("");
		$("select[name=category4]", "form[name=MtmQnaWriteForm]").html("");
		$("#span_cate2").hide();
		$("#span_cate3").hide();
		$("#span_cate4").hide();
	}
	$("#selval2").html("온라인 상담유형을 선택하세요.");
	$("#selval3").html("상세 유형을 선택하세요.");
	$("#selval4").html("구분 유형을 선택하세요.");
}

function selCate3() {
	var cate2 = $("select[name=category2]", "form[name=MtmQnaWriteForm]").val();
	if (cate2 == "문의") {
		$("#span_cate3").show();
		var cate3Value = "<option value=''>상세 유형을 선택하세요.</option>";
		for (i = 0; i < arrCate1_1_1.length; i++) {
			cate3Value += "<option value='" + arrCate1_1_1[i] + "'>" + arrCate1_1_1[i] + "</option>";
		}
		$("select[name=category3]", "form[name=MtmQnaWriteForm]").html(cate3Value);
	} else if (cate2 == "칭찬") {
		$("#span_cate3").show();
		var cate3Value = "<option value=''>상세 유형을 선택하세요.</option>";
		for (i = 0; i < arrCate1_1_2.length; i++) {
			cate3Value += "<option value='" + arrCate1_1_2[i] + "'>" + arrCate1_1_2[i] + "</option>";
		}
		$("select[name=category3]", "form[name=MtmQnaWriteForm]").html(cate3Value);
	} else if (cate2 == "불만") {
		$("#span_cate3").show();
		var cate3Value = "<option value=''>상세 유형을 선택하세요.</option>";
		for (i = 0; i < arrCate1_1_3.length; i++) {
			cate3Value += "<option value='" + arrCate1_1_3[i] + "'>" + arrCate1_1_3[i] + "</option>";
		}
		$("select[name=category3]", "form[name=MtmQnaWriteForm]").html(cate3Value);
	} else if (cate2 == "제안") {
		$("#span_cate3").show();
		var cate3Value = "<option value=''>상세 유형을 선택하세요.</option>";
		for (i = 0; i < arrCate1_1_4.length; i++) {
			cate3Value += "<option value='" + arrCate1_1_4[i] + "'>" + arrCate1_1_4[i] + "</option>";
		}
		$("select[name=category3]", "form[name=MtmQnaWriteForm]").html(cate3Value);

	} else {
		$("select[name=category3]", "form[name=MtmQnaWriteForm]").html("");
		$("select[name=category4]", "form[name=MtmQnaWriteForm]").html("");
		$("#span_cate3").hide();
		$("#span_cate4").hide();
	}
	$("#selval3").html("상세 유형을 선택하세요.");
	$("#selval4").html("구분 유형을 선택하세요.");
}

function selCate4() {
	var cate3 = $("select[name=category3]", "form[name=MtmQnaWriteForm]").val();
	if (cate3 == "상품문의") {
		$("#span_cate4").show();
		var cate4Value = "<option value=''>구분 유형을 선택하세요.</option>";
		for (i = 0; i < arrCate1_1_1_1.length; i++) {
			cate4Value += "<option value='" + arrCate1_1_1_1[i] + "'>" + arrCate1_1_1_1[i] + "</option>";
		}
		$("select[name=category4]", "form[name=MtmQnaWriteForm]").html(cate4Value);
		$("select[name=category4]", "form[name=MtmQnaWriteForm]").val("");
	} else if (cate3 == "배송관련") {
		$("#span_cate4").show();
		var cate4Value = "<option value=''>구분 유형을 선택하세요.</option>";
		for (i = 0; i < arrCate1_1_1_3.length; i++) {
			cate4Value += "<option value='" + arrCate1_1_1_3[i] + "'>" + arrCate1_1_1_3[i] + "</option>";
		}
		$("select[name=category4]", "form[name=MtmQnaWriteForm]").html(cate4Value);
	} else if (cate3 == "교환/환불") {
		$("#span_cate4").show();
		var cate4Value = "<option value=''>구분 유형을 선택하세요.</option>";
		for (i = 0; i < arrCate1_1_1_4.length; i++) {
			cate4Value += "<option value='" + arrCate1_1_1_4[i] + "'>" + arrCate1_1_1_4[i] + "</option>";
		}
		$("select[name=category4]", "form[name=MtmQnaWriteForm]").html(cate4Value);
	} else if (cate3 == "A/S관련") {
		$("#span_cate4").show();
		var cate4Value = "<option value=''>구분 유형을 선택하세요.</option>";
		for (i = 0; i < arrCate1_1_1_5.length; i++) {
			cate4Value += "<option value='" + arrCate1_1_1_5[i] + "'>" + arrCate1_1_1_5[i] + "</option>";
		}
		$("select[name=category4]", "form[name=MtmQnaWriteForm]").html(cate4Value);
	} else if (cate3 == "적립금") {
		$("#span_cate4").show();
		var cate4Value = "<option value=''>구분 유형을 선택하세요.</option>";
		for (i = 0; i < arrCate1_1_1_6.length; i++) {
			cate4Value += "<option value='" + arrCate1_1_1_6[i] + "'>" + arrCate1_1_1_6[i] + "</option>";
		}
		$("select[name=category4]", "form[name=MtmQnaWriteForm]").html(cate4Value);
	} else {
		$("select[name=category4]", "form[name=MtmQnaWriteForm]").html("");
		$("#selval4").html("구분 유형을 선택하세요.");
		$("#span_cate4").hide();
	}
	$("#selval4").html("구분 유형을 선택하세요.");
}

/* 1:1문의 이미지 선택창 열기 */
function openMtmQnaImageSearch() {
	var ImgCount = parseInt($("form[name='MtmQnaWriteForm'] input[name='UploadFilesCount']").val());
	if (ImgCount >= 5) {
		common_msgPopOpen("", "첨부 이미지는 5개까지만 가능합니다.");
	}
	else {
		$("form[name='MtmQnaWriteForm'] input[name='FileName']").trigger('click');
	}
}

/* 1:1문의 이미지 추가 */
function mtmQnaImageAdd() {
	var ImgCount = parseInt($("form[name='MtmQnaWriteForm'] input[name='UploadFilesCount']").val());

	if (ImgCount >= 5) {
		common_msgPopOpen("", "첨부 이미지는 5개까지만 가능합니다.");
		return;
	}

	var img = $("form[name='MtmQnaWriteForm'] input[name='FileName']").val().trim();
	if (img.length > 0) {
		lng = img.length;
		ext = img.substring(lng - 4, lng);
		ext = ext.toLowerCase();
		if (!(ext == ".jpg" || ext == ".gif" || ext == ".png" || ext == "jpeg")) {
			common_msgPopOpen("", "이미지는 gif, jpg, png, jpeg만 업도르 가능합니다.");
			return;
		}

		var formData = new FormData($("form[name='MtmQnaWriteForm']")[0]);
		$.ajax({
			type: "post",
			url: "/ASP/Mypage/Ajax/MtmQnaImageTempUpload.asp",
			data: formData,
			async: false,
			contentType: false,
			cache: false,
			processData: false,
			dataType: "text",
			success: function (data) {
				var splitData = data.split("|||||");
				var result = splitData[0];
				var cont = splitData[1];

				if (result == "OK") {
					var splitData2 = cont.split("^^^^^");
					var imagePath = splitData2[0];
					var imageName = splitData2[1];
					mtmQnaImagePreView(imagePath, imageName);
				}
				else {
					common_msgPopOpen("", cont);
					return;
				}
			},
			error: function (data) {
				//alert(data.responseText);
				common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.");
			}
		});
	}
}

/* 1:1문의 선택 이미지 미리보기 재배치 */
function reMtmQnaImagePreView(imagePath) {
	$("form[name='MtmQnaWriteForm'] .review-photo").html("");
	var i = parseInt($("form[name='MtmQnaWriteForm'] input[name='UploadFilesCount']").val());
	var UploadFiles = $("form[name='MtmQnaWriteForm'] input[name='UploadFiles']").val();

	if (i == 1) {
		var html = "<li class=\"photo-list\">";
		html = html + "<button type=\"button\" onclick=\"mtmQnaImageDelete(" + i + ")\">삭제</button>";
		html = html + "<div class=\"img\">";
		html = html + "<img src=\"" + imagePath + UploadFiles + "\" alt=\"후기 이미지\">";
		html = html + "</div>";
		html = html + "</div>";
		html = html + "</li>";
		$("form[name='MtmQnaWriteForm'] .review-photo").append(html);
	} else {
		var UploadFilesArr = UploadFiles.split("|||||");
		for (var k = 1; k <= i; k++) {
			var html = "<li class=\"photo-list\">";
			html = html + "<button type=\"button\" onclick=\"mtmQnaImageDelete(" + k + ")\">삭제</button>";
			html = html + "<div class=\"img\">";
			html = html + "<img src=\"" + imagePath + UploadFilesArr[k - 1] + "\" alt=\"후기 이미지\">";
			html = html + "</div>";
			html = html + "</div>";
			html = html + "</li>";
			$("form[name='MtmQnaWriteForm'] .review-photo").append(html);
		}
	}

}

/* 1:1문의 선택 이미지 미리보기 */
function mtmQnaImagePreView(imagePath, imageName) {
	var i = parseInt($("form[name='MtmQnaWriteForm'] input[name='UploadFilesCount']").val()) + 1;
	var html = "<li class=\"photo-list\">";
	html = html + "<button type=\"button\" onclick=\"mtmQnaImageDelete(" + i + ")\">삭제</button>";
	html = html + "<div class=\"img\">";
	html = html + "<img src=\"" + imagePath + imageName + "\" alt=\"후기 이미지\">";
	html = html + "</div>";
	html = html + "</div>";
	html = html + "</li>";

	$("form[name='MtmQnaWriteForm'] .review-photo").append(html);

	var uploadFiles = $("form[name='MtmQnaWriteForm'] input[name='UploadFiles']").val().trim();

	if (uploadFiles == "") {
		uploadFiles = imageName;
	}
	else {
		uploadFiles = uploadFiles + "|||||" + imageName;
	}

	$("form[name='MtmQnaWriteForm'] input[name='UploadFiles']").val(uploadFiles);
	$("form[name='MtmQnaWriteForm'] input[name='UploadFilesCount']").val(i);
}

/* 1:1문의 이미지 삭제 */
function mtmQnaImageDelete(index) {
	// 삭제할 임시파일 경로
	var delFileName = $("form[name='MtmQnaWriteForm'] .photo-list .img").eq(index - 1).find("img").attr("src");
	var splitFileName = delFileName.split("/");
	var filepath = delFileName.replace(splitFileName[splitFileName.length - 1], "");
	delFileName = splitFileName[splitFileName.length - 1];

	// 임시 이미지 삭제처리
	$.ajax({
		type: "post",
		url: "/ASP/Mypage/Ajax/MtmQnaImageTempDelete.asp",
		async: true,
		data: "FileName=" + delFileName,
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				// 미리보기 이미지 삭제
				$("form[name='MtmQnaWriteForm'] .photo-list .img").eq(index - 1).remove();

				// 업로드할 이미지 리스트 작성
				var uploadFiles = "";
				$("form[name='MtmQnaWriteForm'] .photo-list .img").each(function () {
					var imageUrl = $(this).find("img").attr("src");
					var splitImageUrl = imageUrl.split("/");
					if (uploadFiles == "") {
						uploadFiles = splitImageUrl[splitImageUrl.length - 1];
					}
					else {
						uploadFiles = uploadFiles + "|||||" + splitImageUrl[splitImageUrl.length - 1];
					}
				});
				$("form[name='MtmQnaWriteForm'] input[name='UploadFiles']").val(uploadFiles);

				// 업로드할 이미지 수
				var i = parseInt($("form[name='MtmQnaWriteForm'] input[name='UploadFilesCount']").val()) - 1;
				$("form[name='MtmQnaWriteForm'] input[name='UploadFilesCount']").val(i);

				reMtmQnaImagePreView(filepath);
			}
		},
		error: function (data) {
			//alert(data.responseText);
			common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.");
		}
	});
}

/* 1:1문의 등록 */
function chk_MtmWrite() {
	var cate1 = $("form[name='MtmQnaWriteForm'] select[name='category1']").val();
	var cate2 = $("form[name='MtmQnaWriteForm'] select[name='category2']").val();
	var cate3 = $("form[name='MtmQnaWriteForm'] select[name='category3']").val();
	var cate4 = $("form[name='MtmQnaWriteForm'] select[name='category4']").val();

	if (cate1 == "") {
		common_msgPopOpen("", "문의 유형을 선택하세요.");
		return;
	}
	if (cate2 == "" && cate2 != null) {
		common_msgPopOpen("", "온라인상담 유형을 선택하세요.");
		return;
	}
	if (cate3 == "" && cate3 != null) {
		common_msgPopOpen("", "상세 유형을 선택하세요.");
		return;
	}
	if (cate4 == "" && cate4 != null) {
		common_msgPopOpen("", "구분 유형을 선택하세요.");
		return;
	}

	var title = alltrim($("form[name='MtmQnaWriteForm'] input[name='Title']").val());
	if (title == "") {
		common_msgPopOpen("", "제목을 입력하세요.", "", "Title");
		return;
	}
	var contents = alltrim($("form[name='MtmQnaWriteForm'] textarea[name='Contents']").val());
	if (contents == "") {
		common_msgPopOpen("", "내용을 입력하세요.");
		$("form[name='MtmQnaWriteForm'] textarea[name='Contents']").focus();
		return;
	}

	var smsReturnFlag = $("form[name='MtmQnaWriteForm'] input[name='SMSReturnFlag']:checked").val();
	if (smsReturnFlag == "1") {
		var hp1 = $("form[name='MtmQnaWriteForm'] select[name='Mobile1']").val();
		var hp23 = alltrim($("form[name='MtmQnaWriteForm'] input[name='Mobile23']").val());
		if (hp1.length == 0) {
			common_msgPopOpen("", "휴대폰번호를 선택해 주십시오.", "", "Mobile1");
			return;
		}
		if (hp23.length < 7) {
			common_msgPopOpen("", "올바른 휴대폰번호를 입력해 주십시오.", "", "Mobile23");
			return;
		}
		if (only_Num(hp23) == false) {
			common_msgPopOpen("", "휴대폰번호를 숫자로만 입력해 주십시오.", "", "Mobile23");
			return;
		}
	}

	var eMailReturnFlag = $("form[name='MtmQnaWriteForm'] input[name='EMailReturnFlag']:checked").val();
	if (eMailReturnFlag == "1") {
		var email = alltrim($("input[name='EMail']").val());
		if (email.length == 0) {
			common_msgPopOpen("", "이메일 계정을 입력해 주십시오.", "", "Email");
			return;
		}
		if (beAllowStr(email, "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890.-_@") == false) {
			common_msgPopOpen("", "이메일 주소를 영문과 숫자로만 입력해 주십시오.", "", "Email");
			return;
		}
		if (email.indexOf(".") < 0 || email.indexOf(".") == 0 || email.indexOf(".") == email.length - 1 || email.indexOf("@") < 0 || email.indexOf("@") == 0 || email.indexOf("@") == email.length - 1) {
			common_msgPopOpen("", "올바른 이메일 주소를 다시 입력해 주십시오.", "", "Email");
			return;
		}
	}

	$.ajax({
		type: "post",
		url: "/ASP/MyPage/Ajax/MtmQnaWriteOK.asp",
		async: false,
		data: $("form[name='MtmQnaWriteForm']").serialize(),
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				common_msgPopOpen("", cont,"location.href='/ASP/Mypage/Qna.asp?QnaType=2'");
				common_PopClose('DimDepth1');
				return;
			}
			else if (result == "LOGIN") {
				common_msgPopOpen("", cont, "PageReload();");
				return;
			}
			else {
				common_msgPopOpen("", cont);
				return;
			}
		},
		error: function (data) {
			//alert(data.responseText);
			common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.");
		}
	});
}

/* 공지사항 뷰 */
function footerNoticeView(idx) {
	$("form[name=FooterNoticeViewForm] input[name=Idx]").val(idx);
	common_PopOpen('DimDepth1', 'FooterNoticeView');
}

/*공지사항 롤링*/
function footerNoticeRolling() {
	$.ajax({
		type: "post",
		url: "/Common/Ajax/Footer_NoticeList_Top5.asp",
		async: true,
		data: "",
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				$("#FooterNoticeList").html(cont);
				fn_NoticeTop5('notification', '', true);
				return;
			}
		},
		error: function (data) {
			//alert(data.responseText);
			common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.");
		}
	});
}

function fn_NoticeTop5(containerID, buttonID, autoStart) {
	var $element = $('#' + containerID).find('.text');
	var autoPlay = autoStart; //자동으로 돌아가는 설정
	var auto = null;
	var speed = 3000; //자동으로 돌아가는 설정 시간
	var timer = null;
	var move = $element.children().outerHeight();//이것에 따라 top값 변경, .notice-list의 li높이
	var first = false;
	var lastChild;

	lastChild = $element.children().eq(-1).clone(true);
	lastChild.prependTo($element);
	$element.children().eq(-1).remove();

	if ($element.children().length == 1) {
		$element.css('top', '0px');
	} else {
		$element.css('top', '-' + move + 'px');
	}

	if (autoPlay) {
		timer = setInterval(moveNextSlide, speed);
	}

	$element.find('>p').bind({
		'mouseenter': function () {
			if (auto) {
				clearInterval(timer);
			}
		},
		'mouseleave': function () {
			if (auto) {
				timer = setInterval(moveNextSlide, speed);
			}
		}
	});


	function moveNextSlide() {
		$element.each(function (idx) {
			var firstChild = $element.children().filter(':first-child').clone(true);
			firstChild.appendTo($element.eq(idx));
			$element.children().filter(':first-child').remove();
			$element.css('top', '0px');
			$element.eq(idx).animate({ 'top': '-' + move + 'px' }, 'normal');
		});
	}
}

/* 상품 검색 */
function TopSearch() {
	$("#TopSearch").show();
	TopSearchWordView('P');
}

function TopSearchWordView(t) {
	if (t == 'P') {
		$("#ts1").addClass("active");
		$("#ts2").removeClass("active");
	} else {
		$("#ts1").removeClass("active");
		$("#ts2").addClass("active");
	}

	$.ajax({
		type: "post",
		url: "/Common/Ajax/TopSearch.asp",
		async: false,
		data: "vType=" + t,
		dataType: "text",
		success: function (data) {
			$("#WordView").html(data);
		},
		error: function (data) {
			alert(data.responseText);
			common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
		}
	});

}

function TopSearchClose() {
	$("#TopSearch").hide();
}

function TopSearchGo() {
	APP_GoUrl('/ASP/Product/SearchProductList.asp?SearchWord=' + $("#SearchText").val());
	$("#TopSearch").hide();

	//document.TopSearchForm.SearchWord.value = $("#SearchText").val();
	//document.TopSearchForm.submit();
}

//사이즈 레이어
function SizeLayerOpen(productcode) {
	$.ajax({
		url: '/ASP/Product/Ajax/ProductSizeList.asp',
		data: "ProductCode=" + productcode,
		async: false,
		type: 'get',
		dataType: 'html',
		success: function (data, textStatus, jqXHR) {

			$("#sizelist").html(data);
			$("#SizePop").show();
		},
		error: function (data, textStatus, jqXHR) {
			//alert(jqXHR);
			//alert(data.responseText);
			alert("리스트를 가져오는 도중 오류가 발생하였습니다.");
		}
	});
}

/* 페이지 이동 */
function move_Page(url) {
	if (url == "/") {
		APP_TopGoUrl("/");
	}
	if (url.toLowerCase() == "/asp/member/login.asp") {
		APP_TopGoUrl("/ASP/Member/Login.asp");
	}
	else {
		//openPop("loading");
		location.href = url;
	}
}

/* 브랜드 검색 */
function TopBrandSearch() {
	var TopBrandSearchWord = $("#TopBrandSearchWord").val().trim();
	if (TopBrandSearchWord.length <= 0) {
		BrandErrPop("찾으시는 브랜드명을 입력하여 주세요.");
		return;
	}

	$.ajax({
		url: '/Common/Ajax/Get_TopSearchBrandList.asp',
		data: "TopBrandSearchWord=" + TopBrandSearchWord,
		async: false,
		type: 'get',
		dataType: 'html',
		success: function (data, textStatus, jqXHR) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				$("#BrandSearchView").html(cont);
				$("#BrandSearchView").show();
			} else {
				BrandErrPop(cont);
				return;
			}
		},
		error: function (data, textStatus, jqXHR) {
			//alert(jqXHR);
			//alert(data.responseText);
			BrandErrPop(data.responseText);
		}
	});
}

/* 브랜드 찾기 검색어 없음 레이어 */
function BrandErrPop(cont) {
	$("#msg").html(cont);
	$("#BrandErrPop").show();
}

function BrandErrPopclose(cont) {
	$("#BrandErrPop").hide();
}

/* 쿠폰다운로드 */
function cp_down(cIdx) {
	$.ajax({
		type		 : "post",
		url			 : "/ASP/Mypage/Ajax/CouponDownLoad.asp",
		async		 : false,
		data		 : "Idx=" + cIdx,
		dataType	 : "text",
		success		 : function (data) {
						var splitData	 = data.split("|||||");
						var result		 = splitData[0];
						var cont		 = splitData[1];


						if (result == "OK") {
							openAlertLayer("alert", "쿠폰이 지급 되었습니다.<br />마이페이지 쿠폰북에서 확인 가능합니다.", "closePop('alertPop', '')", "");
							return;
						}
						else if (result == "LOGIN") {
							openAlertLayer("alert", "로그인 후 이용 가능합니다.", "closePop('alertPop', '');", "");
							return;
						}
						else {
							openAlertLayer("alert", cont, "closePop('alertPop', '');", "");
							return;
						}
		},
		error		 : function (data) {
						openAlertLayer("alert", "처리 도중 오류가 발생하였습니다.", "closePop('alertPop', '')", "");
		}
	});
}