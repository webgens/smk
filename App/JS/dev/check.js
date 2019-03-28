function PageReload() {
	location.href = document.URL;
}

/* OPEN LAYER POPUP */
function openPop(id) {
	//$(".dim").show();

	var pop = $("#" + id);
	var pop_height = pop.height();


	var popMargLeft = ($("#" + id).width() + 2) / 2;
	$("#" + id).css({ 'margin-left': -popMargLeft });

	var top = ($('.dim').height() - $("#" + id).height()) / 2;
	//$("#" + id).fadeIn().css({ 'top': top });
	$("#" + id).show().css({ 'top': top });

	lock();

	if (id == "alertPop") {
		$("#alert_confirm").focus();
	}
	else if (id == "confirmPop") {
		$("#confirm_cancel").focus();
	}
	else if (id == "messagePop") {
		$("#message_btn2").focus();
	}
}

/* CLOSE LAYER POPUP */
function closePop(id, focusId) {
	//$(".dim").hide();
	$("#" + id).hide();
	release();

	if (focusId != undefined && focusId != "") {
		$("#" + focusId).focus();
	}
};

/* BODY SCROLLBAR LOCK */
function lock() {
	$("body").css("overflow", "hidden");
}
/* BODY SCROLLBAR RELEASE */
function release() {
	$("body").css("overflow", "auto");
};

/* ALERT LAYER POPUP */
function openAlertLayer(type, msg, link1, link2) {
	var id = "";
	if (type == "confirm") {
		id = "confirmPop";
		$("#confirm_title").html("SHOEMARKER");
		$("#confirm_content").html(msg);
		$("#confirm_cancel").attr("onclick", link1);
		$("#confirm_confirm").attr("onclick", link2);
	}
	else {
		id = "alertPop";
		$("#alert_title").html("SHOEMARKER");
		$("#alert_content").html(msg);
		$("#alert_confirm").attr("onclick", link1);
	}

	openPop(id);
}
function openAlertLayer2(type, msg, btn1, btn2, link1, link2) {
	var id = "";
	if (type == "confirm") {
		id = "confirmPop";
		$("#confirm_title").html("SHOEMARKER");
		$("#confirm_content").html(msg);
		$("#confirm_cancel").html(btn1);
		$("#confirm_confirm").html(btn2);
		$("#confirm_cancel").attr("onclick", link1);
		$("#confirm_confirm").attr("onclick", link2);
	}
	else {
		id = "alertPop";
		$("#alert_title").html("SHOEMARKER");
		$("#alert_content").html(msg);
		$("#alert_confirm").html(btn1);
		$("#alert_confirm").attr("onclick", link1);
	}

	openPop(id);
}



function alltrim(str) {
	var i;
	var ch;
	var retStr = '';
	var retStr1 = '';
	if (str.length == 0)
		return str;
	for (i = 0; i < str.length; i++) {
		ch = str.charAt(i);
		if (ch == ' ' || ch == '\r' || ch == '\n')
			continue;
		retStr += ch;
	}
	return retStr;
}

function beAllowStr(str, allowStr) {
	var i;
	var ch;
	for (i = 0; i < str.length; i++) {
		ch = str.charAt(i);
		if (allowStr.indexOf(ch) < 0) {
			return false;
		}
	}
	return true;
}

function only_Num(val) {
	var regAlphaNum = /^[0-9]+$/;
	if (!regAlphaNum.test(val)) {
		return false;
	}
	else {
		return true;
	}
}

function only_Num2(val) {
	var regAlphaNum = /^[0-9.]+$/;
	if (!regAlphaNum.test(val)) {
		return false;
	}
	else {
		return true;
	}
}

function only_Num3(val) {
	var regAlphaNum = /^[0-9-]+$/;
	if (!regAlphaNum.test(val)) {
		return false;
	}
	else {
		return true;
	}
}

function only_AlphaNum(val) {
	var regAlphaNum = /^[A-Za-z0-9]+$/;
	if (!regAlphaNum.test(val)) {
		return false;
	}
	else {
		return true;
	}
}

function strCharByte(chStr) {
	if (chStr.substring(0, 2) == '%u') {
		if (chStr.substring(2, 4) == '00')
			return 1;
		else
			return 2;
	}
	else if (chStr.substring(0, 1) == '%') {
		if (parseInt(chStr.substring(1, 3), 16) > 127)
			return 2;
		else
			return 1;
	}
	else
		return 1;
}

/* =================================================================
 fn_numbersonly()
 숫자만 입력
 ================================================================= */
function fn_GetEvent(e) {
	if (navigator.appName == 'Netscape') {
		keyVal = e.which;
	}
	else {
		keyVal = event.keyCode;
	}
	return keyVal;
}
function fn_numbersonly(evt) {
	var myEvent = window.event ? window.event : evt;
	var isWindowEvent = window.event ? true : false;
	var keyVal = fn_GetEvent(evt);
	var result = false;
	if (myEvent.shiftKey) {
		result = false;
	}
	else {
		if ((keyVal >= 48 && keyVal <= 57) || (keyVal >= 96 && keyVal <= 105) || (keyVal == 8) || (keyVal == 9) || (keyVal == 13) || (keyVal == 46)) {
			result = true;
		}
		else {
			result = false;
		}
	}
	if (!result) {
		if (!isWindowEvent) {
			myEvent.preventDefault();
		}
		else {
			myEvent.returnValue = false;
		}
	}
}

/* =================================================================
 execDaumPostcode(zipCode, address)
 다음 주소 검색 팝업창
 zipCode : 우편번호 input id
 address : 주소 input id
 ================================================================= */
function execDaumPostcode(zipCode, address) {
	new daum.Postcode({
		oncomplete: function (data) {
			// 팝업에서 검색결과 항목을 클릭했을때 실행할 코드를 작성하는 부분.
			//alert(data.addressType + "-" + data.userSelectedType + "\n\n1 : " + data.roadAddress + "\n2 : " + data.autoRoadAddress + "\n\n3 : " + data.jibunAddress + "\n4 : " + data.autoJibunAddress);

			// data.addressType : 주소검색방법 R=도로명검색, J=지번검색
			// data.userSelectedType : 주소선택구분 R=도로명주소선택, J=지번주소선택
			var roadAddress = data.roadAddress;
			var jibunAddress = data.jibunAddress;
			if (data.addressType == "R") {
				if (jibunAddress == "" && data.userSelectedType == "R") {
					jibunAddress = data.autoJibunAddress;
				}
			} else {
				if (roadAddress == "" && data.userSelectedType == "J") {
					roadAddress = data.autoRoadAddress;
				}
			}

			// 도로명 주소의 노출 규칙에 따라 주소를 조합한다.
			// 내려오는 변수가 값이 없는 경우엔 공백('')값을 가지므로, 이를 참고하여 분기 한다.
			var fullRoadAddr = roadAddress; // 도로명 주소 변수
			var extraRoadAddr = ""; // 도로명 조합형 주소 변수

			// 법정동명이 있을 경우 추가한다. (법정리는 제외)
			// 법정동의 경우 마지막 문자가 "동/로/가"로 끝난다.
			if (data.bname !== "" && /[동|로|가]$/g.test(data.bname)) {
				extraRoadAddr += data.bname;
			}
			// 건물명이 있고, 공동주택일 경우 추가한다.
			if (data.buildingName !== "" && data.apartment === "Y") {
				extraRoadAddr += (extraRoadAddr !== "" ? ", " + data.buildingName : data.buildingName);
			}
			// 도로명, 지번 조합형 주소가 있을 경우, 괄호까지 추가한 최종 문자열을 만든다.
			if (extraRoadAddr !== "") {
				extraRoadAddr = " (" + extraRoadAddr + ")";
			}
			// 도로명, 지번 주소의 유무에 따라 해당 조합형 주소를 추가한다.
			if (fullRoadAddr !== "") {
				fullRoadAddr += extraRoadAddr;
			}

			// 우편번호와 주소 정보를 해당 필드에 넣는다.
			document.getElementById(zipCode).value = data.zonecode; //5자리 새우편번호 사용
			document.getElementById(address).value = fullRoadAddr;
			closePop('PopupPostSearch');
		},
		width: '100%',
		height: '100%',
		maxSuggestItems: 5
	}).embed(document.getElementById('PopupPostContents'));
	openPop('PopupPostSearch');
}


function checkEmail(val) {

	if (val.length == 0) { return false; }

	if (beAllowStr(val, "1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz@.-_") == false) { return false; }

	var atCnt = 0;
	var dotCnt = 0;
	for (i = 0; i < val.length; i++) {
		ch = val.charAt(i);
		if (ch == "@") { atCnt++; }
		if (ch == ".") { dotCnt++; }
	}

	if (atCnt != 1 || dotCnt < 1) { return false; }

	var atIndex = 0;
	atIndex = val.indexOf("@");
	if (atIndex <= 0) { return false; }

	return true;
}

function checkDate(v_year, v_month, v_day) {

	var err = 0
	if (v_year.length != 4) err = 1
	if (v_month.length != 1 && v_month.length != 2) err = 1
	if (v_day.length != 1 && v_day.length != 2) err = 1


	r_year = eval(v_year);
	r_month = eval(v_month);
	r_day = eval(v_day);

	if (r_month < 1 || r_month > 12) err = 1
	if (r_day < 1 || r_day > 31) err = 1
	if (r_year < 0) err = 1


	if (r_month == 4 || r_month == 6 || r_month == 9 || r_month == 11) {
		if (r_day == 31) err = 1
	}

	// 2,윤년체크
	if (r_month == 2) {
		var g = parseInt(r_year / 4)

		if (isNaN(g)) {
			err = 1
		}
		if (r_day > 29) err = 1
		if (r_day == 29 && ((r_year / 4) != parseInt(r_year / 4))) err = 1
	}

	if (err == 1) {
		return false
	} else {
		return true;
	}
}

function beNum(ch) {
	return (ch >= '0' && ch <= '9');
}

function beNumStr(str) {
	var i;
	var ch;
	for (i = 0; i < str.length; i++) {
		ch = str.charAt(i);
		if (beNum(ch) == false) {
			return false;
		}
	}
	return true;
}

function chk_SameChr(val, len) {
	var b = "";
	var c = "";
	var j = 0;
	for (var i = 0; i < val.length; i++) {
		var c = val.charAt(i).toLowerCase();
		if (b == "") { b = c;}
		if (b == c) { j = j + 1; }
		else { j = 1; }
		if (j >= len) { break; }
		b = c;
	}
	if (j >= len) {
		return false;
	}
	else {
		return true;
	}
}

function dateSelect(id) {
	$("#" + id).focus();
}


function setDate(term, sid, eid) {
	var sDate, eDate, year, month, day;

	// 시작일자
	sDate = new Date();

	if (term == "") {
		sDate.setDate(sDate.getDate());
	} else if (term == "15d") {
		sDate.setDate(sDate.getDate() - 15);
	} else if (term == "1m") {
		sDate.setMonth(sDate.getMonth() - 1);
	} else if (term == "3m") {
		sDate.setMonth(sDate.getMonth() - 3);
	} else if (term == "6m") {
		sDate.setMonth(sDate.getMonth() - 6);
	} else if (term == "1y") {
		sDate.setFullYear(sDate.getFullYear() - 1);
	}

	year = sDate.getFullYear();
	month = sDate.getMonth() + 1;
	day = sDate.getDate();
	if (month < 10) month = "0" + month;
	if (day < 10) day = "0" + day;
	sDate = year + "-" + month + "-" + day;

	// 종료일자
	eDate = new Date();
	year = eDate.getFullYear();
	month = eDate.getMonth() + 1;
	day = eDate.getDate();
	if (month < 10) month = "0" + month;
	if (day < 10) day = "0" + day;
	eDate = year + "-" + month + "-" + day;

	$("#" + sid).val(sDate);
	$("#" + eid).val(eDate);
}
