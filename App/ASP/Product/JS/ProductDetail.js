/* 로그인 체크 */
function conf_Login() {
	openAlertLayer("confirm", "로그인 후 이용 가능합니다.<br />로그인 하시겠습니까?", "closePop('confirmPop', '');", "closePop('confirmPop', '');move_Login('S');"); /* dev_common.js */
}


/* 상품문의 Pop Up 호출 */
function popup_ProductQna(productCode) {
	$.ajax({
		type		 : "post",
		url			 : "/ASP/Product/Ajax/ProductQnaAdd.asp",
		async		 : false,
		data		 : "ProductCode=" + productCode,
		dataType	 : "text",
		success		 : function (data) {
						var splitData	 = data.split("|||||");
						var result		 = splitData[0];
						var cont		 = splitData[1];

						if (result == "OK") {
							$("#DimDepth1").html(data);
							openPop('DimDepth1');
						}
						else {
							openAlertLayer("alert", cont, "closePop('alertPop', '');", "");
						}
		},
		error		 : function (data) {
						openAlertLayer("alert", "처리 도중 오류가 발생하였습니다.", "closePop('alertPop', '');", "");
		}
	});
}


/* 상품문의 입력 */
function ins_ProductQna() {

	var Title = alltrim(document.ProductQnaForm.Title.value);
	if (Title.length == 0) {
		openAlertLayer("alert", "제목을 입력하여 주십시오.", "closePop('alertPop', 'Title');", "");
		return;
	}

	var Contents = alltrim(document.ProductQnaForm.Contents.value);
	if (Contents.length == 0) {
		openAlertLayer("alert", "내용을 입력하여 주십시오.", "closePop('alertPop', 'Contents');", "");
		return;
	}

	var SMSReturnFlag = document.ProductQnaForm.SMSReturnFlag.value;
	if (SMSReturnFlag == "1") {
		var Mobile1 = alltrim(document.ProductQnaForm.Mobile1.value);
		var Mobile2 = alltrim(document.ProductQnaForm.Mobile2.value);

		if (only_Num(Mobile2) == false) {
			openAlertLayer("alert", "휴대전화번호를 숫자로만 입력하여 주십시오.", "closePop('alertPop', 'Mobile2');", "");
			return;
		}
		if (Mobile1 == "010") {
			if (Mobile2.length != 8) {
				openAlertLayer("alert", "휴대전화번호를 정확히 입력하여 주십시오.", "closePop('alertPop', 'Mobile2');", "");
				return;
			}
		}
		else {
			if (Mobile2.length < 7) {
				openAlertLayer("alert", "휴대전화번호를 정확히 입력하여 주십시오.", "closePop('alertPop', 'Mobile2');", "");
				return;
			}
		}
	}

	var EMailReturnFlag = document.ProductQnaForm.EMailReturnFlag.value;
	if (EMailReturnFlag == "1") {
		var Email = alltrim(document.ProductQnaForm.Email.value);
		if (Email.length == 0) {
			openAlertLayer("alert", "이메일을 입력하여 주십시오.", "closePop('alertPop', 'Email');", "");
			return;
		}
		if (beAllowStr(Email, "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890.-_@") == false) {
			openAlertLayer("alert", "이메일을 영문과 숫자로만 입력하여 주십시오.", "closePop('alertPop', 'Email');", "");
			return;
		}
		if (Email.indexOf("@") < 0 || Email.indexOf("@") == 0) {
			openAlertLayer("alert", "이메일을 정확히 입력하여 주십시오.", "closePop('alertPop', 'Email');", "");
			return;
		}
		var arrEmail	 = Email.split("@");
		var Email1		 = Email.split("@")[0];
		var Email2		 = Email.split("@")[1];

		if (Email1.length == 0 || Email2.length == 0) {
			openAlertLayer("alert", "이메일을 정확히 입력하여 주십시오.", "closePop('alertPop', 'Email');", "");
			return;
		}
		if (Email2.indexOf(".") < 0 || Email2.indexOf(".") == 0 || Email2.indexOf(".") == Email2.length - 1) {
			openAlertLayer("alert", "이메일을 정확히 입력하여 주십시오.", "closePop('alertPop', 'Email');", "");
			return;
		}
	}


	var productCode = $("input[name='ProductCode']", "form[name='ProductQnaForm']").val();

	$.ajax({
		type		 : "post",
		url			 : "/ASP/Product/Ajax/ProductQnaAddOK.asp",
		async		 : false,
		data		 : $("form[name='ProductQnaForm']").serialize(),
		dataType	 : "text",
		success		 : function (data) {
						var splitData	 = data.split("|||||");
						var result		 = splitData[0];
						var cont		 = splitData[1];

						if (result == "OK") {
							openAlertLayer("alert", "상품문의가 등록 되었습니다.", "closePop('alertPop', '');closePop('DimDepth1');", "");
							ProductCounselList(1, productCode);
						}
						else {
							openAlertLayer("alert", cont, "closePop('alertPop', '');", "");
							return;
						}
		},
		error		 : function (data) {
						openAlertLayer("alert", "처리 도중 오류가 발생하였습니다.", "closePop('alertPop', '');", "");
		}
	});
}


/* 상품문의 리스트 */
function list_ProductCounsel(page, productcode) {
	$("#CounselPage").val(page);
	$.ajax({
		url			 : '/ASP/Product/Ajax/ProductCouselList.asp',
		data		 : "ProductCode=" + productcode + "&Page=" + page,
		async		 : false,
		type		 : 'get',
		dataType	 : 'html',
		success		 : function (data) {
						arrData	 = data.split("|||||");
						Data	 = arrData[0];
						RecCnt	 = arrData[1];
						PageCnt	 = arrData[2];

						$("#ProductCounselCount").html(RecCnt);

						$("#productcounsel_more_btn").show();
						if (parseInt(page) >= parseInt(PageCnt)) {
							$("#productcounsel_more_btn").hide();
						}

						// 목록 로딩시키기
						if (page == 1) {
							$("#productcounselList").html(Data);

						} else {
							$("#productcounselList").append(Data);
						}
		},
		error		 : function (data) {
						openAlertLayer("alert", "처리 도중 오류가 발생하였습니다.", "closePop('alertPop', '');", "");
		}
	});
}

/* 상품문의 다음 리스트 */
function list_ProductCounselNext(productcode) {
	var page = $("#CounselPage").val();
	page = parseInt(page) + 1;
	list_ProductCounsel(page, productcode);
}
