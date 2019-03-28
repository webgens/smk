
/**************************************************************************************/
/* 슈즈상품권 START
/**************************************************************************************/
/* 슈즈상품권 리스트 */
function get_ShoesGiftList(page) {
	var sDate = $("#SDate").val();
	var eDate = $("#EDate").val();

	$("#Page").val(page);

	$.ajax({
		type		 : "post",
		url			 : "/ASP/Mypage/Ajax/ShoesGiftList.asp",
		async		 : false,
		data		 : $("#form").serialize(),
		dataType	 : "text",
		success		 : function (data) {
						var splitData	 = data.split("|||||");
						var result		 = splitData[0];
						var cont		 = splitData[1];


						if (result == "OK") {
							$("#ShoesGiftList").html(cont);
							return;
						}
						else if (result == "LOGIN") {
							PageReload();
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

/* 슈즈상품권번호 체크 */
function chk_ShoesGift() {
	var cpno = $("#cpno", "form[name='formAddShoesGift']").val();
	if (cpno.length == 0) {
		openAlertLayer("alert", "슈즈 상품권 번호를 입력하여 주십시오.", "closePop('alertPop', 'cpno')", "");
		return;
	}
	if (only_Num(cpno) == false) {
		openAlertLayer("alert", "슈즈 상품권 번호를 숫자로만 입력하여 주십시오.", "closePop('alertPop', 'cpno')", "");
		return;
	}

	
	$.ajax({
		type		 : "post",
		url			 : "/ASP/Mypage/Ajax/ShoesGiftCheck.asp",
		async		 : false,
		data		 : $("#formAddShoesGift").serialize(),
		dataType	 : "text",
		success		 : function (data) {
						var splitData	 = data.split("|||||");
						var result		 = splitData[0];
						var cont		 = splitData[1];


						if (result == "OK") {
							common_PopOpen('DimDepth1', 'ShoesGiftAdd')
							return;
						}
						else if (result == "LOGIN") {
							PageReload();
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

/* 슈즈 상품권 등록 */
function ins_ShoesGift() {
	var agr = $("input:checkbox[name='Agr']", "form[name='addShoesGift']").is(":checked");
	if (agr == false) {
		openAlertLayer("alert", "슈즈 상품권 전환 규약에 동의 하여 주십시오.", "closePop('alertPop', '')", "");
		return;
	}

	openAlertLayer("confirm", "슈즈 상품권을 등록 하시겠습니까?", "closePop('confirmPop', '')", "closePop('confirmPop', '');ins_ShoesGiftExec();");
}
function ins_ShoesGiftExec() {
	$.ajax({
		type		 : "post",
		url			 : "/ASP/Mypage/Ajax/ShoesGiftAddOk.asp",
		async		 : false,
		dataType	 : "text",
		success		 : function (data) {
						var splitData	 = data.split("|||||");
						var result		 = splitData[0];
						var cont		 = splitData[1];


						if (result == "OK") {
							openAlertLayer("alert", "등록 되었습니다.", "closePop('alertPop', '');PageReload();", "");
							return;
						}
						else if (result == "LOGIN") {
							PageReload();
							return;
						}
						else {
							openAlertLayer("alert", cont, "closePop('alertPop', '');PageReload();", "");
							return;
						}
		},
		error		 : function (data) {
						openAlertLayer("alert", "처리 도중 오류가 발생하였습니다", "closePop('alertPop', '');", "");
		}
	});
}
/**************************************************************************************/
/* 슈즈상품권 END
/**************************************************************************************/


/**************************************************************************************/
/* 쿠폰북 START
/**************************************************************************************/
/* 쿠폰 구분 클릭 */
function clk_CouponType(num) {
	$(".tab-links > .tab-link").removeClass("current");
	$(".tab-links > .tab-link").eq(num).addClass("current");

	if (num == 0) {
		$("#Useable").val("Y");
		$(".selectbox").show();
		$(".my-recode-ct li").eq(0).trigger("click");
		get_CouponList(1);
	}
	else {
		$(".selectbox").hide();
		get_IngCouponList();
	}
}


/* 쿠폰 리스트 */
function get_CouponList(page) {

	$("#Page").val(page);

	$.ajax({
		type		 : "post",
		url			 : "/ASP/Mypage/Ajax/CouponList.asp",
		async		 : false,
		data		 : $("#form").serialize(),
		dataType	 : "text",
		success		 : function (data) {
						var splitData	 = data.split("|||||");
						var result		 = splitData[0];
						var cont		 = splitData[1];
						var cnt			 = splitData[2];


						if (result == "OK") {
							$("#CouponList").html(cont);
							$("#CouponCnt").html(cnt);
							return;
						}
						else if (result == "LOGIN") {
							PageReload();
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

/* 쿠폰 리스트 */
function get_IngCouponList() {

	$.ajax({
		type		 : "post",
		url			 : "/ASP/Mypage/Ajax/IngCouponList.asp",
		async		 : false,
		data		 : $("#form").serialize(),
		dataType	 : "text",
		success		 : function (data) {
						var splitData	 = data.split("|||||");
						var result		 = splitData[0];
						var cont		 = splitData[1];
						var cnt			 = splitData[2];


						if (result == "OK") {
							$("#CouponList").html(cont);
							$("#CouponCnt").html(cnt);
							return;
						}
						else if (result == "LOGIN") {
							PageReload();
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

/* 쿠폰다운로드 */
function couponDown(cIdx) {
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
							openAlertLayer("alert", "쿠폰이 지급 되었습니다.<br />마이페이지 MY쿠폰에서 확인 가능합니다.", "closePop('alertPop', '')", "");
							return;
						}
						else if (result == "LOGIN") {
							openAlertLayer("alert", "로그인 후 이용 가능합니다.", "closePop('alertPop', '');APP_TopGoUrl('/ASP/Member/Login.asp');", "");
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
/**************************************************************************************/
/* 쿠폰북 END
/**************************************************************************************/


/**************************************************************************************/
/* 포인트 START
/**************************************************************************************/
/* 포인트 리스트 */
function get_PointList(page) {
	//location.href = "/ASP/Mypage/Ajax/PointList.asp?Page=" + page;
	//return;

	$.ajax({
		type		 : "post",
		url			 : "/ASP/Mypage/Ajax/PointList.asp",
		async		 : false,
		data		 : "Page=" + page,
		dataType	 : "text",
		success		 : function (data) {
						var splitData	 = data.split("|||||");
						var result		 = splitData[0];
						var cont		 = splitData[1];


						if (result == "OK") {
							$("#PointList").html(cont);
							return;
						}
						else if (result == "LOGIN") {
							PageReload();
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
/**************************************************************************************/
/* 슈즈상품권 END
/**************************************************************************************/


/* 회원탈퇴 체크 */
function chk_MyDraw() {
	var cCnt = $("input[name='wdReason']:checked", "form[name='withDrawForm']").length;
	if (cCnt == 0) {
		common_msgPopOpen("", "탈퇴 사유를 선택해 주십시오.");
		return;
	}

	/* 비밀번호 */
	var pwd = alltrim($("input[name='Pwd']", "form[name='withDrawForm']").val());
	if (pwd.length == 0) {
		common_msgPopOpen("", "비밀번호를 입력해 주십시오.", "", "withDrawForm.Pwd");
		return;
	}

	common_msgPopOpen("", "회원탈퇴를 하시겠습니까?", "myDrawOk();", "", "C");
}

function myDrawOk() {
	$.ajax({
		type: "post",
		url: "/ASP/Mypage/Ajax/WithdrawOk.asp",
		async: false,
		data: $("#withDrawForm").serialize(),
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];
			var uid = splitData[2];

			if (result == "OK") {

				AceCounter_Withdraw(uid);

				common_msgPopOpen("", "탈퇴처리가 완료되었습니다.<br />이용해 주셔서 감사합니다.", "location.replace('/ASP/Member/Logout.asp');");
				APP_goMain();
				return;
			}
			else {
				common_msgPopOpen("", cont);
				return;
			}
		},
		error: function (data) {
			//alert(data.responseText)
			common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.");
		}
	});
}


/* 배송지 추가 체크 */
function chk_MyAddr() {
	var AddressName = $("input[name='AddressName']", "form[name='MyAddress']");
	if (AddressName.val().length == 0) {
		common_msgPopOpen("", "배송지명을 입력해 주십시오.", "", "AddressName");
		return;
	}
	var ReceiveName = $("input[name='ReceiveName']", "form[name='MyAddress']");
	if (ReceiveName.val().length == 0) {
		common_msgPopOpen("", "받으시는분을 입력해 주십시오.", "", "ReceiveName");
		return;
	}
	var ReceiveTel1 = $("select[name='ReceiveTel1']", "form[name='MyAddress']");
	var ReceiveTel23 = $("input[name='ReceiveTel23']", "form[name='MyAddress']");
	if (ReceiveTel1.val() != "" || ReceiveTel23.val().length > 0) {
		if (ReceiveTel1.val()=="" || ReceiveTel23.val().length < 7) {
			common_msgPopOpen("", "올바른 연락처를 입력해 주십시오.", "", "ReceiveTel23");
			return;
		}
		if (only_Num(ReceiveTel23.val()) == false) {
			common_msgPopOpen("", "연락처를 숫자로만 입력하여 주십시오.", "", "ReceiveTel23");
			return;
		}
	}
	var ReceiveHP1 = $("select[name='ReceiveHP1']", "form[name='MyAddress']");
	var ReceiveHP23 = $("input[name='ReceiveHP23']", "form[name='MyAddress']");
	if (ReceiveHP1.val()=="" || ReceiveHP23.val().length < 7) {
		common_msgPopOpen("", "올바른 휴대폰 번호를 입력해 주십시오.", "", "ReceiveHP23");
		return;
	}
	if (only_Num(ReceiveHP23.val()) == false) {
		common_msgPopOpen("", "휴대폰 번호를 숫자로만 입력하여 주십시오.", "", "ReceiveHP23");
		return;
	}
	var ReceiveZipCode = $("input[name='ReceiveZipCode']", "form[name='MyAddress']");
	var ReceiveAddr1 = $("input[name='ReceiveAddr1']", "form[name='MyAddress']");
	if (ReceiveZipCode.val().length == 0 || ReceiveAddr1.val().length == 0) {
		common_msgPopOpen("", "배송지 주소를 입력해 주십시오.", "", "ReceiveAddr1");
		return;
	}

	var ReceiveAddr2 = $("input[name='ReceiveAddr2']", "form[name='MyAddress']");
	if (ReceiveAddr2.val().length == 0) {
		common_msgPopOpen("", "상세 주소를 입력해 주십시오.", "", "ReceiveAddr2");
		return;
	}

	$.ajax({
		type: "post",
		url: "/ASP/Mypage/Ajax/MyAddrOk.asp",
		async: false,
		data: $("#MyAddress").serialize(),
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				common_msgPopOpen("", cont, "location.reload();");
				return;
			}
			else {
				common_msgPopOpen("", cont);
				return;
			}
		},
		error: function (data) {
			//alert(data.responseText)
			common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.");
		}
	});	
}

/* 배송지 입력 */
function insert_MyAddr(type, idx) {
	if (type == "insert") {
		$("input[name='addrType']", "form[name='MyAddrListForm']").val("insert");
		$("input[name='idx']", "form[name='MyAddrListForm']").val("");
		common_PopOpen('DimDepth1', 'MyAddr');
	} else if (type == "modify") {
		$("input[name='addrType']", "form[name='MyAddrListForm']").val("modify");
		$("input[name='idx']", "form[name='MyAddrListForm']").val(idx);
		common_PopOpen('DimDepth1', 'MyAddr');
	}
}

/* 기본배송지 설정 */
function chg_MainFlag(idx) {
	if (idx == "") {
		common_msgPopOpen("", "배송지 정보가 없습니다.");
		return;
	}

	common_msgPopOpen("", "기본 배송지 정보로 수정 하시겠습니까?", "chg_MainFlagOk(" + idx +");", "", "C");
}

function chg_MainFlagOk(idx) {
	$.ajax({
		type: "post",
		url: "/ASP/Mypage/Ajax/MyAddrMainFlagOk.asp",
		async: false,
		data: "idx = " + idx,
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				common_msgPopOpen("", cont, "PageReload();");
				return;
			}
			else {
				common_msgPopOpen("", cont);
				return;
			}
		},
		error: function (data) {
			//alert(data.responseText)
			common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.");
		}
	});	
}


/* 배송지 삭제 */
function del_MyAddr(idx) {
	if (idx == "") {
		common_msgPopOpen("", "삭제정보가 없습니다.");
		return;
	}

	common_msgPopOpen("", "해당 배송지 정보를 삭제 하시겠습니까?", "del_MyAddrOk(" + idx +");", "", "C");
}

function del_MyAddrOk(idx) {
	$.ajax({
		type: "post",
		url: "/ASP/Mypage/Ajax/MyAddrDel.asp",
		async: false,
		data: "idx = " + idx,
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				common_msgPopOpen("", cont, "PageReload();");
				return;
			}
			else {
				common_msgPopOpen("", cont);
				return;
			}
		},
		error: function (data) {
			//alert(data.responseText)
			common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.");
		}
	});	
}


/* 나의정보수정 비밀번호 체크 */
function chk_MyPwd() {
	/* 비밀번호 */
	var pwd = alltrim($("input[name='Pwd']", "form[name='chkPwdForm1']").val());
	if (pwd.length == 0) {
		common_msgPopOpen("", "현재 비밀번호를 입력해 주십시오.", "", "chkPwdForm1.Pwd");
		return;
	}

	//location.href = "/ASP/Mypage/Ajax/MyInfoModifyPwChkOk.asp?" + $("#chkPwdForm1").serialize();
	//return;

	$.ajax({
		type: "post",
		url: "/ASP/Mypage/Ajax/MyInfoModifyPwChkOk.asp",
		async: false,
		data: $("#chkPwdForm1").serialize(),
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];


			if (result == "OK") {
				location.href = "/ASP/Mypage/MyInfoModify.asp";
				return;
			}
			else if (result == "FAIL") {
				common_msgPopOpen("", cont);
				return;
			}
			else {
				common_msgPopOpen("", cont);
				return;
			}
		},
		error: function (data) {
			//alert(data.responseText)//alert("처리 도중 오류가 발생하였습니다.");
			common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.");
		}
	});
}


/* 회원정보 수정 폼 체크 */
function chk_MyInfoModify() {

	var pwd = alltrim($("input[name='Pwd']", "form[name='MyInfoModify']").val());
	if (pwd.length == 0) {
		common_msgPopOpen("", "비밀번호를 입력해 주십시오.", "", "MyInfoModify.Pwd");
		return;
	}
	if (pwd.length < 6 || pwd.length > 12) {
		common_msgPopOpen("", "비밀번호를 영문, 숫자조합 6자리이상 12자리 이내로 입력해 주십시오.", "", "MyInfoModify.Pwd");
		return;
	}
	if (only_AlphaNum(pwd) == false) {
		common_msgPopOpen("", "비밀번호를 영문, 숫자조합 6자리이상 12자리 이내로 입력해 주십시오.", "", "MyInfoModify.Pwd");
		return;
	}
	var userID = alltrim($("input[name='UserID']", "form[name='MyInfoModify']").val());
	if (pwd == userID) {
		common_msgPopOpen("", "아이디와 같은 비밀번호는 사용할 수 없습니다.", "", "MyInfoModify.Pwd");
		return;
	}
	if (chk_SameChr(pwd, 4) == false) {
		common_msgPopOpen("", "비밀번호는 4자리이상 동일한 문자를 사용할 수 없습니다.", "", "MyInfoModify.Pwd");
		return;
	}



	/* 비밀번호 확인 */
	var pwd1 = alltrim($("input[name='Pwd1']", "form[name='MyInfoModify']").val());
	if (pwd1.length == 0) {
		common_msgPopOpen("", "비밀번호를 다시 한번 입력해 주십시오.", "", "MyInfoModify.Pwd1");
		return;
	}

	if (pwd != pwd1) {
		common_msgPopOpen("", "비밀번호가 일치하지 않습니다.", "", "MyInfoModify.Pwd1");
		return;
	}


	var name = alltrim($("input[name='Name']", "form[name='MyInfoModify']").val());
	if (name.length == 0) {
		common_msgPopOpen("", "이름을 입력해 주십시오.", "", "MyInfoModify.Name");
		return;
	}

	var birth = alltrim($("input[name='Birth']", "form[name='MyInfoModify']").val());
	if (birth.length == 0) {
		common_msgPopOpen("", "생년월일을 입력해 주십시오.", "", "MyInfoModify.Birth");
		return;
	}
	if (only_Num(birth) == false) {
		common_msgPopOpen("", "생년월일을 숫자로만 입력해 주십시오.", "", "MyInfoModify.Birth");
		return;
	}
	if (birth.length != 8) {
		common_msgPopOpen("", "생년월일을 숫자 8자리로 입력해 주십시오.", "", "MyInfoModify.Birth");
		return;
	}


	var bYear = String(birth).substring(0, 4);
	var bMonth = String(birth).substring(4, 6);
	var bDay = String(birth).substring(6, 8);
	if (checkDate(bYear, bMonth, bDay) == false) {
		common_msgPopOpen("", "생년월일 입력이 잘 못 되었습니다.", "", "MyInfoModify.Birth");
		return;
	}


	var sexFlag = $("input[name='SexFlag']", "form[name='MyInfoModify']").val();
	if (sexFlag == "N") {
		var sCnt = $("input:radio[name='Sex']:checked", "form[name='MyInfoModify']").length;
		if (sCnt == 0) {
			common_msgPopOpen("", "성별을 선택해 주십시오.");
			$("input:radio[name='Sex']:checked", "form[name='MyInfoModify']").eq(0).focus();
			return;
		}
	}

	var zipcode = $("input[name='ZipCode']", "form[name='MyInfoModify']").val();
	var addr1 = $("input[name='Addr1']", "form[name='MyInfoModify']").val();
	if (zipcode.length == 0 || addr1.length == 0) {
		openAlertLayer("alert", "우편번호를 검색하여 입력해 주십시오.", "closePop('alertPop', 'ZipCode');", "");
		return;
	}
	var addr2 = $("input[name='Addr2']", "form[name='MyInfoModify']").val();
	if (addr2.length == 0) {
		openAlertLayer("alert", "상세주소를 입력해 주십시오.", "closePop('alertPop', 'Addr2');", "");
		return;
	}

	var hp1 = $("select[name='HP1']", "form[name='MyInfoModify']").val();
	if (hp1.length == 0) {
		common_msgPopOpen("", "휴대폰번호를 선택해 주십시오.", "", "MyInfoModify.HP1");
		return;
	}

	var hp23 = $("input[name='HP23']", "form[name='MyInfoModify']").val();
	if (hp23.length < 7) {
		common_msgPopOpen("", "올바른 휴대폰번호를 입력해 주십시오.", "", "MyInfoModify.HP23");
		return;
	}
	if (only_Num(hp23) == false) {
		common_msgPopOpen("", "휴대폰번호를 숫자로만 입력해 주십시오.", "", "MyInfoModify.HP23");
		return;
	}

	var email = alltrim($("input[name='Email']", "form[name='MyInfoModify']").val());
	if (email.length == 0) {
		common_msgPopOpen("", "이메일 계정을 입력해 주십시오.", "", "MyInfoModify.Email");
		return;
	}
	if (beAllowStr(email, "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890.-_@") == false) {
		common_msgPopOpen("", "이메일 주소를 영문과 숫자로만 입력해 주십시오.", "", "MyInfoModify.Email");
		return;
	}
	if (email.indexOf(".") < 0 || email.indexOf(".") == 0 || email.indexOf(".") == email.length - 1 || email.indexOf("@") < 0 || email.indexOf("@") == 0 || email.indexOf("@") == email.length - 1) {
		common_msgPopOpen("", "올바른 이메일 주소를 다시 입력해 주십시오.", "", "MyInfoModify.Email");
		return;
	}

	/* 보호자 입력정보 */
	var FTFlag = $("input[name='FTFlag']", "form[name='MyInfoModify']").val();
	if (FTFlag == "Y") {

		var parentName = alltrim($("input[name='ParentName']", "form[name='MyInfoModify']").val());
		if (parentName.length == 0) {
			common_msgPopOpen("", "보호자 이름을 입력해 주십시오.", "", "MyInfoModify.ParentName");
			return;
		}
		var parentBirth = alltrim($("input[name='ParentBirth']", "form[name='MyInfoModify']").val());
		if (parentBirth.length == 0) {
			common_msgPopOpen("", "생년월일을 입력해 주십시오.", "", "MyInfoModify.ParentBirth");
			return;
		}
		if (only_Num(parentBirth) == false) {
			common_msgPopOpen("", "생년월일을 숫자로만 입력해 주십시오.", "", "MyInfoModify.ParentBirth");
			return;
		}
		if (parentBirth.length != 8) {
			common_msgPopOpen("", "생년월일을 숫자 8자리로 입력해 주십시오.", "", "MyInfoModify.ParentBirth");
			return;
		}

		var pYear = String(parentBirth).substring(0, 4);
		var pMonth = String(parentBirth).substring(4, 6);
		var pDay = String(parentBirth).substring(6, 8);
		if (checkDate(pYear, pMonth, pDay) == false) {
			common_msgPopOpen("", "생년월일 입력이 잘 못 되었습니다.", "", "MyInfoModify.ParentBirth");
			return;
		}

		var php1 = $("select[name='PHP1']", "form[name='MyInfoModify']").val();
		if (php1.length == 0) {
			common_msgPopOpen("", "보호자 휴대폰번호를 선택해 주십시오.", "", "MyInfoModify.PHP1");
			return;
		}

		var php2 = $("input[name='PHP2']", "form[name='MyInfoModify']").val();
		if (php2.length < 7) {
			common_msgPopOpen("", "올바른 보호자 휴대폰번호를 입력해 주십시오.", "", "MyInfoModify.PHP2");
			return;
		}
		if (only_Num(php2) == false) {
			common_msgPopOpen("", "휴대폰번호를 숫자로만 입력해 주십시오.", "", "MyInfoModify.PHP2");
			return;
		}

		var parentEmail = alltrim($("input[name='ParentEmail']", "form[name='MyInfoModify']").val());
		if (parentEmail.length == 0) {
			common_msgPopOpen("", "이메일 계정을 입력해 주십시오.", "", "MyInfoModify.ParentEmail");
			return;
		}
		if (beAllowStr(parentEmail, "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890.-_@") == false) {
			common_msgPopOpen("", "이메일 계정을 영문과 숫자로만 입력해 주십시오.", "", "MyInfoModify.ParentEmail");
			return;
		}
		if (parentEmail.indexOf(".") < 0 || parentEmail.indexOf(".") == 0 || parentEmail.indexOf(".") == parentEmail.length - 1) {
			common_msgPopOpen("", "이메일 도메인을 다시 입력해 주십시오.", "", "MyInfoModify.ParentEmail");
			return;
		}
	}

	common_msgPopOpen("", "입력하신 사항으로 정보를 수정 하시겠습니까?", "MyInfoModifyOk();", "", "C");
}

function MyInfoModifyOk() {
	//location.href = "/ASP/Mypage/Ajax/MyInfoModifyOk.asp?" + $("#MyInfoModify").serialize();
	//return;

	$.ajax({
		type: "post",
		url: "/ASP/Mypage/Ajax/MyInfoModifyOk.asp",
		async: false,
		data: $("#MyInfoModify").serialize(),
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];


			if (result == "OK") {
				common_msgPopOpen("", cont, "location.href='/ASP/Mypage/';");
				return;
			}
			else if (result == "FAIL_LOGIN") {
				common_msgPopOpen("", cont, "location.href='/';");
				return;
			}
			else {
				common_msgPopOpen("", cont);
				return;
			}
		},
		error: function (data) {
			common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.");
		}
	});
}



/* 쿠폰교환 - 바우처 */
function chg_Coupon() {
	var couponNum = alltrim($("input[name='CouponNum']", "form[name='cForm']").val());
	if (couponNum.length == 0) {
		openAlertLayer("alert", "쿠폰번호를 입력하여 주십시오.", "closePop('alertPop', 'CouponNum')", "");
		return;
	}


	$.ajax({
		type		 : "post",
		url			 : "/ASP/Mypage/Ajax/CouponChangeOk.asp",
		async		 : false,
		data		 : $("#cForm").serialize(),
		dataType	 : "text",
		success		 : function (data) {
						var splitData	 = data.split("|||||");
						var result		 = splitData[0];
						var cont		 = splitData[1];


						if (result == "OK") {
							openAlertLayer("alert", "쿠폰을 발급 되었습니다.", "closePop('alertPop', '')", "");
							document.cForm.reset();
							clk_CouponType(0);
							return;
						}
						else if (result == "LOGIN") {
							PageReload();
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


/* 구매후기 입력 */
function insert_MyReview(type, ordercode, idx) {
	$("input[name='ordercode']", "form[name='MyReviewListForm']").val(ordercode);
	$("input[name='idx']", "form[name='MyReviewListForm']").val(idx);
	if (type == "add") {
		$("input[name='reviewGubun']", "form[name='MyReviewListForm']").val("add");
		common_PopOpen('DimDepth1', 'MyReviewWrite');
	} else if (type == "view") {
		$("input[name='reviewGubun']", "form[name='MyReviewListForm']").val("view");
		common_PopOpen('DimDepth1', 'MyReviewWrite');
	}
}

//에이스 카운터 탈퇴
function AceCounter_Withdraw(uid) {
	var m_jn = 'withdraw';          //  가입탈퇴 ( 'join','withdraw' ) 
	var m_jid = uid;				// 가입시입력한 ID

	var _AceGID = (function () { var Inf = ['app.shoemarker.co.kr', 'app.shoemarker.co.kr', 'AZ1A74686', 'AM', '0', 'NaPm,Ncisy', 'ALL', '0']; var _CI = (!_AceGID) ? [] : _AceGID.val; var _N = 0; if (_CI.join('.').indexOf(Inf[3]) < 0) { _CI.push(Inf); _N = _CI.length; } return { o: _N, val: _CI }; })();
	var _AceCounter = (function () { var G = _AceGID; var _sc = document.createElement('script'); var _sm = document.getElementsByTagName('script')[0]; if (G.o != 0) { var _A = G.val[G.o - 1]; var _G = (_A[0]).substr(0, _A[0].indexOf('.')); var _C = (_A[7] != '0') ? (_A[2]) : _A[3]; var _U = (_A[5]).replace(/\,/g, '_'); _sc.src = (location.protocol.indexOf('http') == 0 ? location.protocol : 'http:') + '//cr.acecounter.com/Mobile/AceCounter_' + _C + '.js?gc=' + _A[2] + '&py=' + _A[1] + '&up=' + _U + '&rd=' + (new Date().getTime()); _sm.parentNode.insertBefore(_sc, _sm); return _sc.src; } })();
}