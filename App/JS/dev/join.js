/* 14 구분 회원인증 페이지 이동 */
function move_Certification(form, jType) {
	$("input[name='JoinType']").val(jType);
	$("#formMoveCert").submit();
}

/* 필수 약관 선택 여부 */
function agr_MemberTerms(form, type) {
	if (document.form.agreechk1.checked == false) {
		openAlertLayer("alert", "쇼핑몰 이용약관에 동의를 하셔야 합니다.", "closePop('alertPop', 'agreement1');", "");
		return;
	}

	if (document.form.agreechk2.checked == false) {
		openAlertLayer("alert", "개인정보 이용 및 수집에 대한 동의를 하셔야 합니다.", "closePop('alertPop', 'agreement2');", "");
		return;
	}

	//휴대폰인증
	if (type == "Nice") {
		auth_HP(form);
	}
	//아이핀인증
	else if (type == "Ipin") {
		auth_Ipin(form);
	}
}


/* 핸드폰인증 */
function auth_HP(form) {
	APP_PopupGoUrl('/Common/AuthHP/Nice.asp?' + $("#" + form).serialize(), '0', '');
	/*
	window.name = "Parent_window";
	window.open('', 'popupNice', 'width=450, height=550, top=100, left=100, fullscreen=no, menubar=no, status=no, toolbar=no, titlebar=yes, location=no, scrollbar=no');
	eval("document." + form).target = "popupNice";
	eval("document." + form).action = "/Common/AuthHP/Nice.asp";
	eval("document." + form).submit();
	*/
}


/* 아이핀인증 */
function auth_Ipin(form) {
	APP_PopupGoUrl('/Common/AuthIpin/Ipin.asp?' + $("#" + form).serialize(), '0', '');
	/*
	window.name = "Parent_window";
	window.open('', 'popupIpin', 'width=450, height=550, top=100, left=100, fullscreen=no, menubar=no, status=no, toolbar=no, titlebar=yes, location=no, scrollbar=no');
	eval("document." + form).target = "popupIpin";
	eval("document." + form).action = "/Common/AuthIpin/Ipin.asp";
	eval("document." + form).submit();
	*/
}


/* 본인인증 완료 후 가입정보 기입 페이지로 이동 */
function goJoin() {
	document.form.target = "_self";
	document.form.action = "/ASP/Member/JoinForm.asp";
	document.form.submit();
}

/* 본인인증 완료 후 가입정보 기입 페이지로 이동(SNS 정회원 전환) */
function goJoinChgMem() {
	document.form.target = "_self";
	document.form.action = "/ASP/Member/JoinChgMemForm.asp";
	document.form.submit();
}


/* 회원 아이디 체크 */
function JoinUserIDCheckResult(RCType) {
	if (RCType == "1") {
		document.getElementById("idAvailable").innerHTML = "사용 가능한 아이디 입니다.";
		document.Joinform.UserIDCheckFlag.value = "Y";
	} else {
		document.getElementById("idAvailable").innerHTML = "이미 사용 중인 아이디 입니다.";
		document.Joinform.UserIDCheckFlag.value = "N";
	}
}


/* 회원 아이디 체크 */
function IsID(formname) {
	var form = eval("document.Joinform." + formname);
	if (form.value.length < 6 || form.value.length > 16) {
		return;
	}
	for (var i = 0; i < form.value.length; i++) {
		var chr = form.value.substr(i, 1);
		if ((chr < '0' || chr > '9') && (chr < 'a' || chr > 'z') && (chr < 'A' || chr > 'Z')) {
			return;
		}
	}
	return true;
}

/* 비밀번호 체크 */
function IsPW(formname) {
	var form = eval("document.Joinform." + formname);
	if (form.value.length < 8 || form.value.length > 16) {
		return;
	}
	for (var i = 0; i < form.value.length; i++) {
		var chr = form.value.substr(i, 1);
		if ((chr < '0' || chr > '9') && (chr < 'a' || chr > 'z') && (chr < 'A' || chr > 'Z')) {
			return;
		}
	}
	return true;
}

function Password_Confirm()
{
	if (document.Joinform.Password.value != document.Joinform.Password_Check.value)
	{
		document.getElementById("pwdchkAvalilable").style.display = "block";
	}
	else
	{
		document.getElementById("pwdchkAvalilable").style.display = "none";
	}
}

/* 이메일 선택 */
function EmailChg(val) {
	if (val == "etc") {
		$("#Email2").val('');
		$("#Email2").attr("readonly", false);
		$("#Email2").focus();
	}
	else {
		$("#Email2").val(val);
		$("#Email2").attr("readonly", true);
		$("#Email2").focus();
	}
}


/* 부모 이메일 선택 */
function ParentEmailChg(val)
{
	if (val == "etc")
	{
		$("#ParentEmail2").val('');
		$("#ParentEmail2").attr("readonly", false);
		$("#ParentEmail2").focus();
	}
	else
	{
		$("#ParentEmail2").val(val);
		$("#ParentEmail2").attr("readonly", true);
		$("#ParentEmail2").focus();
	}
}

/* 회원 아이디 체크 */
function chk_UserID() {
	var userID = alltrim($("input[name='UserID']", "form[name='Joinform']").val());
	if (userID.length == 0) {
		openAlertLayer("alert", "아이디를 입력하여 주십시오.", "closePop('alertPop', 'UserID');", "");
		return;
	}
	if (userID.length < 6 || userID.length > 12) {
		openAlertLayer("alert", "아이디를 영문, 숫자조합 6자리이상 12자리 이내로 입력해 주십시오.", "closePop('alertPop', 'UserID');", "");
		return;
	}
	if (only_AlphaNum(userID) == false) {
		openAlertLayer("alert", "아이디를 영문, 숫자조합 6자리이상 12자리 이내로 입력해 주십시오.", "closePop('alertPop', 'UserID');", "");
		return;
	}

	$.ajax({
		type		 : "post",
		url			 : "/ASP/Member/Ajax/CheckUserID.asp",
		async		 : false,
		data		 : "UserID=" + userID,
		dataType	 : "text",
		success		 : function (data) {
						var splitData	 = data.split("|||||");
						var result		 = splitData[0];
						var cont		 = splitData[1];


						if (result == "OK") {
							$("#ID_Msg").css({ "color": "blue" });
							$("#ID_Msg").text("등록 가능한 아이디 입니다.");
							$("#ID_Msg").show();
							$("#CheckID").val(userID);
							$("#CheckIDAvailable").val("Y");
							return;
						}
						else if (result == "EXISTS") {
							$("#ID_Msg").css({ "color": "red" })
							$("#ID_Msg").text(cont);
							$("#ID_Msg").show();
							$("#CheckID").val("");
							$("#CheckIDAvailable").val("");
							return;
						}
						else {
							alert(cont);
							return;
						}
		},
		error		 : function (data) {
						common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.");
		}
	});
}

/* 회원가입 폼 체크 */
function chk_Join() {

	var joinType = $("input[name='JoinType']", "form[name='Joinform']").val();

	var userID = alltrim($("input[name='UserID']", "form[name='Joinform']").val());
	if (userID.length == 0) {
		openAlertLayer("alert", "아이디를 입력하여 주십시오.", "closePop('alertPop', 'UserID');", "");
		return;
	}
	if (userID.length < 6 || userID.length > 12) {
		openAlertLayer("alert", "아이디를 영문, 숫자조합 6자리이상 12자리 이내로 입력해 주십시오.", "closePop('alertPop', 'UserID');", "");
		return;
	}
	if (only_AlphaNum(userID) == false) {
		openAlertLayer("alert", "아이디를 영문, 숫자조합 6자리이상 12자리 이내로 입력해 주십시오.", "closePop('alertPop', 'UserID');", "");
		return;
	}


	var checkIDAvailable = $("input[name='CheckIDAvailable']", "form[name='Joinform']").val();
	if (checkIDAvailable != "Y") {
		openAlertLayer("alert", "아이디 중복 체크를 해 주십시오.", "closePop('alertPop', 'IDChkBtn');", "");
		return;
	}
	var checkID = $("input[name='CheckID']", "form[name='Joinform']").val();
	if (checkID != userID) {
		openAlertLayer("alert", "아이디 중복 체크를 다시 해 주십시오.", "closePop('alertPop', 'IDChkBtn');", "");
		return;
	}

	var pwd = alltrim($("input[name='Pwd']", "form[name='Joinform']").val());
	if (pwd.length == 0) {
		openAlertLayer("alert", "비밀번호를 입력해 주십시오.", "closePop('alertPop', 'Pwd');", "");
		return;
	}
	if (pwd.length < 6 || pwd.length > 12) {
		openAlertLayer("alert", "비밀번호를 영문, 숫자조합 6자리이상 12자리 이내로 입력해 주십시오.", "closePop('alertPop', 'Pwd');", "");
		return;
	}
	if (!only_AlphaNum(pwd)) {
		openAlertLayer("alert", "비밀번호를 영문, 숫자조합 6자리이상 12자리 이내로 입력해 주십시오.", "closePop('alertPop', 'Pwd');", "");
		return;
	}

	if (userID == pwd) {
		openAlertLayer("alert", "아이디와 동일한 비밀번호를 사용하실 수 없습니다.", "closePop('alertPop', 'Pwd');", "");
		return;
	}

	if (chk_SameChr(pwd, 4) == false) {
		openAlertLayer("alert", "비밀번호는 4자리이상 동일한 문자를 사용할 수 없습니다.", "closePop('alertPop', 'Pwd');", "");
		return;
	}

	var pwd1 = alltrim($("input[name='Pwd1']", "form[name='Joinform']").val());
	if (pwd1.length == 0) {
		openAlertLayer("alert", "비밀번호를 다시 한번 입력해 주십시오.", "closePop('alertPop', 'Pwd1');", "");
		return;
	}
	if (pwd != pwd1) {
		openAlertLayer("alert", "비밀번호가 일치하지 않습니다.", "closePop('alertPop', 'Pwd1');", "");
		return;
	}


	var name = alltrim($("input[name='Name']", "form[name='Joinform']").val());
	if (name.length == 0) {
		openAlertLayer("alert", "이름을 입력해 주십시오.", "closePop('alertPop', 'Name');", "");
		return;
	}


	var birth = alltrim($("input[name='Birth']", "form[name='Joinform']").val());
	if (birth.length == 0) {
		openAlertLayer("alert", "생년월일을 입력해 주십시오.", "closePop('alertPop', 'Birth');", "");
		return;
	}
	if (only_Num(birth) == false) {
		openAlertLayer("alert", "생년월일을 숫자로만 입력해 주십시오.", "closePop('alertPop', 'Birth');", "");
		return;
	}
	if (birth.length != 8) {
		openAlertLayer("alert", "생년월일을 숫자 8자리로 입력해 주십시오.", "closePop('alertPop', 'Birth');", "");
		return;
	}


	var bYear	 = String(birth).substring(0, 4);
	var bMonth	 = String(birth).substring(4, 6);
	var bDay	 = String(birth).substring(6, 8);

	if (checkDate(bYear, bMonth, bDay) == false) {
		openAlertLayer("alert", "생년월일 입력이 잘 못 되었습니다.", "closePop('alertPop', 'Birth');", "");
		return;
	}
	var today = $("input[name='Today']", "form[name='Joinform']").val();
	if (joinType == "D") {
		if ((parseFloat(today) - parseFloat(birth)) >= 140000) {
			openAlertLayer("alert", "만14세 이상 인증을 받아 주십시오.", "closePop('alertPop', 'Birth');", "");
			return;
		}
	}
	else {
		if ((parseFloat(today) - parseFloat(birth)) < 140000) {
			alert("만14세 이하 인증을 받아 주십시오.");
			return;
		}
	}


	var sCnt = $("input:radio[name='Sex']:checked", "form[name='Joinform']").length;
	if (sCnt == 0) {
		openAlertLayer("alert", "성별을 선택해 주십시오.", "closePop('alertPop', 'SexArea');", "");
		return;
	}


	var sCnt = $("input:radio[name='Sex']:checked", "form[name='Joinform']").length;
	if (sCnt == 0) {
		openAlertLayer("alert", "성별을 선택해 주십시오.", "closePop('alertPop', 'SexArea');", "");
		return;
	}

	var mobileFlag = $("input[name='MobileFlag']", "form[name='Joinform']").val();
	if (mobileFlag == "N") {

		var hp1 = $("select[name='HP1']", "form[name='Joinform']").val();
		if (hp1.length == 0) {
			openAlertLayer("alert", "휴대폰번호를 선택해 주십시오.", "closePop('alertPop', 'HP1');", "");
			return;
		}

		var hp2 = $("input[name='HP2']", "form[name='Joinform']").val();
		if (hp2.length == 0) {
			openAlertLayer("alert", "휴대폰번호를 입력해 주십시오.", "closePop('alertPop', 'HP2');", "");
			return;
		}
		if (hp1 == "010" && hp2.length != 4) {
			openAlertLayer("alert", "휴대폰번호를 숫자 4자리로 입력해 주십시오.", "closePop('alertPop', 'HP2');", "");
			return;
		}
		if (hp1 != "010" && hp2.length < 3) {
			openAlertLayer("alert", "휴대폰번호를 숫자 3자리 이상으로 입력해 주십시오.", "closePop('alertPop', 'HP2');", "");
			return;
		}
		if (only_Num(hp2) == false) {
			openAlertLayer("alert", "휴대폰번호를 숫자로만 입력해 주십시오.", "closePop('alertPop', 'HP2');", "");
			return;
		}


		var hp3 = $("input[name='HP3']", "form[name='Joinform']").val();
		if (hp3.length == 0) {
			openAlertLayer("alert", "휴대폰번호를 입력해 주십시오.", "closePop('alertPop', 'HP3');", "");
			return;
		}
		if (hp3.length != 4) {
			openAlertLayer("alert", "휴대폰번호를 숫자 4자리로 입력해 주십시오.", "closePop('alertPop', 'HP3');", "");
			return;
		}
		if (only_Num(hp3) == false) {
			openAlertLayer("alert", "휴대폰번호를 숫자로만 입력해 주십시오.", "closePop('alertPop', 'HP3');", "");
			return;
		}
	}


	var email = alltrim($("input[name='Email']", "form[name='Joinform']").val());
	if (email.length == 0) {
		openAlertLayer("alert", "이메일을 입력해 주십시오.", "closePop('alertPop', 'Email');", "");
		return;
	}
	if (beAllowStr(email, "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890.-_@") == false) {
		openAlertLayer("alert", "이메일을 영문과 숫자로만 입력해 주십시오.", "closePop('alertPop', 'Email');", "");
		return;
	}
	if (checkEmail(email) == false) {
		openAlertLayer("alert", "이메일 형식이 잘 못 입력 되었습니다.", "closePop('alertPop', 'Email');", "");
		return;
	}

	var zipcode = $("input[name='ZipCode']", "form[name='Joinform']").val();
	var addr1 = $("input[name='Addr1']", "form[name='Joinform']").val();
	if (zipcode.length == 0 || addr1.length == 0) {
		openAlertLayer("alert", "우편번호를 검색하여 입력해 주십시오.", "closePop('alertPop', 'ZipCode');", "");
		return;
	}
	var addr2 = $("input[name='Addr2']", "form[name='Joinform']").val();
	if (addr2.length == 0) {
		openAlertLayer("alert", "상세주소를 입력해 주십시오.", "closePop('alertPop', 'Addr2');", "");
		return;
	}


	if (joinType == "D") {

		var parentName = alltrim($("input[name='ParentName']", "form[name='Joinform']").val());
		if (parentName.length == 0) {
			openAlertLayer("alert", "보호자 이름을 입력해 주십시오.", "closePop('alertPop', 'ParentName');", "");
			return;
		}


		var parentBirth = alltrim($("input[name='ParentBirth']", "form[name='Joinform']").val());
		if (parentBirth.length == 0) {
			openAlertLayer("alert", "생년월일을 입력해 주십시오.", "closePop('alertPop', 'ParentBirth');", "");
			return;
		}
		if (only_Num(parentBirth) == false) {
			openAlertLayer("alert", "생년월일을 숫자로만 입력해 주십시오.", "closePop('alertPop', 'ParentBirth');", "");
			return;
		}
		if (parentBirth.length != 8) {
			openAlertLayer("alert", "생년월일을 숫자 8자리로 입력해 주십시오.", "closePop('alertPop', 'ParentBirth');", "");
			return;
		}


		var pYear	 = String(parentBirth).substring(0, 4);
		var pMonth	 = String(parentBirth).substring(4, 6);
		var pDay	 = String(parentBirth).substring(6, 8);

		if (checkDate(pYear, pMonth, pDay) == false) {
			openAlertLayer("alert", "생년월일 입력이 잘 못 되었습니다.", "closePop('alertPop', 'ParentBirth');", "");
			return;
		}


		var parentMobileFlag = $("input[name='ParentMobileFlag']", "form[name='Joinform']").val();
		if (parentMobileFlag == "N") {
			var php1 = $("select[name='PHP1']", "form[name='Joinform']").val();
			if (php1.length == 0) {
				openAlertLayer("alert", "휴대폰번호를 선택해 주십시오.", "closePop('alertPop', 'PHP1');", "");
				return;
			}

			var php2 = $("input[name='PHP2']", "form[name='Joinform']").val();
			if (php2.length == 0) {
				openAlertLayer("alert", "휴대폰번호를 입력해 주십시오.", "closePop('alertPop', 'PHP2');", "");
				return;
			}
			if (php1 == "010" && php2.length != 4) {
				openAlertLayer("alert", "휴대폰번호를 숫자 4자리로 입력해 주십시오.", "closePop('alertPop', 'PHP2');", "");
				return;
			}
			if (php1 != "010" && php2.length < 3) {
				openAlertLayer("alert", "휴대폰번호를 숫자 3자리 이상으로 입력해 주십시오.", "closePop('alertPop', 'PHP2');", "");
				return;
			}
			if (only_Num(php2) == false) {
				openAlertLayer("alert", "휴대폰번호를 숫자로만 입력해 주십시오.", "closePop('alertPop', 'PHP2');", "");
				return;
			}


			var php3 = $("input[name='PHP3']", "form[name='Joinform']").val();
			if (php3.length == 0) {
				openAlertLayer("alert", "휴대폰번호를 입력해 주십시오.", "closePop('alertPop', 'PHP3');", "");
				return;
			}
			if (php3.length != 4) {
				openAlertLayer("alert", "휴대폰번호를 숫자 4자리로 입력해 주십시오.", "closePop('alertPop', 'PHP3');", "");
				return;
			}
			if (only_Num(php3) == false) {
				openAlertLayer("alert", "휴대폰번호를 숫자로만 입력해 주십시오.", "closePop('alertPop', 'PHP3');", "");
				return;
			}
		}


		var parentEmail = alltrim($("input[name='ParentEmail']", "form[name='Joinform']").val());
		if (parentEmail.length == 0) {
			openAlertLayer("alert", "이메일을 입력해 주십시오.", "closePop('alertPop', 'ParentEmail');", "");
			return;
		}
		if (beAllowStr(parentEmail, "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890.-_@") == false) {
			openAlertLayer("alert", "이메일을 영문과 숫자로만 입력해 주십시오.", "closePop('alertPop', 'ParentEmail');", "");
			return;
		}
		if (checkEmail(parentEmail) == false) {
			openAlertLayer("alert", "이메일 형식이 잘 못 입력 되었습니다.", "closePop('alertPop', 'ParentEmail');", "");
			return;
		}

	}

	openAlertLayer("confirm", "입력하신 사항으로 회원가입 하시겠습니까?", "closePop('confirmPop', '')", "closePop('confirmPop', '');joinOk();");
}

function joinOk() {

	var sMode	 = $("input[name='SMode']", "form[name='Joinform']").val();
	var joinUrl	 = "";

	if (sMode == "MemberJoin") {
		joinUrl = "/ASP/Member/Ajax/JoinOk.asp"
	}
	else if (sMode == "JoinChgMem") {
		joinUrl = "/ASP/Member/Ajax/JoinChgMemOk.asp"
	}

	$.ajax({
		type		 : "post",
		url			 : joinUrl,
		async		 : false,
		data		 : $("#Joinform").serialize(),
		dataType	 : "text",
		success		 : function (data) {
						var splitData	 = data.split("|||||");
						var result		 = splitData[0];
						var cont		 = splitData[1];


						if (result == "OK") {
							location.replace("/ASP/Member/JoinComplete.asp");
							return;
							/*
							if (sMode == "MemberJoin") {
								location.replace("/ASP/Member/JoinComplete.asp");
								return;
							}
							else if (sMode == "JoinChgMem") {
								openAlertLayer("alert", "회원전환이 완료되었습니다.<br />다시 재접속하여 주시기 바랍니다.", "closePop('alertPop', '');location.href='/ASP/Member/Logout.asp';", "");
								return;
							}
							*/
						}
						else {
							openAlertLayer("alert", cont, "closePop('alertPop', '');", "");
							return;
						}
		},
		error		 : function (data) {
						alert(data.responseText)//openAlertLayer("alert", "처리 도중 오류가 발생하였습니다.", "closePop('alertPop', '');", "");
		}
	});
}


/* SNS계정 간편로그인 사용자 체크 */
function chk_SnsJoin() {
	var agr2 = $("input[name='agreechk2']").is(":checked");
	if (agr2 == false) {
		openAlertLayer("alert", "개인정보 이용 및 수집에 대해 동의하여 주십시오.", "closePop('alertPop', 'agreement2');", "");
		return;
	}
	var snsEmail = $("input[name='SnsEmail']", "form[name='formSnsAgreement']").val();
	if (snsEmail.length == 0) {
		openAlertLayer("alert", "이메일을 입력하여 주십시오.", "closePop('alertPop', 'SnsEmail');", "");
		return;
	}
	if (!checkEmail(snsEmail)) {
		openAlertLayer("alert", "올바른 이메일을 입력하여 주십시오.", "closePop('alertPop', 'SnsEmail');", "");
		return;
	}

	$.ajax({
		url			 : '/ASP/Member/Ajax/SnsJoinOk.asp',
		data		 : $("form[name='formSnsAgreement']").serialize(),
		async		 : false,
		type		 : 'post',
		dataType	 : 'html',
		success		 : function (data) {
						var splitData	 = data.split("|||||");
						var result		 = splitData[0];
						var msg			 = splitData[1];

						if (result == "OK") {
							openAlertLayer("alert", "SNS 간편로그인 가입이 완료되었습니다.", "APP_HistoryBack_Login();closePop('alertPop', '');", "");
							return;
						}
						else if (result == "DIDUP") {
							openAlertLayer("alert", msg, "APP_HistoryBack();closePop('alertPop', '');", "");
							return;
						}
						else if (result == "FAIL") {
							openAlertLayer("alert", msg, "APP_HistoryBack();closePop('alertPop', '');", "");
							return;
						}
						else {
							openAlertLayer("alert", msg, "APP_HistoryBack();closePop('alertPop', '');", "");
							return;
						}
		},
		error		 : function (data) {
						openAlertLayer("alert", "처리 도중 오류가 발생하였습니다.", "APP_HistoryBack();closePop('alertPop', '');", "");
		}
	});
}


/* SNS 회원 정회원 전환시 정회원 정보가 있을 경우 통합 폼 페이지 띄우기 */
function SDupInfoChk() {
	APP_PopupGoUrl("/ASP/Member/SnsSDupInfoView.asp", "1");
}


/* SNS 회원이 선택한 기존 정회원과 통합 처리 폼 체크 */
function chk_MyIdCombine() {
	var CombineID = $("select[name='CombineID']", "form[name='MyIdCombine']").val();
	if (CombineID.length == 0) {
		openAlertLayer("alert", "통합하실 계정을 선택하여 주십시오.", "closePop('alertPop', 'CombineID');", "");
		return;
	}

	openAlertLayer("confirm", "선택하신 ID(" + CombineID + ")로 통합하시겠습니까?", "closePop('confirmPop', '')", "closePop('confirmPop', '');MyIdCombineOk();");
}
/* SNS 회원이 선택한 기존 정회원과 통합 처리 */
function MyIdCombineOk() {
	$.ajax({
		url			 : '/ASP/Member/Ajax/SnsMemCombineOk.asp',
		data		 : $("form[name='MyIdCombine']").serialize(),
		async		 : false,
		type		 : 'post',
		dataType	 : 'html',
		success		 : function (data) {
						var splitData	 = data.split("|||||");
						var result		 = splitData[0];
						var msg			 = splitData[1];

						if (result == "OK") {
							openAlertLayer("alert", "계정 통합처리가 완료 되었습니다.<br />다시 로그인 하여 주십시오.", "closePop('alertPop', '');APP_PopupHistoryBack_Move('/ASP/Member/Login.asp');", "");
							return;
						}
						else if (result == "FAIL") {
							openAlertLayer("alert", msg, "closePop('alertPop', '');", "");
							return;
						}
						else {
							openAlertLayer("alert", msg, "closePop('alertPop', '');", "");
							return;
						}
		},
		error		 : function (data) {
						openAlertLayer("alert", "처리 도중 오류가 발생하였습니다.", "closePop('alertPop', '');", "");
		}
	});
}
