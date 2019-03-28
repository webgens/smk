/* LOGIN PAGE ID/PWD FOCUS */
function init_Login() {
	/*
	$("#id_msg").hide();
	$("#pwd_msg").hide();
	*/
}

/* CHECK LOGIN FORM */
function chk_Login() {
	var userID = alltrim($("input[name='UserID']", "form[name='formLogin']").val());
	if (userID.length == 0) {
		openAlertLayer("alert", "아이디를 입력하여 주십시오.", "closePop('alertPop', 'UserID')", "");
		return;
	}

	var pwd = alltrim($("input[name='Pwd']", "form[name='formLogin']").val());
	if (pwd.length == 0) {
		openAlertLayer("alert", "비밀번호를 입력하여 주십시오.", "closePop('alertPop', 'Pwd')", "");
		return;
	}

	var progID = $("input[name='ProgID']", "form[name='formLogin']").val();


	$.ajax({
		type		 : "post",
		url			 : "/ASP/Member/Ajax/LoginOk.asp",
		async		 : false,
		data		 : $("#formLogin").serialize(),
		success		 : function (data) {
						var splitData	 = data.split("|||||");
						var result		 = splitData[0];
						var cont		 = splitData[1];
						var age			 = splitData[2];
						var gender		 = splitData[3];

						if (result == "OK") {

							AceCounter_Login(age, gender, userID);
							
							var wptg_tagscript_vars = wptg_tagscript_vars || [];
							wptg_tagscript_vars.push(
							(function () {
								return {
									wp_hcuid: userID,  /*고객넘버 등 Unique ID (ex. 로그인  ID, 고객넘버 등 )를 암호화하여 대입. *주의 : 로그인 하지 않은 사용자는 어떠한 값도 대입하지 않습니다.*/
									ti: "24585",
									ty: "Login",                        /*트래킹태그 타입 */
									device: "web",                  /*디바이스 종류  (web 또는  mobile)*/
									items: [{
										i: "로그인",          /*전환 식별 코드  (한글 , 영어 , 번호 , 공백 허용 )*/
										t: "로그인",          /*전환명  (한글 , 영어 , 번호 , 공백 허용 )*/
										p: "1",                   /*전환가격  (전환 가격이 없을 경우 1로 설정 )*/
										q: "1"                   /*전환수량  (전환 수량이 고정적으로 1개 이하일 경우 1로 설정 )*/
									}]
								};
							}));

							if (cont != "") {
								$("#messagePop").html(cont);
								openPop('messagePop');
							}
							else {
								location.href = progID;
							}
							return;
						}
						else if (result == "ID") {
							openAlertLayer("alert", "존재하지 않는 아이디 입니다.", "closePop('alertPop', 'UserID')", "");
							return;
						}
						else if (result == "PWD") {
							openAlertLayer("alert", "비밀번호가 일치하지 않습니다.", "closePop('alertPop', 'Pwd')", "");
							return;
						}
						else if (result == "NEWAGREE") {
							document.formLogin.reset();
							openAlertLayer("confirm", "신규 약관에 동의하여 주십시오.", "closePop('confirmPop', '')", "closePop('confirmPop', '');APP_GoUrl('/ASP/Member/NewAgreement.asp');");
							return;
						}
						else if (result == "DORMANCY") {
							document.formLogin.reset();
							APP_GoUrl("/ASP/Member/DormancyRelease.asp");
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

/* NEW TERMS AGREEMENT */
function chk_NewAgreement() {
	var agr1 = $("input[name='Agr1']").is(":checked");
	if (agr1 == false) {
		openAlertLayer("alert", "쇼핑몰 이용약관에 동의 하셔야 됩니다.", "closePop('alertPop', 'agreement1')", "");
		return;
	}
	var agr2 = $("input[name='Agr2']").is(":checked");
	if (agr2 == false) {
		openAlertLayer("alert", "개인정보 이용 및 수집에 대해 동의 하셔야 됩니다.", "closePop('alertPop', 'agreement2')", "");
		return;
	}
	
	openAlertLayer("confirm", "약관 동의 처리 하시겠습니까?", "closePop('confirmPop', '')", "closePop('confirmPop', '');chk_NewAgreementOk();");
}

function chk_NewAgreementOk() {
	$.ajax({
		type		 : "post",
		url			 : "/ASP/Member/Ajax/NewAgreementOk.asp",
		async		 : false,
		data		 : $("#formNewAgreement").serialize(),
		success		 : function (data) {
						var splitData	 = data.split("|||||");
						var result		 = splitData[0];
						var cont		 = splitData[1];
						var cont2		 = splitData[2];

						if (result == "OK") {
							if (cont2 != "") {
								$("#messagePop").html(cont2);
								openPop('messagePop');
							}
							else {
								openAlertLayer("alert", "신규 약관에 동의 완료 되었습니다..", "closePop('alertPop', '');APP_TopGoUrl('" + cont + "');", "");
							}
							return;
						}
						else if (result == "LOGIN") {
							openAlertLayer("alert", "로그인 정보가 없습니다.<br />다시 로그인 하여 주십시오.", "closePop('alertPop', '');APP_HistoryBack();", "");
							return;
						}
						else {
							openAlertLayer("alert", cont, "closePop('alertPop', '')", "");
							return;
						}
		},
		error		 : function (data) {
						openAlertLayer("alert", "처리 도중 오류가 발생하였습니다", "closePop('alertPop', '')", "");
		}
	});
}


/* 아이디찾기 핸드폰/이메일로 찾기 라디오버튼 클릭 */
function chg_FindID_Normal_Type() {
	var findIDType = $("input:radio[name='FindIDType']:checked").val();
	document.formFindID.reset();
	$("input[name='FindIDType']", "form[name='formFindID']").val(findIDType);

	$("#FI_N_Rst_Msg").text("");
	$("#FI_N_Rst").hide();
	$("#FI_N_Form").show();

	$("#FI_N_Btn").show();
	$("#FI_N_Fail_Btn").hide();
	$("#FI_N_Succ_Btn").hide();

	if (findIDType == "mobile") {
		$("#FI_N_Mobile").show();
		$("#FI_N_Email").hide();
	}
	else {
		$("#FI_N_Mobile").hide();
		$("#FI_N_Email").show();
	}
}


/* 비밀번호찾기 핸드폰/이메일로 찾기 라디오버튼 클릭 */
function chg_FindPW_Normal_Type() {
	var findPWType = $("input:radio[name='FindPWType']:checked").val();

	document.formFindPW.reset();
	$("input[name='FindPWType']", "form[name='formFindPW']").val(findPWType);

	$("#FW_N_Rst_Msg").text("");
	$("#FW_N_Rst").hide();
	$("#FW_N_Form").show();

	$("#FW_N_Btn").show();
	$("#FW_N_Fail_Btn").hide();

	if (findPWType == "mobile") {
		$("#FW_N_Mobile").show();
		$("#FW_N_Email").hide();
	}
	else {
		$("#FW_N_Mobile").hide();
		$("#FW_N_Email").show();
	}
}

/* 아이디/비밀번호찾기 이메일로 찾기 이메일도메인선택 */
function chg_Email(form) {
	var email3 = $("select[name='Email3']", "form[name='" + form + "']").val();
	if (email3 == "@") {
		$("input[name='Email2']", "form[name='" + form + "']").val("");
		$("#Email_Domain2").show();
	}
	else {
		$("input[name='Email2']", "form[name='" + form + "']").val(email3);
		$("#Email_Domain2").hide();
	}
}

/* 아이디찾기 폼 체크 */
function chk_FindID_Normal() {
	var name = alltrim($("input[name='Name']", "form[name='formFindID']").val());
	if (name.length == 0) {
		openAlertLayer("alert", "이름을 입력하여 주십시오.", "closePop('alertPop', 'Name')", "");
		return;
	}

	var findIDType = $("input[name='FindIDType']", "form[name='formFindID']").val();
	if (findIDType == "mobile") {

		var hp2 = alltrim($("input[name='HP2']", "form[name='formFindID']").val());
		if (hp2.length == 0) {
			openAlertLayer("alert", "핸드폰번호를 입력하여 주십시오.", "closePop('alertPop', 'HP2')", "");
			return;
		}
		if (only_Num(hp2) == false) {
			openAlertLayer("alert", "핸드폰번호를 숫자로만 입력하여 주십시오.", "closePop('alertPop', 'HP2')", "");
			return;
		}

		var hp3 = alltrim($("input[name='HP3']", "form[name='formFindID']").val());
		if (hp3.length == 0) {
			openAlertLayer("alert", "핸드폰번호를 입력하여 주십시오.", "closePop('alertPop', 'HP3')", "");
			return;
		}
		if (only_Num(hp3) == false) {
			openAlertLayer("alert", "핸드폰번호를 숫자로만 입력하여 주십시오.", "closePop('alertPop', 'HP3')", "");
			return;
		}

	}
	else {

		var email = alltrim($("input[name='Email']", "form[name='formFindID']").val());
		if (email.length == 0) {
			openAlertLayer("alert", "이메일을 입력하여 주십시오.", "closePop('alertPop', 'Email')", "");
			return;
		}

	}


	$.ajax({
		type		 : "post",
		url			 : "/ASP/Member/Ajax/FindIDOk.asp",
		async		 : false,
		data		 : $("#formFindID").serialize(),
		success		 : function (data) {
						var splitData	 = data.split("|||||");
						var result		 = splitData[0];
						var cont		 = splitData[1];

						document.formFindID.reset();
						$("input[name='FindIDType']", "form[name='formFindID']").val($("input:radio[name='FindIDType']:checked").val());

						$("#FI_N_Rst_Msg").text(cont);
						$("#FI_N_Rst").show();
						$("#FI_N_Form").hide();

						if (result == "OK") {
							$("#FI_N_Btn").hide();
							$("#FI_N_Fail_Btn").hide();
							$("#FI_N_Succ_Btn").show();
							return;
						}
						else {
							$("#FI_N_Btn").hide();
							$("#FI_N_Fail_Btn").show();
							$("#FI_N_Succ_Btn").hide();
							return;
						}
		},
		error		 : function (data) {
						openAlertLayer("alert", "처리 도중 오류가 발생하였습니다.", "closePop('alertPop', '')", "");
		}
	});
}

/* 비밀번호 폼 체크 */
function chk_FindPW_Normal() {
	var name = alltrim($("input[name='Name']", "form[name='formFindPW']").val());
	if (name.length == 0) {
		openAlertLayer("alert", "이름을 입력하여 주십시오.", "closePop('alertPop', 'Name1')", "");
		return;
	}

	var userID = alltrim($("input[name='UserID']", "form[name='formFindPW']").val());
	if (userID.length == 0) {
		openAlertLayer("alert", "아이디를 입력하여 주십시오.", "closePop('alertPop', 'UserID1')", "");
		return;
	}

	var findPWType = $("input[name='FindPWType']", "form[name='formFindPW']").val();
	if (findPWType == "mobile") {

		var hp2 = alltrim($("input[name='HP2']", "form[name='formFindPW']").val());
		if (hp2.length == 0) {
			openAlertLayer("alert", "핸드폰번호를 입력하여 주십시오.", "closePop('alertPop', 'HP2')", "");
			return;
		}
		if (only_Num(hp2) == false) {
			openAlertLayer("alert", "핸드폰번호를 숫자로만 입력하여 주십시오.", "closePop('alertPop', 'HP2')", "");
			return;
		}
		var hp3 = alltrim($("input[name='HP3']", "form[name='formFindPW']").val());
		if (hp3.length == 0) {
			openAlertLayer("alert", "핸드폰번호를 입력하여 주십시오.", "closePop('alertPop', 'HP3')", "");
			return;
		}
		if (only_Num(hp3) == false) {
			openAlertLayer("alert", "핸드폰번호를 숫자로만 입력하여 주십시오.", "closePop('alertPop', 'HP3')", "");
			return;
		}

	}
	else {

		var email1 = alltrim($("input[name='Email']", "form[name='formFindPW']").val());
		if (email1.length == 0) {
			openAlertLayer("alert", "이메일을 입력하여 주십시오.", "closePop('alertPop', 'Email')", "");
			return;
		}
	}


	$.ajax({
		type		 : "post",
		url			 : "/ASP/Member/Ajax/FindPWOk.asp",
		async		 : false,
		data		 : $("#formFindPW").serialize(),
		success		 : function (data) {
						var splitData	 = data.split("|||||");
						var result		 = splitData[0];
						var cont		 = splitData[1];

						document.formFindPW.reset();
						$("input[name='FindPWType']", "form[name='formFindPW']").val($("input:radio[name='FindPWType']:checked").val());

						if (result == "OK") {
							$(".tab-panel").removeClass("active");
							$(".tab-panel1").show();
							return;
						}
						else {
							$("#FW_N_Rst_Msg").text(cont);
							$("#FW_N_Rst").show();
							$("#FW_N_Form").hide();

							$("#FW_N_Btn").hide();
							$("#FW_N_Fail_Btn").show();
							return;
						}
		},
		error		 : function (data) {
						openAlertLayer("alert", "처리 도중 오류가 발생하였습니다.", "closePop('alertPop', '')", "");
		}
	});
}


/* 아이디찾기 핸드폰인증 결과 */
function msg_FindID_AuthHP_Result(result, cont) {
	$("#FI_HP_Rst_Msg").text(cont);
	$("#FI_HP_Rst").show();

	if (result == "OK") {
		$("#FI_HP_Btn").hide();
		$("#FI_HP_Fail_Btn").hide();
		$("#FI_HP_Succ_Btn").show();
		return;
	}
	else {
		$("#FI_HP_Btn").hide();
		$("#FI_HP_Fail_Btn").show();
		$("#FI_HP_Succ_Btn").hide();
		return;
	}
}


/* 아이디찾기 핸드폰인증 다시찾기 */
function re_FindID_AuthHP() {
	$("#FI_HP_Rst_Msg").text("");
	$("#FI_HP_Rst").hide();

	$("#FI_HP_Btn").show();
	$("#FI_HP_Fail_Btn").hide();
	$("#FI_HP_Succ_Btn").hide();
}


/* 아이디찾기 아이핀인증 결과 */
function msg_FindID_AuthIpin_Result(result, cont) {
	$("#FI_Ipin_Rst_Msg").text(cont);
	$("#FI_Ipin_Rst").show();

	if (result == "OK") {
		$("#FI_Ipin_Btn").hide();
		$("#FI_Ipin_Fail_Btn").hide();
		$("#FI_Ipin_Succ_Btn").show();
		return;
	}
	else {
		$("#FI_Ipin_Btn").hide();
		$("#FI_Ipin_Fail_Btn").show();
		$("#FI_Ipin_Succ_Btn").hide();
		return;
	}
}


/* 아이디찾기 아이핀인증 다시찾기 */
function re_FindID_AuthIpin() {
	$("#FI_Ipin_Rst_Msg").text("");
	$("#FI_Ipin_Rst").hide();

	$("#FI_Ipin_Btn").show();
	$("#FI_Ipin_Fail_Btn").hide();
	$("#FI_Ipin_Succ_Btn").hide();
}


/* 비밀번호찾기 핸드폰인증 결과 */
function msg_FindPW_AuthHP_Result(result, cont) {
	if (result == "OK") {
		$(".tab-panel").removeClass("active");
		$(".tab-panel1").show();
		return;
	}
	else {
		$("#FW_HP_Rst_Msg").text(cont);
		$("#FW_HP_Rst").show();
		$("#FW_HP_Form").hide();

		$("#FW_HP_Btn").hide();
		$("#FW_HP_Fail_Btn").show();
		return;
	}
}


/* 비밀번호찾기 핸드폰인증 다시찾기 */
function re_FindPW_AuthHP() {
	document.form.reset();

	$("#FW_HP_Rst_Msg").text("");
	$("#FW_HP_Rst").hide();
	$("#FW_HP_Form").show();

	$("#FW_HP_Btn").show();
	$("#FW_HP_Fail_Btn").hide();

	$(".tab-panel1").hide();
}


/* 비밀번호찾기 아이핀 결과 */
function msg_FindPW_AuthIpin_Result(result, cont) {
	if (result == "OK") {
		$(".tab-panel").removeClass("active");
		$(".tab-panel1").show();
		return;
	}
	else {
		$("#FW_Ipin_Rst_Msg").text(cont);
		$("#FW_Ipin_Rst").show();
		$("#FW_Ipin_Form").hide();

		$("#FW_Ipin_Btn").hide();
		$("#FW_Ipin_Fail_Btn").show();
		return;
	}
}


/* 비밀번호찾기 아이핀인증 다시찾기 */
function re_FindPW_AuthIpin() {
	document.form1.reset();

	$("#FW_Ipin_Rst_Msg").text("");
	$("#FW_Ipin_Rst").hide();
	$("#FW_Ipin_Form").show();

	$("#FW_Ipin_Btn").show();
	$("#FW_Ipin_Fail_Btn").hide();

	$(".tab-panel1").hide();
}


/* 비밀번호찾기 인증완료 후 비밀번호 입력 초기화 */
function intTabPanel1(){
	$(".tab-panel1").hide();
}



/* 비밀번호 수정 체크 */
function chk_ChgPwd() {
	/* 비밀번호 */
	var pwd = alltrim($("input[name='Pwd']", "form[name='formChgPwd']").val());
	if (pwd.length < 6 || pwd.length > 12) {
		openAlertLayer("alert", "비밀번호를 영문, 숫자조합 6자리이상 12자리 이내로 입력해 주십시오.", "closePop('alertPop', 'newPw')", "");
		return;
	}
	if (only_AlphaNum(pwd) == false) {
		openAlertLayer("alert", "비밀번호를 영문, 숫자조합 6자리이상 12자리 이내로 입력해 주십시오.", "closePop('alertPop', 'newPw')", "");
		return;
	}

	/* 비밀번호 확인 */
	var pwd1 = alltrim($("input[name='Pwd1']", "form[name='formChgPwd']").val());
	if (pwd1.length == 0) {
		openAlertLayer("alert", "비밀번호를 다시 한번 입력해 주십시오.", "closePop('alertPop', 'newPwCheck')", "");
		return;
	}

	if (pwd != pwd1) {
		openAlertLayer("alert", "비밀번호가 일치하지 않습니다.", "closePop('alertPop', 'newPwCheck')", "");
		return;
	}

	openAlertLayer("confirm", "비밀번호를 수정 하시겠습니까?", "closePop('confirmPop', '')", "closePop('confirmPop', '');chk_ChgPwdOk();");
}


function chk_ChgPwdOk() {
	$.ajax({
		type		 : "post",
		url			 : "/ASP/Member/Ajax/PwdModifyOk.asp",
		async		 : false,
		data		 : $("#formChgPwd").serialize(),
		dataType	 : "text",
		success		 : function (data) {
						var splitData	 = data.split("|||||");
						var result		 = splitData[0];
						var cont		 = splitData[1];


						if (result == "OK") {
							openAlertLayer("alert", "수정 되었습니다.", "closePop('alertPop', '');APP_HistoryBack();", "");
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

/* 휴대폰인증으로 비밀번호 찾기 폼 체크 */
function chk_FindPW_AuthHP(form) {
	var name = alltrim($("input[name='Name']", "form[name='form']").val());
	if (name.length == 0) {
		openAlertLayer("alert", "이름을 입력하여 주십시오.", "closePop('alertPop', 'Name2');", "");
		return;
	}

	var userID = alltrim($("input[name='UserID']", "form[name='form']").val());
	if (userID.length == 0) {
		openAlertLayer("alert", "아이디를 입력하여 주십시오.", "closePop('alertPop', 'UserID2');", "");
		return;
	}

	auth_HP(form);
}


/* 아이핀인증으로 비밀번호 찾기 폼 체크 */
function chk_FindPW_AuthIpin(form) {
	var name = alltrim($("input[name='Name']", "form[name='form1']").val());
	if (name.length == 0) {
		openAlertLayer("alert", "이름을 입력하여 주십시오.", "closePop('alertPop', 'Name3');", "");
		return;
	}

	var userID = alltrim($("input[name='UserID']", "form[name='form1']").val());
	if (userID.length == 0) {
		openAlertLayer("alert", "아이디를 입력하여 주십시오.", "closePop('alertPop', 'UserID3');", "");
		return;
	}

	auth_Ipin(form);
}


/* 네이버 로그인 */
function pop_NaverLogin() {
	APP_PopupGoUrl("/API/Naver.asp", '0', '');
}

/* 구글 로그인 */
function pop_GoogleLogin() {
	APP_PopupGoUrl("/API/Google.asp", '0', '');
}

/* 카카오 로그인 */
function pop_KakaoLogin() {
	APP_PopupGoUrl("/API/Kakao.asp", '0', '');
}

/* 페이스북 로그인 */
function pop_FacebookLogin() {
	APP_PopupGoUrl("/API/Facebook.asp", '0', '');
}



/* SNS로그인 사용자 체크 */
function snsLogin() {
	$.ajax({
		url			 : '/ASP/Member/Ajax/SnsLoginOk.asp',
		data		 : $("form[name='SimpleLoginForm']").serialize(),
		async		 : false,
		type		 : 'post',
		dataType	 : 'html',
		success		 : function (data) {
						var splitData	 = data.split("|||||");
						var result		 = splitData[0];
						var msg			 = splitData[1];

						if (result == "OK") {
							location.href = $("#ProgID").val();
						}
						else if (result == "FAIL") {
							openAlertLayer("alert", msg, "closePop('alertPop', '')", "");
							return;
						}
						else if (result == "FAIL_LOGIN") {
							openAlertLayer("alert", msg, "closePop('alertPop', '');", "");
							return;
						}
						else if (result == "NEWAGREE") {
							openAlertLayer("confirm", "신규 약관에 동의하여 주십시오.", "closePop('confirmPop', '')", "closePop('confirmPop', '');APP_GoUrl('/ASP/Member/NewAgreement.asp');");
							return;
						}
						else if (result == "DORMANCY") {
							APP_GoUrl("/ASP/Member/DormancyRelease.asp");
							return;
						}
						else {
							openAlertLayer("alert", msg, "closePop('alertPop', '')", "");
							return;
						}
		},
		error		 : function (data) {
						openAlertLayer("alert", "처리 도중 오류가 발생하였습니다.", "closePop('alertPop', '')", "");
		}
	});
}

/* MYPAGE > SNS계정연결 */
function snsConnection() {
	$.ajax({
		url			 : '/ASP/Mypage/Ajax/MySnsConnection.asp',
		data		 : $("form[name='SimpleLoginForm']").serialize(),
		async		 : false,
		type		 : 'post',
		dataType	 : 'html',
		success		 : function (data) {
						var splitData	 = data.split("|||||");
						var result		 = splitData[0];
						var msg			 = splitData[1];

						if (result == "OK") {
							openAlertLayer("alert", "SNS연결 설정이 되었습니다.", "closePop('alertPop', '');location.reload();", "");
						}
						else if (result == "FAIL") {
							openAlertLayer("alert", msg, "closePop('alertPop', '');", "");
							return;
						}
						else if (result == "DIDUP") {	/* 연결하려는 SNS계정의 회원이 있을 경우 통합 */
							openAlertLayer("confirm", msg, "closePop('confirmPop', '');", "closePop('confirmPop', '');MemSnsCombine();");
							return;
						}
						else {
							openAlertLayer("alert", "Error : " + msg, "closePop('alertPop', '');", "");
							return;
						}
		},
		error		 : function (data) {
						openAlertLayer("alert", "계정연결 처리 중 오류가 발생하였습니다.", "closePop('alertPop', '');", "");
		}
	});
}

/* MYPAGE > SNS계정연결시 해당 SNS회원 정보가 있을 경우 통합처리 */
function MemSnsCombine() {
	$.ajax({
		url			 : '/ASP/Mypage/Ajax/MemSnsCombineOk.asp',
		data		 : $("form[name='SimpleLoginForm']").serialize(),
		async		 : false,
		type		 : 'post',
		dataType	 : 'html',
		success		 : function (data) {
						var splitData	 = data.split("|||||");
						var result		 = splitData[0];
						var msg			 = splitData[1];

						if (result == "OK") {
							openAlertLayer("alert", msg, "closePop('alertPop', '');location.reload();", "");
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



/* 비회원 주문조회 로그인 */
function chk_NLogin() {

	var name = alltrim($("input[name='Name']", "form[name='formNLogin']").val());
	if (name.length == 0) {
		openAlertLayer("alert", "이름을 입력해 주십시오.", "closePop('alertPop', 'Name');", "");
		return;
	}

	var hp1 = $("select[name='HP1']", "form[name='formNLogin']").val();
	if (hp1.length == 0) {
		openAlertLayer("alert", "휴대폰번호를 선택해 주십시오.", "closePop('alertPop', 'HP1');", "");
		return;
	}

	var hp2 = $("input[name='HP2']", "form[name='formNLogin']").val();
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


	var hp3 = $("input[name='HP3']", "form[name='formNLogin']").val();
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

	var email = alltrim($("input[name='Email']", "form[name='formNLogin']").val());
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


	$.ajax({
		type		 : "post",
		url			 : "/ASP/Mypage/Ajax/NLoginOk.asp",
		async: false,
		data		 : $("#formNLogin").serialize(),
		success		 : function (data) {
						var splitData	 = data.split("|||||");
						var result		 = splitData[0];
						var cont		 = splitData[1];

						if (result == "OK") {
							location.href = "/ASP/Mypage/OrderList.asp";
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



/* 로그아웃 */
function sm_Logout() {
	$.ajax({
		url			 : '/ASP/Member/Ajax/Logout.asp',
		async		 : false,
		type		 : 'post',
		dataType	 : 'html',
		success		 : function (data) {
						openAlertLayer("alert", "로그아웃 되었습니다.", "closePop('alertPop', '');location.href='/';", "");
		},
		error		 : function (data) { }
	});
}

//에이스 카운터 로그인
function AceCounter_Login(age, gender, uid) {
	var m_ag = age;         // 로그인사용자 나이
	var m_id = uid;    		// 로그인사용자 아이디
	var m_gd = gender;         // 로그인사용자 성별 ('man' , 'woman')

	var _AceGID = (function () { var Inf = ['app.shoemarker.co.kr', 'app.shoemarker.co.kr', 'AZ1A74686', 'AM', '0', 'NaPm,Ncisy', 'ALL', '0']; var _CI = (!_AceGID) ? [] : _AceGID.val; var _N = 0; if (_CI.join('.').indexOf(Inf[3]) < 0) { _CI.push(Inf); _N = _CI.length; } return { o: _N, val: _CI }; })();
	var _AceCounter = (function () { var G = _AceGID; var _sc = document.createElement('script'); var _sm = document.getElementsByTagName('script')[0]; if (G.o != 0) { var _A = G.val[G.o - 1]; var _G = (_A[0]).substr(0, _A[0].indexOf('.')); var _C = (_A[7] != '0') ? (_A[2]) : _A[3]; var _U = (_A[5]).replace(/\,/g, '_'); _sc.src = (location.protocol.indexOf('http') == 0 ? location.protocol : 'http:') + '//cr.acecounter.com/Mobile/AceCounter_' + _C + '.js?gc=' + _A[2] + '&py=' + _A[1] + '&up=' + _U + '&rd=' + (new Date().getTime()); _sm.parentNode.insertBefore(_sc, _sm); return _sc.src; } })();
}

