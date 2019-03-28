var agent		 = navigator.userAgent.toLowerCase();
var isIOS		 = (agent.indexOf("iphone") > -1 || agent.indexOf("ipad") > -1 || agent.indexOf("ipod") > -1);
var isANDROID	 = (agent.match('android') != null);
var isMSIE80	 = (agent.indexOf("msie 6.0") > -1 || agent.indexOf("msie 7.0") > -1 || agent.indexOf("msie 8.0") > -1);


/* 디버깅 모드 */
var APP_Debug_Flag = false;


/* 토스트카운트 */
var mflag = 0;


var app_msg = "";


function APP_BadgeUpdate(i) {
	toApp({ method: "updateBadge", num: i });
}

function APP_TopGo() {
	toApp({ method: "clearTop", url: home_domain });
}

function APP_TopGoUrl(targetUrl) {
	if (isApp == 'Y')	 { toApp({ method: "clearTop", url: home_domain + targetUrl }); }
	else				 { location.href = home_domain + targetUrl; }
}

function APP_GoUrl(targetUrl) {
	if (isApp == 'Y')	 { toApp({ method:"openWindow", url: home_domain + targetUrl }); }
	else				 { location.href = home_domain + targetUrl; }
}


function APP_PopupGoUrl(targetUrl, urlType, popOption) {
	if (isApp == 'Y') {
		if (urlType == '0') {
			if (isIOS)	 { toApp({ method: "openPop", url: home_domain + targetUrl }); }
			else		 { toApp({ method: "openPop", url: home_domain + targetUrl, external: true }); }
		}
		else {
			if (isIOS)	 { toApp({ method: "openPopNoBtn", url: home_domain + targetUrl }); }
			else		 { toApp({ method: "openPopNoBtn", url: home_domain + targetUrl, external: false }); }
		}
	}
	else {
		window.open(targetUrl);
	}
}

function toApp(param) {
	location.href = "app://" + encodeURIComponent(JSON.stringify(param));
}

function openExternal(goUrl) {
	toApp({ method:"openExternal", url:goUrl });
}

function finish() {
	toApp({ method: "finish", });
}

function sendPush(mnum, tit, cont, idx) {
	toApp({ method: "sendPush", memberNum: mnum, title: tit, alert: cont, pushIdx: idx });
}

function kakao_share(par, tit, img, description, tp, sp, dc, btn) {
	//toApp({ method: "kakaoLink", param: "ProductCode=10637", title: "스노우블락", image: "http://m.shoemarker.co.kr/Upload/ProductImage/020104/10637_1_0500_0500.jpg", desc: "슈마커", regularPrice: "39000", discountPrice: "29000", discountRate: "10", buttonLabel: "상품보러가기" });
	toApp({ method: "kakaoLink", param: par, title: tit, image: img, desc: description, regularPrice: parseInt(tp), discountPrice: parseInt(sp), discountRate: parseInt(dc), buttonLabel: btn });
}

var app = {
				installationId: function ($installationId, $deviceToken, $deviceType, $appVersion, $appModel) {
					var url = document.URL;
					url = url.toLowerCase();

					if (url.search(home_domain + "/gate.asp") >= 0) {
						APP_DeviceInfoAdd($installationId, $deviceToken, $deviceType, $appVersion, $appModel, "0");
					} 
					else {
						APP_DeviceInfoAdd($installationId, $deviceToken, $deviceType, $appVersion, $appModel, "1");
					}

				},
				appVersion: function ($appVersion) { },
				popupResult: function ($result) {
					var result = $result.split("|");

					var my_url = document.URL;
					my_url = my_url.toLowerCase();


					/* 창닫히고 부모창이 메인 페이지일 경우 - 메인 페이지 신규 처방전 수 */
					if ((my_url == home_domain) || (my_url == home_domain + "/") || (my_url.search(home_domain + "/index.asp")) >= 0) {
						//get_NewPrescription();
						//get_BadgeCount();
						//return;
					}
					/* 창닫히고 부모창이 처방전 리스트 페이지일 경우 */
					else if (my_url.indexOf("/prescriptionlist.asp") >= 0) {
						//historyBack();
					}
					/* 창닫히고 부모창이 1:1문의 리스트 페이지일 경우 */
					else if (my_url.indexOf("/inquirylist.asp") >= 0) {
						//historyBack();
					}

					if (result[0] == "afterlogin") {									/* 로그인 후 페이지 이동 */
						move_AfterLogin();
					}

					else if (result[0] == "aftersnslogin") {							/* SNS 로그인 후 페이지 이동 */
						move_AfterSnsLogin(result[1]);
					}

					else if (result[0] == "ordersnsconnect") {							/* 주문시 SNS 로그인 후 기존 아이디 연결 */
						move_AfterSnsConnect();
					}

					else if (result[0] == "alertlayer") {								/* 팝업 닫고 alert 띄우기 */
						openAlertLayer("alert", result[1], "closePop('alertPop', '');", "");
					}

					else if (result[0] == "move") {										/* 팝업 닫고 페이지 이동 */
						move_Page(result[1]);	/* dev_common.js */
					}

					else if (result[0] == "find_id_hp_result") {						/* 아이디찾기 핸드폰 인증 결과 */
						msg_FindID_AuthHP_Result(result[1], result[2]);
					}

					else if (result[0] == "find_id_ip_result") {						/* 아이디찾기 아이핀 인증 결과 */
						msg_FindID_AuthIpin_Result(result[1], result[2]);
					}

					else if (result[0] == "find_pw_hp_result") {						/* 비밀번호찾기 핸드폰 인증 결과 */
						msg_FindPW_AuthHP_Result(result[1], result[2]);
					}

					else if (result[0] == "find_pw_ip_result") {						/* 비밀번호찾기 아이핀 인증 결과 */
						msg_FindPW_AuthIpin_Result(result[1], result[2]);
					}

					else if (result[0] == "dormancy_result") {							/* 휴명계정해제 결과 */
						after_DormancyAuth(result[1], result[2]);
					}

					else if (result[0] == "join_auth_result") {							/* 회원가입 인증 결과 */
						after_JoinAuth(result[1], result[2]);
					}

					else if (result[0] == "login") {
						if (result[1] == "/") {
							location.replace("/Index.asp");
						}
						else {
							location.replace("/Index.asp");
							APP_GoUrl(result[1]);
						}
					}
					else if (result[0] == "main") {
						location.replace("/Index.asp");
					}
					else if (result[0] == "WriteOK") {
						location.reload();
					}
					else if (result[0] == "list") {
						location.reload();
					}
					else if (result[0] == "returnUrl") {
						if (result[1] != "") {
							location.replace(result[1]);
						}
					}
					else if (result[0] == "OrderConfirm") {
						//location.replace(result[1]);

						APP_TopGoUrl(result[1]);
					}
					else if (result[0] == "Cart") {
						location.replace("/index.asp");
					}
					else if (result[0] == "other") { }

					get_GNB_CartCount();
				},


				addImage: function ($path, $img) {
					//alert("kkkkkkkk");
					//alert("$img=" + $img);


					document.getElementById("PPath").value = $path;
					document.getElementById("PImg").value = $img;
					alert($img)
					/*
					var url = document.URL;
					url = url.toLowerCase();

					// 본인 사진변경일 경우
					if (url.indexOf("/mypage/") > -1) {
						imgChg($path);
					}else{
						var img = new Image;
						img.src = $img;
						img.path = $path;
						img.height = 100;
						document.getElementById("img_list").appendChild(img);
						img.onclick = removethis;
					}
					*/
				},

				/* 사진 처방전 저장하는 중 */
				uploadStart: function () {
					$(".dim").show();
					$("#loading").show();
				},

				uploadComplete: function ($result) {
					if ($result.indexOf("|||||") > 0) {
						var splitData	 = $result.split("|||||");
						var result		 = splitData[0];
						var cont		 = splitData[1];

						if (result == "OK") {
							$(".dim").hide();
							$("#loading").hide();
							get_NewPrescription();
							openAlertLayer('alertPop', 'alert', "등록 되었습니다.", 'closeAlertLayer("alertPop");APP_GoUrl("/ASP/Prescription/PrescriptionList.asp");', '');
						}
						else {
							$(".dim").hide();
							$("#loading").hide();
							openAlertLayer('alertPop', 'alert', cont, 'closeAlertLayer("alertPop");', '');
						}
					}
					else {
						$(".dim").hide();
						$("#loading").hide();
						openAlertLayer('alertPop', 'alert', "처리중 오류가 발생하였습니다.", 'closeAlertLayer("alertPop");', '');
					}
				},


				backKey: function () {
					// 열린 레이어팝업 찾기
					var popCnt = 0;
					$(".popup").each(function () {
						if ($(this).css("display") != "none") {
							popCnt += 1;
						}
					});
					// 열린 레이어팝업이 있으면 닫기
					if (popCnt > 0) {
						closePop();
						return;
					}

					var url = document.URL;
					url = url.toLowerCase();

					if (
							(url == home_domain) || (url == home_domain + "/") || (url.search(home_domain + "/index.asp")) >= 0 
								|| (url.search(home_domain + "/asp/member/login.asp") >= 0) 
						)
					{
						// 열린 레이어팝업 찾기
						var popID = "";
						//$("[id^=PopDiv]").each(function () {
						//	if ($(this).css("display") != "none") {
						//		popID = $(this).attr("id");
						//	}
						//});
						// 열린 레이어팝업이 있으면 닫기
						if (popID != "") {
							$("." + popID + "Close").trigger("click");

						} else {
							//메인일때
							if (new Date() - mflag > 2000) {
								mflag = new Date() * 1;
								toApp({
									method: "toast",
									msg: "'뒤로'버튼을 한번 더 누르시면 종료 됩니다."
								});
							} else {
								toApp({
									method: "finish",
								});
							}
						}
					}


					/*
					else {
						toApp({ method: "finish", });
					}
					*/
				},


				doneKey: function () {
					var url = document.URL;
					url = url.toLowerCase();
					if ((url == home_domain) || (url == home_domain + "/") || (url == home_domain + "/index.asp") || (url == home_domain + "/asp/member/login.asp") ) {
						//메인일때
						if (new Date() - mflag > 2000) {
							mflag = new Date() * 1;
							toApp({ method: "toast", msg: "'뒤로'버튼을 한번 더 누르시면 종료 됩니다." });
						}
						else {
							toApp({ method: "finish", });
						}
					}
					else {
						//메인이 아닐때
						toApp({ method: "popupResult", result: "other|0" });
					}
				},


				resume: function () {
					//alert("살아남!");
				},

				userLocation: function(lat, long) {
					set_MyLocation(lat, long);
				}
};

/* LOGIN 후 페이지 이동 */
function APP_HistoryBack_Login() {
	if (isApp == 'Y')	 { toApp({ method: "popupResult", result: "afterlogin|" }); }
	else				 { opener.move_AfterLogin(); self.close(); }
}

/* SNS LOGIN 후 페이지 이동 */
function APP_HistoryBack_SNS_Login(val) {
	if (isApp == 'Y')	 { toApp({ method: "popupResult", result: "aftersnslogin|" + val }); }
	else				 { opener.move_AfterSnsLogin(val); self.close(); }
}

/* 주문 로그인 SNS LOGIN 후 기존 아이디 연결 - 로그인 페이지 다시 이동 */
function APP_PopupHistoryBack_Order_Sns_Connect() {
	if (isApp == 'Y')	 { toApp({ method: "popupResult", result: "ordersnsconnect|" }); }
	else				 { opener.move_AfterSnsConnect(); self.close(); }
}




/* 팝업 닫고 alert 띄우기 */
function APP_PopupHistoryBack_Alert(msg) {
	if (isApp == 'Y')	 { toApp({ method: "popupResult", result: "alertlayer|" + msg }); }
	else				 { opener.openAlertLayer("alert", msg, "closePop('alertPop', '');", ""); self.close(); }
}

/* 팝업 닫고 이동 */
function APP_PopupHistoryBack_Move(move_url) {
	if (isApp == 'Y')	 { toApp({ method: "popupResult", result: "move|" + move_url }); }
	else				 { opener.move_Page(move_url); self.close(); }
}

/* 아이디 찾기 핸드폰 인증 결과 */
function APP_PopupHistoryBack_ID_HP_Result(rCode, rMessage) {
	if (isApp == 'Y')	 { toApp({ method: "popupResult", result: "find_id_hp_result|" + rCode + "|" + rMessage }); }
	else				 { opener.msg_FindID_AuthHP_Result(rCode, rMessage); self.close(); }
}

/* 아이디 찾기 아이핀 인증 결과 */
function APP_PopupHistoryBack_ID_IP_Result(rCode, rMessage) {
	if (isApp == 'Y')	 { toApp({ method: "popupResult", result: "find_id_ip_result|" + rCode + "|" + rMessage }); }
	else				 { opener.msg_FindID_AuthIpin_Result(rCode, rMessage); self.close(); }
}

/* 비밀번호 찾기 핸드폰 인증 결과 */
function APP_PopupHistoryBack_PW_HP_Result(rCode, rMessage) {
	if (isApp == 'Y')	 { toApp({ method: "popupResult", result: "find_pw_hp_result|" + rCode + "|" + rMessage }); }
	else				 { opener.msg_FindPW_AuthHP_Result(rCode, rMessage); self.close(); }
}

/* 비밀번호 찾기 아이핀 인증 결과 */
function APP_PopupHistoryBack_PW_IP_Result(rCode, rMessage) {
	if (isApp == 'Y')	 { toApp({ method: "popupResult", result: "find_pw_ip_result|" + rCode + "|" + rMessage }); }
	else				 { opener.msg_FindPW_AuthIpin_Result(rCode, rMessage); self.close(); }
}

/* 휴면게정 처리 결과 */
function APP_PopupHistoryBack_DOR_Result(rCode, rMessage) {
	if (isApp == 'Y')	 { toApp({ method: "popupResult", result: "dormancy_result|" + rCode + "|" + rMessage }); }
	else				 { opener.after_DormancyAuth(rCode, rMessage); self.close(); }
}

/* 회원가입 인증 처리 결과 */
function APP_PopupHistoryBack_JoinAuth_Result(rCode, rMessage) {
	if (isApp == 'Y')	 { toApp({ method: "popupResult", result: "join_auth_result|" + rCode + "|" + rMessage }); }
	else				 { opener.after_JoinAuth(rCode, rMessage); self.close(); }
}







function APP_HistoryBack() {
	if (isApp == 'Y')	 { toApp({ method: "popupResult", result: "other|0" }); }
	else				 { history.back(); }
}

function APP_PopupHistoryBack() {
	if (isApp == 'Y')	 { toApp({ method: "popupResult", result: "other|0" }); }
	else				 { self.close(); }
}

function APP_HistoryBack_View() {
	toApp({ method:"popupResult", result:"list|0" });
}

function APP_HistoryBack_Url(returnUrl) {
	toApp({ method: "popupResult", result: "returnUrl|" + returnUrl });
}

function APP_goMain() {
	toApp({ method:"popupResult", 	result:"main|main.asp" });
}

function LoginConfirm(targetUrl) {
	toApp({
		method:"popupResult",
		result:"login|"+targetUrl
	});
}

function WriteConfirm(targetUrl)
{
	toApp({
		method:"popupResult",
		result:"WriteOK|"+targetUrl
	});
}

function OrderConfirm(targetUrl)
{
	toApp({
		method:"popupResult",
		result:"OrderConfirm|"+targetUrl
	});
}

function main_Replace() { }

function Send_SMS(body)
{
	toApp({
		method:"SendSMS",
		sms:body
	});
}

/* 디버그용 ALERT 함수 */
function APP_Debug(str) {
	if (APP_Debug_Flag) { /*alert(str);*/ }
}

function log($str) {
	document.getElementById("logs").innerHTML += $str + "<br/>";
}

/* 초기 접속시 앱 디바이스 정보 입력 처리 */
function APP_DeviceInfoAdd(installationId, deviceToken, deviceType, appVersion, appModel, t) {

	$.ajax({
		type		 : "get",
		url			 : "/Common/Ajax/DeviceInfo.asp",
		async		 : true,
		data		 : "installationId=" + installationId + "&deviceToken=" + deviceToken + "&deviceType=" + deviceType + "&appVersion=" + appVersion + "&appModelName=" + appModel,
		dataType	 : "text",
		success		 : function (data) {
						location.replace('/?intro=Y');
		},
		error		 : function (data) {
						location.replace('/?intro=Y');// alert(data.responseText)
		}
	});
}


/* 카메라 호출 */
function APP_AddPrescription() {
	toApp({ method: "openCamera", userId: u_no, uploadUrl: home_domain + "/ASP/Prescription/Ajax/PrescriptionAddOk.asp" });
}


/* 내 위치 좌표 */
function APP_UserLocation() {
	if (isApp == "Y")	 { toApp({ method: "getUserLocation" }); }
	else				 { set_MyLocation(37.4852224, 126.8949784); }
}


/* 검색 */
function LinkgoUrl(url) {
	APP_GoUrl(encodeURI(url));
}