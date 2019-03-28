<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'DormancyRelease.asp - 휴면계정해제 - 본인인증
'Date		: 2018.12.19
'Update		:
'*****************************************************************************************'

'# 페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'-----------------------------------------------------------------------------------------'

'# 페이지 코드----------------------------------------------------------------------------'
DIM PageCode1, PageCode2, PageCode3, PageCode4
PageCode1 = "01"
PageCode2 = "03"
PageCode3 = "01"
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->

<%
IF U_ID <> "" THEN
		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=&Msg=&Script=APP_TopGoUrl('/');"
		Response.End
END IF


'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn						'# ADODB Connection 개체
DIM oRs							'# ADODB Recordset 개체
DIM oCmd						'# ADODB Command 개체

DIM i
DIM j
DIM x
DIM y

DIM wQuery						'# WHERE 절
DIM sQuery						'# SORT 절


DIM UserID
DIM UserName
DIM DormancyFlag
DIM NewAgreementFlag
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


UserID			 = TRIM(Decrypt(Request.Cookies("TEMP_UID")))
UserName		 = TRIM(Decrypt(Request.Cookies("TEMP_UNAME")))
DormancyFlag	 = TRIM(Decrypt(Request.Cookies("TEMP_DOR")))
NewAgreementFlag = TRIM(Decrypt(Request.Cookies("TEMP_NEW")))


IF DormancyFlag = "N" THEN
		IF NewAgreementFlag = "N" THEN
				Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=신규 약관에 동의하여 주십시오.&Script=location.href='/ASP/Member/NewAgreement.asp';"
				Response.End
		ELSE
				Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=다시 로그인하여 주십시오.&Script=APP_HistoryBack();"
				Response.End
		END IF
END IF

	

SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>


<!-- #include virtual="/INC/Header.asp" -->
	<script type="text/javascript" src="/JS/dev/join.js?ver=<%=U_DATE%><%=U_TIME%>"></script>
	<script type="text/javascript">
		function after_DormancyAuth(rCode, rMessage) {
			if (rCode == "OK") {
				openAlertLayer("alert", rMessage, "closePop('alertPop', '');move_Main();", "");
			}
			else if (rCode == "OK2") {
				var splitMsg	 = rMessage.split("///");
				var name		 = splitMsg[0];
				var couponName	 = splitMsg[1];
				var progID		 = splitMsg[2];
				var orderFlag	 = splitMsg[3];
				var msg			 = "";

				msg = msg + "			<div class='area-pop'>";
				msg = msg + "				<div class='alert'>";
				msg = msg + "					<div class='tit-pop'>";
				msg = msg + "						<p class='tit' id='confirm_title'>SHOEMARKER</p>";
				msg = msg + "						<button id='confirm_close' onclick=\"closePop('messagePop')\" class='btn-hide-pop'>닫기</button>";
				msg = msg + "					</div>";
				msg = msg + "					<div class='container-pop'>";
				msg = msg + "						<div class='contents'>";
				msg = msg + "							<div class='ly-cont'>";
				msg = msg + "								<p id='confirm_content' class='t-level4' style='text-align:left'>";
				msg = msg +									"휴면계정 해제 처리 되었습니다.<br>";
				msg = msg +									name + "님께<br>";
				msg = msg +									couponName;
				msg = msg + "								</p>";
				msg = msg + "							</div>";
				msg = msg + "						</div>";
				msg = msg + "						<div class='btns'>";

				if (orderFlag != "Y") {
					msg = msg + "							<button type='button' id='message_btn1' onclick=\"APP_TopGoUrl('/ASP/Mypage/CouponList.asp');\" class='button ty-black'>쿠폰 확인</button>";
				}

				msg = msg + "							<button type='button' id='message_btn2' onclick=\"APP_TopGoUrl('" + progID + "');\" class='button ty-red'>확인</button>";
				msg = msg + "						</div>";
				msg = msg + "					</div>";
				msg = msg + "				</div>";
				msg = msg + "			</div>";

				$("#messagePop").html(msg);
				openPop("messagePop")
			}
			else if (rCode == "NEWAGREEMENT") {
				openAlertLayer("alert", rMessage, "closePop('alertPop', '');move_NewAgreement();", "");
			}
			else if (rCode == "DOR_NOTEXISTS") {
				openAlertLayer("alert", rMessage, "closePop('alertPop', '');APP_HistoryBack()", "");
			}
			else {
				openAlertLayer("alert", rMessage, "closePop('alertPop', '');", "");
			}
		}

		function move_NewAgreement() {
			//openPop("loading");
			location.href = "/ASP/Member/NewAgreement.asp";
		}

		function move_Main() {
			//openPop("loading");
			APP_HistoryBack_Login();
		}
	</script>

<%TopSubMenuTitle = "휴면계정해제"%>
<!-- #include virtual="/INC/TopSub.asp" -->


<!-- Main -->
<main id="container" class="container">
	<div class="content">
		<form name="form" id="form">
		<input type="hidden" name="SMode" value="DormancyRelease" />
		<section class="join-agreement">
			<h1 class="h-level1">휴면계정해제</h1>
			<ul class="agreement-list">
				<li>
					<div class="fieldset">
						<label><%=UserName%>(<%=UserID%>) 님은 현재 휴면계정이십니다.<br />본인인증 후 슈마커 쇼핑몰 서비스를 이용 가능하십니다.</label>
					</div>
				</li>
			</ul>
			<div class="agree-info">
				<ul>
					<li class="bullet-ty1">슈마커 쇼핑몰은 입점몰의 원활한 운영을 위하여 이용자의 개인정보 일부를...</li>
					<li class="bullet-ty1">슈마커에서는 “정보통신망 이용촉진 및 정보보호 등에 관한 법률” 제 23의 2 “주민등록번호의 사용 제안” 에 의거 고객님의 주민번호를 수집, 보관, 이용하지 않습니다.</li>
				</ul>
			</div>
			<div class="fieldset confirm-btn">
				<!-- 14세 이상 인증 버튼 -->
				<div class="confirm-adult">
					<a href="javascript:void(0)" onclick="auth_HP('form');" class="button is-expand ty-red">휴대폰 인증</a>
					<a href="javascript:void(0)" onclick="auth_Ipin('form');" class="button is-expand ty-black">아이핀 인증</a>
				</div>
			</div>
		</section>
		</form>
	</div>
</main>

<!-- #include virtual="/INC/FooterNoBNB.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>