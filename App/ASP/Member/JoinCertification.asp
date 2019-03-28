<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'JoinCertification.asp - 회원 가입 - 본인인증
'Date		: 2018.11.29
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
PageCode2 = "02"
PageCode3 = "02"
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

DIM JoinType					'# 14세 구분 (U:14세이상 / D:14세미만)
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


JoinType		 = sqlFilter(Request.Form("JoinType"))
IF JoinType = "" THEN
		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=14세 구분값이 없습니다.&Script=location.href='/ASP/Member/Join.asp';"
		Response.End
END IF


Response.Cookies("JOIN_TYPE")		 = Encrypt(JoinType)


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>


<!-- #include virtual="/INC/Header.asp" -->
	<script type="text/javascript" src="/JS/dev/join.js?ver=<%=U_DATE%><%=U_TIME%>"></script>
	<script type="text/javascript">
		function after_JoinAuth(rCode, rMessage) {
			if (rCode == "OK") {
				goJoin();
			}
			else if (rCode == "UNDER14" || rCode == "UNDER20") {
				openAlertLayer("alert", rMessage, "closePop('alertPop', '');move_Join();", "");
			}
			else if (rCode == "MEMBER") {
				openAlertLayer("alert", rMessage, "closePop('alertPop', '');APP_HistoryBack()", "");
			}
			else {
				openAlertLayer("alert", rMessage, "closePop('alertPop', '');", "");
			}
		}

		function move_Join() {
			//openPop("loading");
			location.href = "/ASP/Member/Join.asp";
		}
	</script>

<%TopSubMenuTitle = "회원가입"%>
<!-- #include virtual="/INC/TopSub.asp" -->


<!-- Main -->
<main id="container" class="container">
	<div class="content">
		<form name="form" id="form" method="post">
		<input type="hidden" name="SMode" value="MemberJoin" />
		<section class="join-agreement">
			<h1 class="h-level1">회원가입</h1>
			<p class="current-step step1">STEP 01<span></span></p>
			<ul class="agreement-list">
				<li>
					<div class="fieldset">
						<span class="checkbox">
							<input type="checkbox" id="agree-all" name="agree" data-allchk="agreement">
						</span>
						<label for="agree-all">전체동의</label>
					</div>
					<div class="agree-info">
						<ul>
							<li class="bullet-ty1">아래 모든 약관 (필수/선택 포함) 및 광고성 정보수신 내용을 확인하고 전체 동의합니다.</li>
							<li class="bullet-ty1">선택 항목에 대한 동의를 거부하더라도 회원가입에 영향을 미치지 않습니다.</li>
						</ul>
					</div>
				</li>
				<li>
					<div class="fieldset">
						<span class="checkbox">
							<input type="checkbox" id="agreement1" name="agreechk1" value="Y" data-allparts="agreement">
						</span>
						<label for="agreement1">[필수] 이용약관</label>
						<a href="javascript:PolicyView('22');" class="icon is-notext ico-go"></a>
					</div>
				</li>
				<li>
					<div class="fieldset">
						<span class="checkbox">
							<input type="checkbox" id="agreement2" name="agreechk2" value="Y" data-allparts="agreement">
						</span>
						<label for="agreement2">[필수] 개인정보 이용 및 수집에 대한 동의</label>
						<a href="javascript:PolicyView(23);" class="icon is-notext ico-go"></a>
					</div>
				</li>
				<li>
					<div class="fieldset">
						<span class="checkbox">
							<input type="checkbox" id="agreement3" name="agreechk3" value="Y" data-allparts="agreement">
						</span>
						<label for="agreement3">[선택] 마케팅 및 광고 활용 동의</label>
						<a href="javascript:PolicyView(25);" class="icon is-notext ico-go"></a>
					</div>
				</li>
				<li>
					<div class="fieldset">
						<span class="checkbox">
							<input type="checkbox" id="agreement4" name="agreechk4" value="Y" data-allparts="agreement">
						</span>
						<label for="agreement4">[선택] 개인정보 제3자 제공 동의</label>
						<a href="javascript:PolicyView(20);" class="icon is-notext ico-go"></a>
					</div>
				</li>
				<li>
					<div class="fieldset">
						<span class="checkbox">
							<input type="checkbox" id="agreement5" name="ThirdPartyFlag" value="Y" data-allparts="agreement">
						</span>
						<label for="agreement5">[선택] 개인정보 취급 위탁관련</label>
						<a href="javascript:PolicyView(21);" class="icon is-notext ico-go"></a>
					</div>
				</li>
				<li>
					<div class="fieldset">
						<span class="checkbox">
							<input type="checkbox" id="agreement6" name="SMSFlag" value="Y" data-allparts="agreement">
						</span>
						<label for="agreement6">[선택] 슈마커 할인, 이벤트 소식 문자 수신 동의</label>
					</div>
				</li>
				<li>
					<div class="fieldset">
						<span class="checkbox">
							<input type="checkbox" id="agreement7" name="EmailFlag" value="Y" data-allparts="agreement">
						</span>
						<label for="agreement7">[선택] 슈마커 할인, 이벤트 소식 이메일 수신 동의</label>
					</div>
				</li>
			</ul>
			<div class="agree-info">
				<ul>
					<li class="bullet-ty1">슈마커 쇼핑몰은 입점몰의 원활한 운영을 위하여 이용자의 개인정보 일부를...</li>
					<li class="bullet-ty1">슈마커에서는 “정보통신망 이용촉진 및 정보보호 등에 관한 법률” 제 23의 2 “주민등록번호의 사용 제안” 에 의거 고객님의 주민번호를 수집, 보관, 이용하지 않습니다.</li>
				</ul>
			</div>
			<%IF JoinType="D" THEN%>
			<div class="fieldset confirm-btn">
				<a href="javascript:agr_MemberTerms('form', 'Nice');" class="button is-expand ty-red">보호자 휴대폰 인증</a>
				<a href="javascript:agr_MemberTerms('form', 'Ipin');" class="button is-expand ty-black">보호자 아이핀 인증</a>
			</div>
			<div class="inf-type1">
				<p class="tit">14세 미만은 법률에 의거, 보호자(법적대리인)의 동의가 필요합니다.</p>
			</div>
			<%ELSE%>
			<div class="fieldset confirm-btn">
				<div class="confirm-adult">
					<a href="javascript:agr_MemberTerms('form', 'Nice');" class="button is-expand ty-red">휴대폰 인증</a>
					<a href="javascript:agr_MemberTerms('form', 'Ipin');" class="button is-expand ty-black">아이핀 인증</a>
				</div>
			</div>
			<%END IF%>
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