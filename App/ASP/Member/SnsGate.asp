<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'snsGate.asp - sns 로그인설정
'Date		: 2018.12.14
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
PageCode3 = "01"
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->

<%
IF U_ID <> "" AND U_MFLAG = "Y" THEN
		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=&Msg=&Script=APP_PopupHistoryBack_Move('/');"
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

DIM OrderFlag
DIM SubFlag
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

	
OrderFlag		 = TRIM(Decrypt(Request.Cookies("LN_ORDER")))
SubFlag			 = TRIM(Decrypt(Request.Cookies("LN_SUB")))


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>


<!-- #include virtual="/INC/Header.asp" -->
	<script type="text/javascript" src="/JS/dev/join.js?ver=<%=U_DATE%><%=U_TIME%>"></script>
<!-- #include virtual="/INC/PopupTop.asp" -->


<!-- Main -->
<main id="container" class="container">
	<div class="content">
		<form name="form" method="post">
		<input type="hidden" name="SMode" value="MemberJoin" />
		<section class="join-agreement">
			<h1 class="h-level1">SNS 로그인 설정</h1>
			<ul class="agreement-list">
				<li style="text-align:center;border:0;">
					<div class="fieldset" style="margin-bottom:-15px;">
						<label for="agree-all">기존에 이용하시던 아이디/비밀번호를 입력하시면</label>
					</div>
					<div class="fieldset">
						<label for="agree-all">과거 주문내역/회원정보와 함께 자동 연결됩니다.</label>
					</div>
				</li>
			</ul>
			<div class="fieldset confirm-btn">
				<%IF OrderFlag = "Y" OR SubFlag = "Y" THEN%>
				<a href="javascript:void(0)" onclick="APP_PopupHistoryBack_Order_Sns_Connect()" class="button is-expand ty-black">기존회원 연결</a>
				<%ELSE%>
				<a href="javascript:void(0)" onclick="location.href='/ASP/Member/Login.asp?snsLink=Y'" class="button is-expand ty-black">기존회원 연결</a>
				<%END IF%>
			</div>


			<ul class="agreement-list">
				<li style="text-align:center;border:0;">
					<div class="fieldset" style="margin-bottom:-15px;">
						<label for="agree-all">방금 인증받은 SNS계정으로 이용할 수 있으며,</label>
					</div>
					<div class="fieldset">
						<label for="agree-all">회원전용 쿠폰/혜택이 제한될 수 있습니다.</label>
					</div>
				</li>
			</ul>
			<div class="fieldset confirm-btn">
				<a href="/ASP/Member/SnsAgreement.asp" class="button is-expand ty-red">간편로그인 계속하기</a>
			</div>

		</section>
		</form>
	</div>
</main>

<!-- #include virtual="/INC/FooterNoBNB.asp" -->
<!-- #include virtual="/INC/PopupBottom.asp" -->

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>