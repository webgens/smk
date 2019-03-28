<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'NewAgreement.asp - 리뉴얼 신규 약관 동의 페이지
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
PageCode2 = "04"
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
		Response.Redirect("/")
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

DIM MemberNum
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


MemberNum		 = TRIM(Decrypt(Request.Cookies("TEMP_UNUM")))

IF MemberNum	 = "" THEN
		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=로그인 정보가 없습니다.<br />다시 로그인하여 주십시오.&Script=APP_TopGoUrl('/ASP/Member/Login.asp');"
		Response.End
END IF


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>


<!-- #include virtual="/INC/Header.asp" -->
<script type="text/javascript" src="/JS/dev/login.js?ver=<%=U_DATE%><%=U_TIME%>"></script>

<%TopSubMenuTitle = "휴면계정해제"%>
<!-- #include virtual="/INC/TopSub.asp" -->


<!-- Main -->
<main id="container" class="container">
	<div class="content">
		<form name="formNewAgreement" id="formNewAgreement">
		<section class="join-agreement">
			<h1 class="h-level1">신규약관동의</h1>
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
							<li class="bullet-ty1">선택 항목에 대한 동의를 거부하더라도 영향을 미치지 않습니다.</li>
						</ul>
					</div>
				</li>
				<li>
					<div class="fieldset">
						<span class="checkbox">
							<input type="checkbox" id="agreement1" name="Agr1" value="Y" data-allparts="agreement">
						</span>
						<label for="agreement1">[필수] 이용약관</label>
						<a href="javascript:PolicyView('22');" class="icon is-notext ico-go"></a>
					</div>
				</li>
				<li>
					<div class="fieldset">
						<span class="checkbox">
							<input type="checkbox" id="agreement2" name="Agr2" value="Y" data-allparts="agreement">
						</span>
						<label for="agreement2">[필수] 개인정보 이용 및 수집에 대한 동의</label>
						<a href="javascript:PolicyView(23);" class="icon is-notext ico-go"></a>
					</div>
				</li>
				<li>
					<div class="fieldset">
						<span class="checkbox">
							<input type="checkbox" id="agreement3" name="Agr3" value="Y" data-allparts="agreement">
						</span>
						<label for="agreement3">[선택] 마케팅 및 광고 활용 동의</label>
						<a href="javascript:PolicyView(25);" class="icon is-notext ico-go"></a>
					</div>
				</li>
				<li>
					<div class="fieldset">
						<span class="checkbox">
							<input type="checkbox" id="agreement4" name="Agr4" value="Y" data-allparts="agreement">
						</span>
						<label for="agreement4">[선택] 개인정보 제3자 제공 동의</label>
						<a href="javascript:PolicyView(20);" class="icon is-notext ico-go"></a>
					</div>
				</li>
				<li>
					<div class="fieldset">
						<span class="checkbox">
							<input type="checkbox" id="agreement5" name="Agr5" value="Y" data-allparts="agreement">
						</span>
						<label for="agreement5">[선택] 개인정보 취급 위탁관련</label>
						<a href="javascript:PolicyView(21);" class="icon is-notext ico-go"></a>
					</div>
				</li>
				<li>
					<div class="fieldset">
						<span class="checkbox">
							<input type="checkbox" id="agreement6" name="Agr6" value="Y" data-allparts="agreement">
						</span>
						<label for="agreement6">[선택] 슈마커 할인, 이벤트 소식 문자 수신 동의</label>
					</div>
				</li>
				<li>
					<div class="fieldset">
						<span class="checkbox">
							<input type="checkbox" id="agreement7" name="Agr7" value="Y" data-allparts="agreement">
						</span>
						<label for="agreement7">[선택] 슈마커 할인, 이벤트 소식 이메일 수신 동의</label>
					</div>
				</li>
			</ul>
			<div class="fieldset confirm-btn">
				<a href="javascript:chk_NewAgreement();" class="button is-expand ty-red">약관동의</a>
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