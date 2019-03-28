<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'SnsAgreement.asp - SNS 약관 동의 페이지
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

DIM SNS_UID
DIM SNS_Email
DIM SNS_KName
DIM SNS_Kind
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

SNS_UID		= Decrypt(Request.Cookies("SNS_UID"))
SNS_Email	= Decrypt(Request.Cookies("SNS_Email"))
SNS_KName	= Decrypt(Request.Cookies("SNS_KName"))
SNS_Kind	= Decrypt(Request.Cookies("SNS_Kind"))

IF SNS_UID = "" OR SNS_Kind = "" THEN
		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=SNS 정보가 없습니다.<br />다시 로그인하여 주십시오.&Script=APP_PopupHistoryBack('/');"
		Response.End
END IF


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>


<!-- #include virtual="/INC/Header.asp" -->
	<script type="text/javascript" src="/JS/dev/join.js?ver=<%=U_DATE%><%=U_TIME%>"></script>
<!-- #include virtual="/INC/PopupTop.asp" -->


<!-- Main -->
<main id="container" class="container">
	<div class="content">
		<form name="formSnsAgreement" id="formSnsAgreement">
		<section class="join-agreement">
			<h1 class="h-level1">약관동의</h1>
			<ul class="agreement-list">
				<li>
					<div class="fieldset">
						<span class="checkbox">
							<input type="checkbox" id="agreement2" name="agreechk2" value="Y" data-allparts="agreement">
						</span>
						<label for="agreement2">[필수] 개인정보 이용 및 수집에 대한 동의</label>
						<a href="javascript:PolicyView(23);" class="icon is-notext ico-go"></a>
					</div>
				</li>
			</ul>

			<div class="fieldset">
				<label for="join-email" class="fieldset-label">이메일주소</label>
				<div class="fieldset-row">
					<span class="input is-expand">
						<input type="email" id="SnsEmail" name="SnsEmail" maxlength="30" placeholder="이메일주소를 입력해주세요." value="<%=SNS_Email%>" <%IF SNS_Email<>"" THEN Response.write "readonly='readonly'"%> >
					</span>
				</div>
			</div>

			<div class="agree-info">
				<ul>
					<li class="bullet-ty1">슈마커 쇼핑몰은 입점몰의 원활한 운영을 위하여 이용자의 개인정보 일부를...</li>
					<li class="bullet-ty1">슈마커에서는 “정보통신망 이용촉진 및 정보보호 등에 관한 법률” 제 23의 2 “주민등록번호의 사용 제안” 에 의거 고객님의 주민번호를 수집, 보관, 이용하지 않습니다.</li>
				</ul>
			</div>

			<div class="fieldset confirm-btn">
				<a href="javascript:chk_SnsJoin();" class="button is-expand ty-red">간편로그인</a>
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