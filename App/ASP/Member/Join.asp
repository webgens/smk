<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'Join.asp - 회원 가입 - 14세 구분
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
PageCode2 = "01"
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
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>


<!-- #include virtual="/INC/Header.asp" -->
	<script type="text/javascript" src="/JS/dev/join.js?ver=<%=U_DATE%><%=U_TIME%>"></script>

<%TopSubMenuTitle = "회원가입"%>
<!-- #include virtual="/INC/TopSub.asp" -->

    <!-- Main -->
    <main id="container" class="container">
        <div class="content">

            <section class="join-intro intro-section">
                <h1 class="h-level1">회원가입</h1>
                <div class="join-inform">
                    <p class="t-level5">슈마커 회원으로 가입하고 <br> 다양한 멤버십 혜택을 받아가세요!</p>
                </div>
                <div class="fieldset">
                    <a href="javascript:move_Certification('','U')" class="button is-expand ty-red"><strong>14세 이상</strong> 회원가입</a>
                    <a href="javascript:move_Certification('','D')" class="button is-expand ty-black"><strong>14세 미만</strong> 회원가입</a>
                </div>
            </section>
            <section class="join-introduce">
                <div class="inf-type1">
                    <p class="tit">안전한 회원가입을 위해 본인확인 인증을 진행하고 있습니다.</p>
                    <ul>
                        <li class="bullet-ty1">슈마커 쇼핑몰은 회원님의 개인정보를 신중히 취급하며 안전하게 관리하고 있습니다.</li>
                        <li class="bullet-ty1">입력하신 개인정보는 신용평가기관에 본인확인 목적으로만 사용 됩니다.
                        </li>
                        <li class="bullet-ty1">타인의 정보 및 주민등록번호를 부정하게 사용하는 경우 3년 이하 징역 또는 1천만원 이하의 벌금에 처해지게 됩니다.</li>
                    </ul>
                </div>
            </section>
        </div>
    </main>

	<form name="formMoveCert" id="formMoveCert" method="post" action="JoinCertification.asp">
		<input type="hidden" name="JoinType" />
	</form>

<!-- #include virtual="/INC/FooterNoBNB.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>