<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'JoinForm.asp - 회원 가입 완료 - 내용입력
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
PageCode3 = "04"
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
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>
<!-- #include virtual="/INC/Header.asp" -->

<%TopSubMenuTitle = "회원가입완료"%>
<!-- #include virtual="/INC/TopSub.asp" -->

    <!-- Main -->
    <main id="container" class="container">
        <div class="content">

            <section class="join-complete">
                <h1 class="h-level1">회원가입</h1>
                <p class="current-step step3">Step 03<span></span></p>

                <div class="join-inform">
                    <p class="t-level5">슈마커 회원이 되어주셔서 감사합니다.<br>신규 멤버십 회원에게만 지급해 드리는 쿠폰을<br>지금 바로 확인해보세요!</p>
                    <div class="join-coupon">
                        <img src="/Images/img/img_join_coupon.png" alt="쿠폰 신규회원 10% 할인, 가입 후 첫구매 5,000원 할인, 모바일 앱 쿠폰 3% 할인 중복사용가능">
                    </div>
                </div>

                <div class="buttongroup is-expand">
                    <a href="javascript:void(0)" onclick="APP_TopGoUrl('/ASP/Member/Login.asp')" class="button ty-red">로그인</a>
                    <a href="javascript:void(0)" onclick="APP_TopGoUrl('/')" class="button ty-white">메인으로</a>
                </div>
            </section>

        </div>
    </main>

<!-- This script is for AceCounter START -->
<script type="text/javascript">
var m_jn = 'join';          //  가입탈퇴 ( 'join','withdraw' ) 
var m_jid = '<%=Decrypt(Request.Cookies("JoinTempID"))%>' ;			// 가입시입력한 ID
</script>
<!-- AceCounter END -->

<!--
<script type="text/javascript" src="//wcs.naver.net/wcslog.js"></script>
<script type="text/javascript">
	var _nasa = {}; _nasa["cnv"] = wcs.cnv('2', '<%=Decrypt(Request.Cookies("JoinTempID"))%>');
</script>
//-->


<!-- WIDERPLANET  SCRIPT START 2019.1.24 -->
<div id="wp_tg_cts" style="display:none;"></div>
<script type="text/javascript">
	var wptg_tagscript_vars = wptg_tagscript_vars || [];
	wptg_tagscript_vars.push(
	(function () {
		return {
			wp_hcuid: "",  /*고객넘버 등 Unique ID (ex. 로그인  ID, 고객넘버 등 )를 암호화하여 대입.
                     *주의 : 로그인 하지 않은 사용자는 어떠한 값도 대입하지 않습니다.*/
			ti: "24585",
			ty: "Join",                        /*트래킹태그 타입 */
			device: "mobile",                  /*디바이스 종류  (web 또는  mobile)*/
			items: [{
				i: "회원 가입",          /*전환 식별 코드  (한글 , 영어 , 번호 , 공백 허용 )*/
				t: "회원 가입",          /*전환명  (한글 , 영어 , 번호 , 공백 허용 )*/
				p: "1",                   /*전환가격  (전환 가격이 없을 경우 1로 설정 )*/
				q: "1"                   /*전환수량  (전환 수량이 고정적으로 1개 이하일 경우 1로 설정 )*/
			}]
		};
	}));
</script>
<script type="text/javascript" async src="//cdn-aitg.widerplanet.com/js/wp_astg_4.0.js"></script>
<!-- // WIDERPLANET  SCRIPT END 2019.1.24 -->

<!-- Facebook Pixel Code -->
<script>
  fbq('track', 'CompleteRegistration');
</script>
<!-- End Facebook Pixel Code -->

<!-- kakao pixel script //-->
<script type="text/javascript" charset="UTF-8" src="//t1.daumcdn.net/adfit/static/kp.js"></script>
<script type="text/javascript">
	kakaoPixel('5354511058043421336').pageView();
	kakaoPixel('5354511058043421336').completeRegistration();
</script>
<!-- kakao pixel script //-->

<!-- #include virtual="/INC/FooterNoBNB.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->

<%
Response.Cookies("JoinTempID") = ""

SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>