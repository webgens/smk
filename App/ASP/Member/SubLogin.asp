<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'SubLogin.asp - 서브 페이지 로그인 폼 페이지
'Date		: 2019.01.19
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

DIM SavedCookieID				'# 저장 아이디
DIM SNS_Kind
DIM SnsLink
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

	
Response.Cookies("LN_SUB")		 = Encrypt("Y")
IF TRIM(Decrypt(Request.Cookies("LN_PROGID"))) = "" THEN
		ProgID			 = Request("ProgID")
		IF ProgID		 = "" THEN ProgID = "/"
		Response.Cookies("LN_PROGID")		 = Encrypt(ProgID)
ELSE
		ProgID			 = TRIM(Decrypt(Request.Cookies("LN_PROGID")))
END IF



SavedCookieID	 = TRIM(Decrypt(Request.Cookies("SMEM_ID")))
SnsLink			 = sqlFilter(Request("SnsLink"))
IF SnsLink = "" THEN
		'# SNS 회원정보 초기화
		Response.Cookies("SNS_UID")		 = ""
		Response.Cookies("SNS_Kind")		 = ""
		Response.Cookies("SNS_Email")	 = ""
		Response.Cookies("SNS_KName")	 = ""
		Response.Cookies("SNS_UserID")	 = ""
		Response.Cookies("SNS_UNUM")		 = ""
ELSE
		SNS_Kind	 = Decrypt(Request.Cookies("SNS_Kind"))
END IF


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>


<!-- #include virtual="/INC/Header.asp" -->
    <script type="text/javascript" src="/JS/dev/login.js?ver=<%=U_DATE%><%=U_TIME%>"></script>
	<script type="text/javascript">
		function move_AfterLogin() {
			location.href = $("#ProgID").val();
		}

		function move_MyCouponList() {
			location.href = "/ASP/Mypage/CouponList.asp";
		}


		function move_AfterSnsLogin(val) {
			if (val != "") {
				var splitVal	 = val.split("///");
				var uNum		 = splitVal[0];
				var id			 = splitVal[1];
				var email		 = splitVal[2];
				var name		 = splitVal[3];
				var kind		 = splitVal[4];

				$("input[name='UID']",		 "form[name='SimpleLoginForm']").val(id);
				$("input[name='Email']",	 "form[name='SimpleLoginForm']").val(email);
				$("input[name='KName']",	 "form[name='SimpleLoginForm']").val(name);
				$("input[name='SNSKind']",	 "form[name='SimpleLoginForm']").val(kind);

				if (uNum == "") {
					snsLogin();
				}
				else {
					snsConnection();
				}
			}
		}

		function move_AfterSnsConnect() {
			location.href = "/ASP/Member/SubLogin.asp?SnsLink=Y";
		}
	</script>

<%TopSubMenuTitle = "LOGIN"%>
<!-- #include virtual="/INC/TopSub.asp" -->



	<main id="container" class="container" style="margin-bottom:50px;">
	    <div class="sub_content" style="padding-bottom:0">

			<div style="height:30px"></div>

	        <section class="login-form login-members">

				<form name="formLogin" id="formLogin" onsubmit="return false" autocomplete="off">
				<input type="hidden" name="ProgID" id="ProgID" value="<%=ProgID%>" />
	            <fieldset>
	                <!-- psd_181212수정 -->
	                <legend class="tit">회원 로그인</legend>
	                <!-- //psd_181212수정 -->
	                <div class="fieldset">
	                    <span class="input is-expand">
							<input type="text" name="UserID" id="UserID" title="아이디" value="<%=SavedCookieID%>" maxlength="30" placeholder="아이디" onfocus="init_Login()">
						</span>
	                </div>
	                <div class="fieldset">
	                    <span class="input is-expand">
							<input type="password" name="Pwd" id="Pwd" title="비밀번호" maxlength="20" placeholder="비밀번호">
						</span>
	                </div>
	                <div class="fieldset login">
	                    <span class="checkbox">
							<input type="checkbox" id="keep-login" name="saveid" value="Y" <%IF SavedCookieID <> "" THEN%>checked="checked"<%END IF%>>
						</span>
	                    <label for="keep-login" class="lab-keep-login">로그인 상태 유지</label>
	                </div>
                    <button type="submit" class="button is-expand ty-red" onclick="chk_Login()">로그인</button>
	                <ul class="util">
                        <li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Member/Join.asp')">회원가입</a></li>
                        <li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Member/FindID.asp')">아이디/비밀번호 찾기</a></li>
	                </ul>
					
					<%IF SNS_Kind <> "" AND Request.Cookies("SNS_UID") <> "" THEN %>
					<div style="text-align:center;color:#c00;padding-bottom:20px;">로그인하시면 <% 
					IF SNS_Kind="N" THEN 
						Response.Write "네이버"
					ELSEIF SNS_Kind="K" THEN 
						Response.Write "카카오"
					ELSEIF SNS_Kind="F" THEN 
						Response.Write "페이스북"
					ELSEIF SNS_Kind="G" THEN 
						Response.Write "구글"
					END IF
					%>계정으로 연동됩니다.</div>
					<%END IF %>


	            </fieldset>
				</form>
	        </section>
			
			<%IF SNS_Kind = "" OR Request.Cookies("SNS_UID") = "" THEN %>
            <section id="SLogin" class="login-different">
                <ul>
                    <li>
                        <button type="button" class="naver" onclick="pop_NaverLogin();">네이버 계정으로 로그인</button>
                    </li>
                    <li>
                        <button type="button" class="facebook" onclick="pop_FacebookLogin();">페이스북 계정으로 로그인</button>
                    </li>
                    <li>
                        <button type="button" class="kakao" onclick="pop_KakaoLogin();">카카오 계정으로 로그인</button>
                    </li>
                    <!--
					<li>
                        <button type="button" class="google" onclick="pop_GoogleLogin();">구글 계정으로 로그인</button>
                    </li>
					//-->
                </ul>
            </section>
			<%END IF %>

	    </div>
	</main>


	<form name="SimpleLoginForm" id="SimpleLoginForm" method="post">
		<input type="hidden" name="UID">
		<input type="hidden" name="Email">
		<input type="hidden" name="KName">
		<input type="hidden" name="SNSKind">
	</form>


	<script type="text/javascript">
		$(function () {
			if ($("input:checkbox[name='saveid']").is(":checked")) {
				$("input[name='Pwd']").focus();
			}
			else {
				$("input[name='UserID']").focus();
			}
		});
	</script>


<script type="text/javascript" async src="//cdn-aitg.widerplanet.com/js/wp_astg_4.0.js"></script>

<!-- #include virtual="/INC/FooterNone.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>