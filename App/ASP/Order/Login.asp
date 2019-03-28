<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'Login.asp - 로그인 폼 페이지
'Date		: 2018.12.28
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

DIM SavedCookieID				'# 저장 아이디
DIM SNS_Kind
DIM snsLink
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

	
Response.Cookies("LN_ORDER")		 = Encrypt("Y")
IF TRIM(Decrypt(Request.Cookies("LN_PROGID"))) = "" THEN
		ProgID			 = Request("ProgID")
		IF ProgID		 = "" THEN ProgID = "/"
		Response.Cookies("LN_PROGID")		 = Encrypt(ProgID)
ELSE
		ProgID			 = TRIM(Decrypt(Request.Cookies("LN_PROGID")))
END IF



SavedCookieID	 = TRIM(Decrypt(Request.Cookies("SMEM_ID")))
snsLink			 = sqlFilter(Request("snsLink"))
IF snsLink = "" THEN
		'# SNS 회원정보 초기화
		Response.Cookies("SNS_UID")		 = ""
		Response.Cookies("SNS_Kind")	 = ""
		Response.Cookies("SNS_Email")	 = ""
		Response.Cookies("SNS_KName")	 = ""
		Response.Cookies("SNS_UserID")	 = ""
		Response.Cookies("SNS_UNUM")	 = ""
ELSE
		SNS_Kind	 = Decrypt(Request.Cookies("SNS_Kind"))
END IF

SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>


<!-- #include virtual="/INC/Header.asp" -->
	<style type="text/css">
		.personal-collect .cnt { height: 88%; overflow: auto; font-size: 9px; color: #767676; line-height: 1.6; }
	</style>
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
			location.href = "/ASP/Order/Login.asp?SnsLink=Y";
		}

		function chg_LoginForm(num) {
			$(".part-2 > a").removeClass("current");
			$(".part-2 > a").eq(num).addClass("current");
			if (num == "0") {
				$("#MLogin").show();
				$("#SLogin").show();
				$("#NLogin").hide();
			}
			else {
				$("#MLogin").hide();
				$("#SLogin").hide();
				$("#NLogin").show();
			}
		}


		/* 비회원 주문하기 */
		function nonMemberOrder() {
			if ($("#Agreement").is(":checked") == false) {
				openAlertLayer("alert", "개인정보 수집/이용에 동의 하셔야<br />비회원 주문이 가능합니다.", "closePop('alertPop', 'Agreement');", "");
				return;
			}

			$.ajax({
				type		 : "post",
				url			 : "/ASP/Order/Ajax/NonMemberOrderOk.asp",
				async		 : false,
				data		 : $("#formNonMember").serialize(),
				dataType	 : "text",
				success		 : function (data) {
								var splitData	 = data.split("|||||");
								var result		 = splitData[0];
								var cont		 = splitData[1];


								if (result == "OK") {
									var progID = $("#ProgID").val();
									location.href = progID;
									return;
								}
								else {
									openAlertLayer("alert", cont, "closePop('alertPop', '');", "");
									return;
								}
				},
				error		 : function (data) {
								openAlertLayer("alert", "처리 도중 오류가 발생하였습니다.", "closePop('alertPop', '');", "");
				}
			});
		}
	</script>

<%TopSubMenuTitle = "LOGIN"%>
<!-- #include virtual="/INC/TopSub.asp" -->


    <!-- Main -->
	<main id="container" class="container" style="margin-bottom:50px;">
	    <div class="sub_content" style="padding-bottom:0">

            <nav class="login-method enter-login">
                <ul>
	                <li class="part-2"><a href="javascript:chg_LoginForm('0')" class="current">회원</a></li>
	                <li class="part-2"><a href="javascript:chg_LoginForm('1')">비회원</a></li>
                </ul>
            </nav>

            <section id="MLogin" class="login-form login-members">
                <!-- 필요 시 form 엘리먼트로 교체 사용 -->
				<form name="formLogin" id="formLogin" onsubmit="return false" autocomplete="off">
				<input type="hidden" name="ProgID" id="ProgID" value="<%=ProgID%>" />
				<input type="hidden" name="OrderFlag" value="Y" />
                <fieldset>
                    <legend class="hidden">로그인 정보 입력</legend>
                    <div class="fieldset">
                        <span class="input is-expand">
							<input type="text" name="UserID" title="아이디" value="<%=SavedCookieID%>" maxlength="30" placeholder="아이디" onfocus="init_Login()">
						</span>
                    </div>
                    <div class="fieldset">
                        <span class="input is-expand">
							<input type="password" name="Pwd" title="비밀번호" maxlength="20" placeholder="비밀번호" onfocus="init_Login()">
						</span>
                    </div>
                    <div class="fieldset login">
                        <span class="checkbox">
							<input type="checkbox" id="pick_up1" name="saveid" value="Y" <%IF SavedCookieID <> "" THEN%>checked="checked"<%END IF%>>
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
                    <li>
                        <button type="button" class="google" onclick="pop_GoogleLogin();">구글 계정으로 로그인</button>
                    </li>
                </ul>
            </section>
			<%END IF %>






            <section id="NLogin" class="login-form" style="display:none;">
                <div class="nonmember-agree">
                    <p class="tit">비회원 구매</p>
                    <div class="personal-collect">
                        <!--<p class="tit2">개인정보 수집 범위</p>-->
                        <ul class="cnt">
<%
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Admin_EShop_Policy_Select_By_IDX"

		.Parameters.Append .CreateParameter("@IDX",		adInteger,	adParamInput,	  ,		24)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing

IF NOT oRs.EOF THEN
	'Policy = oRs("Contents")
	Response.Write oRs("Contents")
END IF
oRs.Close
%>
                            <!--<li>① 슈마커는 고객님의 보다 편리한 쇼핑믈 위해 오프라인 매장에서 직접 상품을 받으실 수 있는 매장픽업 서비스를 운영중에 있습니다.</li>
                            <li>② 슈머커는 개인정보를 필수사항과 선택사항으로 구분하여 수집하고 있습니다.<br>- 필수항목: 성명(실명), 전화연락처(휴대폰 번호), 기타 본 서비스 이용과정에서 발생한 상품 또는 서비스 구매 내역, 접속 기록, 쿠키, 가입인증정보 </li>
                            <li>① 슈마커는 고객님의 보다 편리한 쇼핑믈 위해 오프라인 매장에서 직접 상품을 받으실 수 있는 매장픽업 서비스를 운영중에 있습니다.</li>
                            <li>② 슈머커는 개인정보를 필수사항과 선택사항으로 구분하여 수집하고 있습니다.</li>
                            <li>① 슈마커는 고객님의 보다 편리한 쇼핑믈 위해 오프라인 매장에서 직접 상품을 받으실 수 있는 매장픽업 서비스를 운영중에 있습니다.</li>
                            <li>② 슈머커는 개인정보를 필수사항과 선택사항으로 구분하여 수집하고 있습니다.<br>- 필수항목: 성명(실명), 전화연락처(휴대폰 번호), 기타 본 서비스 이용과정에서 발생한 상품 또는 서비스 구매 내역, 접속 기록, 쿠키, 가입인증정보 </li>
                            <li>① 슈마커는 고객님의 보다 편리한 쇼핑믈 위해 오프라인 매장에서 직접 상품을 받으실 수 있는 매장픽업 서비스를 운영중에 있습니다.</li>
                            <li>② 슈머커는 개인정보를 필수사항과 선택사항으로 구분하여 수집하고 있습니다.</li>
                            <li>① 슈마커는 고객님의 보다 편리한 쇼핑믈 위해 오프라인 매장에서 직접 상품을 받으실 수 있는 매장픽업 서비스를 운영중에 있습니다.</li>
                            <li>② 슈머커는 개인정보를 필수사항과 선택사항으로 구분하여 수집하고 있습니다.<br>- 필수항목: 성명(실명), 전화연락처(휴대폰 번호), 기타 본 서비스 이용과정에서 발생한 상품 또는 서비스 구매 내역, 접속 기록, 쿠키, 가입인증정보 </li>
                            <li>① 슈마커는 고객님의 보다 편리한 쇼핑믈 위해 오프라인 매장에서 직접 상품을 받으실 수 있는 매장픽업 서비스를 운영중에 있습니다.</li>
                            <li>② 슈머커는 개인정보를 필수사항과 선택사항으로 구분하여 수집하고 있습니다.</li>-->
                        </ul>
                    </div>

					<form name="formNonMember" id="formNonMember" method="post">
					<input type="hidden" name="ProgID" value="<%=ProgID%>" />
                    <div class="check">
                        <span class="checkbox">
                            <input type="checkbox" name="Agreement" id="Agreement" value="Y" />
                        </span>
                        <label for="Agreement">개인정보 수집/이용에 동의합니다.</label>
                    </div>
					</form>

                    <button type="button" onclick="nonMemberOrder()" class="button is-expand ty-red">비회원 구매하기</button>
                </div>
            </section>



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