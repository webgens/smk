<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'Login.asp - 로그인 폼 페이지
'Date		: 2018.10.29
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
DIM SnsLink
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'
ProgID			 = Request("ProgID")
IF ProgID		 = "" THEN ProgID = "/"

	
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
	</script>
<!-- #include virtual="/INC/TopMain.asp" -->



	<main id="container" class="container">
	    <div class="content">

	        <nav class="login-method enter-login">
	            <ul>
	                <li class="part-2"><a href="javascript:chg_LoginForm('0')" class="current">회원</a></li>
	                <li class="part-2"><a href="javascript:chg_LoginForm('1')">비회원</a></li>
	            </ul>
	        </nav>

	        <section id="MLogin" class="login-form login-members">

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
                        <button type="button" class="naver" onclick="javascript:pop_NaverLogin();">네이버 계정으로 로그인</button>
                    </li>
                    <li>
                        <button type="button" class="facebook" onclick="javascript:pop_FacebookLogin();">페이스북 계정으로 로그인</button>
                    </li>
                    <li>
                        <button type="button" class="kakao" onclick="javascript:pop_KakaoLogin();">카카오 계정으로 로그인</button>
                    </li>
					<!--
                    <li>
                        <button type="button" class="google" onclick="javascript:pop_GoogleLogin();">구글 계정으로 로그인</button>
                    </li>
					//-->
                </ul>
	        </section>
			<%END IF %>






            <section id="NLogin" class="login-form" style="display:none;">
				
				<form name="formNLogin" id="formNLogin" onsubmit="return false" autocomplete="off">
				<input type="hidden" name="ProgID" value="/ASP/Mypage/OrderList.asp" />
                <fieldset>
                    <!-- psd_181212수정 -->
                    <legend class="tit">비회원 주문 조회</legend>
                    <div class="fieldset">
                        <label class="fieldset-label">이름</label>
                        <div class="fieldset-row">
                            <span class="input is-expand">
                                <input type="text" name="Name" id="Name" maxlength="20" placeholder="이름을 입력해주세요.">
                            </span>
                        </div>
                    </div>
                    <div class="fieldset ty-col2">
                        <label class="fieldset-label">휴대폰 번호</label>
                        <div class="fieldset-row">
                            <span class="select2">
                                <select name="HP1" id="HP1">
										<option value="010">010</option>
										<option value="011">011</option>
										<option value="016">016</option>
										<option value="017">017</option>
										<option value="018">018</option>
										<option value="019">019</option>
                                </select>
                                <span class="value"></span>
                            </span>
							<span class="dash1">-</span>
                            <span class="input2">
                                <input type="text" name="HP2" id="HP2" maxlength="4">
                            </span>
							<span class="dash2">-</span>
                            <span class="input3">
                                <input type="text" name="HP3" id="HP3" maxlength="4">
                            </span>
                        </div>
                    </div>
                    <div class="fieldset">
                        <label class="fieldset-label">이메일</label>
                        <div class="fieldset-row">
                            <span class="input is-expand">
                                <input type="email" name="Email" id="Email" maxlength="50" placeholder="이메일주소를 입력해주세요.">
                            </span>
                        </div>
                    </div>
                </fieldset>
                <button type="button" onclick="chk_NLogin();" class="button is-expand ty-red">비회원 주문조회</button>
                </form>

                <ul class="descript">
                    <li class="bullet-ty1">이름, 휴대폰번호, 이메일주소 모두 입력해 주세요.</li>
                    <li class="bullet-ty1">기억나지 않으실 경우, 고객센터에 문의해주세요.<br><strong>슈마커 고객센터 : <a href="tel:080-030-2809">080-030-2809</a></strong></li>
                </ul>
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



			$(".select2 .value").text($("select[name='HP1']").val());

			$(".select2 select").on("focus", function() {
				$(".select2").addClass("is-focus");
			});
			$(".select2 select").on("blur", function() {
				$(".select2").removeClass("is-focus");
			});
			$(".select2 select").on("change", function() {
				$(".select2 .value").text($("select[name='HP1']").val());
			});
		
			$(".input2 input").on("focus", function() {
				$(".input2").addClass("is-focus");
			});
			$(".input2 input").on("blur", function() {
				$(".input2").removeClass("is-focus");
			});
		
			$(".input3 input").on("focus", function() {
				$(".input3").addClass("is-focus");
			});
			$(".input3 input").on("blur", function() {
				$(".input3").removeClass("is-focus");
			});
		
			$(".input1 input").on("focus", function() {
				$(".input1").addClass("is-focus");
			});
			$(".input1 input").on("blur", function() {
				$(".input1").removeClass("is-focus");
			});
		});
	</script>

<script type="text/javascript" async src="//cdn-aitg.widerplanet.com/js/wp_astg_4.0.js"></script>


<!-- #include virtual="/INC/Footer.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>