<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'JoinChgMemForm.asp - 회원 가입(정회원 전환) - 내용입력
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
PageCode2 = "06"
PageCode3 = "03"
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->

<%
IF U_ID <> "" AND U_MFLAG = "Y" THEN
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

DIM AgreeChk1
DIM AgreeChk2
DIM AgreeChk3
DIM AgreeChk4
DIM ThirdPartyFlag
DIM SMSFlag
DIM EMailFlag
	
DIM JoinType					'# 14세 구분 (U:14세이상 / D:14세미만)
DIM AuthType

DIM SDupInfo
DIM Name
DIM Birth
DIM Gender
DIM Mobile
DIM HP1
DIM HP2
DIM HP3

DIM ParentSDupInfo
DIM ParentName
DIM ParentBirth
DIM ParentGender
DIM ParentMobile
DIM PHP1
DIM PHP2
DIM PHP3

DIM Email
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


AgreeChk1			 = sqlFilter(Request("AgreeChk1"))
AgreeChk2			 = sqlFilter(Request("AgreeChk2"))
AgreeChk3			 = sqlFilter(Request("AgreeChk3"))
AgreeChk4			 = sqlFilter(Request("AgreeChk4"))
ThirdPartyFlag		 = sqlFilter(Request("ThirdPartyFlag"))
SMSFlag				 = sqlFilter(Request("SMSFlag"))
EMailFlag			 = sqlFilter(Request("EMailFlag"))
	
JoinType			 = TRIM(Decrypt(Request.Cookies("JoinType")))
AuthType			 = TRIM(Decrypt(Request.Cookies("AuthType")))

SDupInfo			 = TRIM(Decrypt(Request.Cookies("SDupInfo")))
Name				 = TRIM(Decrypt(Request.Cookies("Name")))
Birth				 = TRIM(Decrypt(Request.Cookies("Birth")))
Gender				 = TRIM(Decrypt(Request.Cookies("Gender")))
Mobile				 = TRIM(Decrypt(Request.Cookies("Mobile")))
IF Mobile <> "" THEN
		IF LEN(Mobile) = 11 THEN
				HP1 = LEFT(Mobile, 3)
				HP2 = MID(Mobile, 4, 4)
				HP3 = MID(Mobile, 8, 4)
		ELSEIF LEN(Mobile) = 10 THEN
				HP1 = LEFT(Mobile, 3)
				HP2 = MID(Mobile, 4, 3)
				HP3 = MID(Mobile, 7, 4)
		ELSE
				Mobile = ""
		END IF		
END IF

ParentSDupInfo		 = TRIM(Decrypt(Request.Cookies("ParentSDupInfo")))
ParentName			 = TRIM(Decrypt(Request.Cookies("ParentName")))
ParentBirth			 = TRIM(Decrypt(Request.Cookies("ParentBirth")))
ParentGender		 = TRIM(Decrypt(Request.Cookies("ParentGender")))
ParentMobile		 = TRIM(Decrypt(Request.Cookies("ParentMobile")))
IF ParentMobile <> "" THEN
		IF LEN(ParentMobile) = 11 THEN
				PHP1 = LEFT(ParentMobile, 3)
				PHP2 = MID(ParentMobile, 4, 4)
				PHP3 = MID(ParentMobile, 8, 4)
		ELSEIF LEN(Mobile) = 10 THEN
				PHP1 = LEFT(ParentMobile, 3)
				PHP2 = MID(ParentMobile, 4, 3)
				PHP3 = MID(ParentMobile, 7, 4)
		ELSE
				ParentMobile = ""
		END IF		
END IF

	
IF JoinType = "" THEN
		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=14세 구분값이 없습니다.&Script=location.href='/ASP/Member/JoinChgMem.asp';"
		Response.End
END IF

IF AgreeChk1 = "" OR AgreeChk2 = "" THEN
		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=이용약관 동의를 하셔야 합니다.&Script=location.href='/ASP/Member/JoinChgMem.asp';"
		Response.End
END IF

IF JoinType = "U" AND SDupInfo = "" THEN
		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=본인인증 정보가 없습니다.&Script=location.href='/ASP/Member/JoinChgMem.asp';"
		Response.End
END IF

IF JoinType = "D" AND ParentSDupInfo = "" THEN
		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=본인인증 정보가 없습니다.&Script=location.href='/ASP/Member/JoinChgMem.asp';"
		Response.End
END IF



SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성



SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Member_Select_By_MemberNum"

		.Parameters.Append .CreateParameter("@MemberNum", adInteger, adParamInput, , U_NUM)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
IF NOT oRs.EOF THEN
		Email = oRs("Email")
ELSE
		Response.Cookies("UIP").Expires			 = Now - 1000
		Response.Cookies("UMFLAG").Expires		 = Now - 1000
		Response.Cookies("UNUM").Expires		 = Now - 1000
		Response.Cookies("UID").Expires			 = Now - 1000
		Response.Cookies("UNAME").Expires		 = Now - 1000
		Response.Cookies("UMFLAG").Expires		 = Now - 1000
		Response.Cookies("UETYPE").Expires		 = Now - 1000
		Response.Cookies("UGROUP").Expires		 = Now - 1000

		
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Call AlertMessage("로그인 정보가 없습니다. 다시 로그인 하여 주십시오.", "location.href='/ASP/Member/Login.asp';")
		Response.End
END IF
oRs.Close
%>
<!-- #include virtual="/INC/Header.asp" -->
    <script type="text/javascript" src="/JS/dev/join.js?ver=<%=U_DATE%><%=U_TIME%>"></script>
    <script type="text/javascript" src="/JS/dev/login.js?ver=<%=U_DATE%><%=U_TIME%>"></script>

<%TopSubMenuTitle = "회원가입"%>
<!-- #include virtual="/INC/TopSub.asp" -->

    <!-- Main -->
    <main id="container" class="container">
        <div class="content">

			<form name="Joinform" id="Joinform" method="post" autocomplete="off">
			<input type="hidden" name="JoinType"		 value="<%=JoinType%>"		 />
			<input type="hidden" name="AgreeChk1"		 value="<%=AgreeChk1%>"		 />
			<input type="hidden" name="AgreeChk2"		 value="<%=AgreeChk2%>"		 />
			<input type="hidden" name="AgreeChk3"		 value="<%=AgreeChk3%>"		 />
			<input type="hidden" name="AgreeChk4"		 value="<%=AgreeChk4%>"		 />
			<input type="hidden" name="ThirdPartyFlag"	 value="<%=ThirdPartyFlag%>" />
			<input type="hidden" name="SMSFlag"			 value="<%=SMSFlag%>"		 />
			<input type="hidden" name="EMailFlag"		 value="<%=EMailFlag%>"		 />
			<input type="hidden" name="UserIDCheckFlag"								 />
			<input type="hidden" name="SMode"			 value="JoinChgMem"			 />
			<input type="hidden" name="Today"			 value="<%=U_DATE%>"		 />
            <section class="join-form">
                <h1 class="h-level1">회원가입(정회원 전환)</h1>
                <p class="current-step step2">STEP 02<span></span></p>

                <div class="formfield">
                    <p class="tit">가입자 정보</p>
                    <fieldset>
                        <legend class="hidden">기본 정보 입력</legend>
                        <div class="fieldset">
                            <label for="join-id" class="fieldset-label">아이디</label>
                            <div class="fieldset-row">
                                <span class="input-id is-expand">
                                    <input type="text" name="UserID" id="UserID" maxlength="12" placeholder="아이디 (영문 숫자포함 6~12)">
                                </span>
                                <span class="btn-wrap">
                                    <button type="button" id="IDChkBtn" onclick="chk_UserID();" class="button2 is-expand ty-red">중복체크</button>
                                </span>
                            </div>
							<p id="ID_Msg" class="message icon ico-caution" style="display:none"></p>
							
							<input type="hidden" name="CheckID" id="CheckID" />
							<input type="hidden" name="CheckIDAvailable" id="CheckIDAvailable" />
                        </div>
                        <div class="fieldset">
                            <label for="join-pw" class="fieldset-label">비밀번호</label>
                            <div class="fieldset-row">
                                <span class="input is-expand">
                                    <input type="password" name="Pwd" id="Pwd" maxlength="12" placeholder="비밀번호 (영문 숫자포함 6~12)">
                                </span>
                            </div>
                        </div>
                        <div class="fieldset">
                            <label for="join-pw-confirm" class="fieldset-label">비밀번호 확인</label>
                            <div class="fieldset-row">
                                <span class="input is-expand">
                                    <input type="password" name="Pwd1" id="Pwd1" maxlength="12" placeholder="비밀번호를 한번 더 입력해 주세요.">
                                </span>
                            </div>
                        </div>
                    </fieldset>
                    <fieldset class="">
                        <legend class="hidden">부가 정보 입력</legend>
                        <div class="fieldset">
                            <label for="join-name" class="fieldset-label">이름</label>
                            <div class="fieldset-row">
                                <span class="input is-expand">
									<%IF Name = "" THEN%>
                                    <input type="text" name="Name" id="Name" maxlength="25" placeholder="이름을 입력해 주세요.">
									<%ELSE%>
                                    <input type="text" name="Name" id="Name" maxlength="25" value="<%=Name%>" readonly="readonly" placeholder="이름을 입력해 주세요.">
									<%END IF%>
                                </span>
                            </div>
                        </div>
                        <div class="fieldset">
                            <label for="join-birth" class="fieldset-label">생년월일</label>
                            <div class="fieldset-row">
                                <span class="input is-expand">
									<%IF Birth = "" THEN%>
                                    <input type="tel" name="Birth" id="Birth" maxlength="8" placeholder="생년월일을 입력해 주세요.(ex. 19991231)">
									<%ELSE%>
                                    <input type="tel" name="Birth" id="Birth" maxlength="8" value="<%=Birth%>" readonly="readonly" class="input-no-border" placeholder="생년월일을 입력해 주세요.(ex. 19991231)">
									<%END IF%>
                                </span>
                            </div>
                        </div>
                        <div class="fieldset ty-row">
                            <label class="fieldset-label">성별</label>
                            <div class="fieldset-row">
                                <div class="radiogroup">
<%
IF Gender = "1" OR Gender = "2" THEN
		IF Gender = "1" THEN
%>
                                    <div class="inner">
                                        <span class="radio">
                                            <input type="radio" name="Sex" id="Sex1" value="M" checked="checked">
                                        </span>
                                        <label for="join-male">남</label>
                                    </div>
<%
		ELSE
%>
                                    <div class="inner">
                                        <span class="radio">
                                            <input type="radio" name="Sex" id="Sex2" value="F" checked="checked">
                                        </span>
                                        <label for="join-female">여</label>
                                    </div>
<%
		END IF
ELSE
%>
                                    <div class="inner">
                                        <span class="radio">
                                            <input type="radio" name="Sex" id="Sex1" value="M">
                                        </span>
                                        <label for="join-male">남</label>
                                    </div>
                                    <div class="inner">
                                        <span class="radio">
                                            <input type="radio" name="Sex" id="Sex2" value="F">
                                        </span>
                                        <label for="join-female">여</label>
                                    </div>
<%
END IF
%>
                                </div>
                            </div>
                        </div>
                    </fieldset>
                    <fieldset class="">
                        <legend class="hidden">인증 정보 입력</legend>
                        <div class="fieldset ty-col2">
                            <label for="join-phone" class="fieldset-label">휴대폰 번호</label>
                            <div class="fieldset-row">
								<%IF Mobile = "" THEN%>
                                <span class="select2">
                                    <select name="HP1" id="HP1" title="휴대폰 국번 선택">
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
                                    <input type="tel" name="HP2" id="HP2" maxlength="4">
                                </span>
								<span class="dash2">-</span>
                                <span class="input3">
                                    <input type="tel" name="HP3" id="HP3" maxlength="4">
                                </span>
								<input type="hidden" name="MobileFlag" value="N" />
								<%ELSE%>
                                <span class="input1">
                                    <input type="tel" name="HP1" id="HP1" maxlength="4" value="<%=HP1%>" readonly="readonly">
                                </span>
								<span class="dash1">-</span>
                                <span class="input2">
                                    <input type="tel" name="HP2" id="HP2" maxlength="4" value="<%=HP2%>" readonly="readonly">
                                </span>
								<span class="dash2">-</span>
                                <span class="input3">
                                    <input type="tel" name="HP3" id="HP3" maxlength="4" value="<%=HP3%>" readonly="readonly">
                                </span>
								<input type="hidden" name="MobileFlag" value="Y" />
								<%END IF%>
                            </div>
                        </div>
                        <div class="fieldset">
                            <label for="join-email" class="fieldset-label">이메일주소</label>
                            <div class="fieldset-row">
                                <span class="input is-expand">
                                    <input type="email"name="Email" id="Email" value="<%=Email%>" placeholder="이메일주소를 입력해 주세요.">
                                </span>
                            </div>
                        </div>
                        <!-- *** 수정 *** 190110 : 주소 입력창 추가 -->
                        <div class="fieldset ty-col2">
                            <label for="enter-addr" class="fieldset-label">주소</label>
                            <div class="fieldset-row">
                                <button type="button" class="button ty-black" onclick="execDaumPostcode('ZipCode','Addr1');">우편번호 찾기</button>
                                <span class="input">
                                    <input type="text" name="ZipCode" id="ZipCode" placeholder="우편번호 검색" readonly="readonly">
                                </span>
                                <span class="input is-expand double">
                                        <input type="text" name="Addr1" id="Addr1" readonly="readonly">
                                </span>
                                <span class="input is-expand double">
                                        <input type="text" name="Addr2" id="Addr2" placeholder="상세주소 입력">
                                </span>
                            </div>
                        </div>
                        <!-- // *** 수정 *** 190110 : 주소 입력창 추가 -->
                    </fieldset>
					<%IF JoinType = "U" THEN%>
                    <button type="button"  onclick="chk_Join();" class="button is-expand ty-red">가입하기</button>
					<%END IF%>
                </div>
				<%IF JoinType = "D" THEN%>
                <!-- 14세 미만 가입일 경우 보호자 정보 입력 -->
                <div class="parent-info">
                    <div class="formfield">
                        <div class="fieldset parent-inf -write">
                            <p class="tit">보호자 정보</p>
                            <p class="message">*만 14세 미만 가입 시 필수 기재사항입니다.</p>
                        </div>
                        <fieldset class="">
                            <legend class="hidden">기본 정보 입력</legend>
                            <div class="fieldset">
                                <label for="pjoin-id" class="fieldset-label">이름</label>
                                <div class="fieldset-row">
                                    <span class="input is-expand">
										<%IF ParentName = "" THEN%>
										<input type="text" name="ParentName" id="ParentName" maxlength="25" placeholder="이름을 입력해 주세요.">
										<%ELSE%>
										<input type="text" name="ParentName" id="ParentName" maxlength="25" value="<%=ParentName%>" readonly="readonly" placeholder="이름을 입력해 주세요.">
										<%END IF%>
                                    </span>
                                </div>
                            </div>
                            <div class="fieldset">
                                <label for="join-pw" class="fieldset-label">생년월일</label>
                                <div class="fieldset-row">
                                    <span class="input is-expand">
										<%IF ParentBirth = "" THEN%>
										<input type="tel" name="ParentBirth" id="ParentBirth" maxlength="8" placeholder="생년월일을 입력해 주세요.(ex. 19991231)">
										<%ELSE%>
										<input type="tel" name="ParentBirth" id="ParentBirth" maxlength="8" value="<%=ParentBirth%>" readonly="readonly" class="input-no-border" placeholder="생년월일을 입력해 주세요.(ex. 19991231)">
										<%END IF%>
                                    </span>
                                </div>
                            </div>
                            <div class="fieldset">
                                <legend class="hidden">인증 정보 입력</legend>
                                <div class="fieldset ty-col2 pt0">
                                    <label for="pjoin-phone" class="fieldset-label">휴대폰 번호</label>
                                    <div class="fieldset-row">
										<%IF ParentMobile = "" THEN%>
										<span class="pselect2">
											<select name="PHP1" id="PHP1" title="휴대폰 국번 선택">
												<option value="010">010</option>
												<option value="011">011</option>
												<option value="016">016</option>
												<option value="017">017</option>
												<option value="018">018</option>
												<option value="019">019</option>
											</select>
											<span class="value"></span>
										</span>
										<span class="pdash1">-</span>
										<span class="pinput2">
											<input type="tel" name="PHP2" id="PHP2" maxlength="4">
										</span>
										<span class="pdash2">-</span>
										<span class="pinput3">
											<input type="tel" name="PHP3" id="PHP3" maxlength="4">
										</span>
										<input type="hidden" name="ParentMobileFlag" value="N" />
										<%ELSE%>
										<span class="pinput1">
											<input type="tel" name="PHP1" id="PHP1" maxlength="4" value="<%=PHP1%>" readonly="readonly">
										</span>
										<span class="pdash1">-</span>
										<span class="pinput2">
											<input type="tel" name="PHP2" id="PHP2" maxlength="4" value="<%=PHP2%>" readonly="readonly">
										</span>
										<span class="pdash2">-</span>
										<span class="pinput3">
											<input type="tel" name="PHP3" id="PHP3" maxlength="4" value="<%=PHP3%>" readonly="readonly">
										</span>
										<input type="hidden" name="ParentMobileFlag" value="Y" />
										<%END IF%>
                                    </div>
                                </div>
                            </div>
                            <div class="fieldset">
                                <label for="pjoin-email" class="fieldset-label">이메일주소</label>
                                <div class="fieldset-row">
                                    <span class="input is-expand">
                                        <input type="email" name="ParentEmail" id="ParentEmail" placeholder="이메일 주소를 입력해 주세요.">
                                    </span>
                                </div>
                            </div>
                        </fieldset>
	                    <button type="button"  onclick="chk_Join();" class="button is-expand ty-red">가입하기</button>
                    </div>
                </div>
				<%END IF%>
            </section>
			</form>
        </div>
    </main>


	<script>
		$(function() {
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

	<%IF JoinType = "D" THEN%>
	<script>
		$(function() {
			$(".pselect2 .value").text($("select[name='PHP1']").val());

			$(".pselect2 select").on("focus", function() {
				$(".pselect2").addClass("is-focus");
			});
			$(".pselect2 select").on("blur", function() {
				$(".pselect2").removeClass("is-focus");
			});
			$(".pselect2 select").on("change", function() {
				$(".pselect2 .value").text($("select[name='PHP1']").val());
			});
		
			$(".pinput2 input").on("focus", function() {
				$(".pinput2").addClass("is-focus");
			});
			$(".pinput2 input").on("blur", function() {
				$(".pinput2").removeClass("is-focus");
			});
		
			$(".pinput3 input").on("focus", function() {
				$(".pinput3").addClass("is-focus");
			});
			$(".pinput3 input").on("blur", function() {
				$(".pinput3").removeClass("is-focus");
			});
		
			$(".pinput1 input").on("focus", function() {
				$(".pinput1").addClass("is-focus");
			});
			$(".pinput1 input").on("blur", function() {
				$(".pinput1").removeClass("is-focus");
			});
		});
	</script>
	<%END IF%>

<!-- #include virtual="/INC/FooterNoBNB.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->



<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>