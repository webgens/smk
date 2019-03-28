<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'SnsList.asp - 마이페이지 > 회원정보 > SNS 계정설정
'Date		: 2018.12.17
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
PageCode1 = "05"
PageCode2 = "05"
PageCode3 = "04"
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->
<!-- #include virtual="/Common/SubCheckID.asp" -->

<%

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



Dim NaverLogin : NaverLogin = "N"
Dim GoogleLogin : GoogleLogin = "N"
Dim FacebookLogin : FacebookLogin = "N"
Dim KakaoLogin : KakaoLogin = "N"

Dim NUID, FUID, GUID, KUID
Dim NEmail, FEmail, GEmail, KEmail
Dim NDate, FDate, GDate, KDate
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성




'*****************************************************************************************'
'회원 기본정보 SELECT START
'-----------------------------------------------------------------------------------------'
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
	.ActiveConnection = oConn
	.CommandType = adCmdStoredProc
	.CommandText = "USP_Front_EShop_Member_SNS_Select_By_MemberNum"
	.Parameters.Append .CreateParameter("@MemberNum", adInteger, adParamInput, , U_NUM)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing

Do While Not oRs.eof
	Select Case oRs("SNSKind")
		Case "N" 
			NaverLogin = "Y"
			NUID = oRs("SNSID")
			NEmail = oRs("Email")
			NDate = oRs("CreateDT")
		Case "F" 
			FacebookLogin = "Y"
			FUID = oRs("SNSID")
			FEmail = oRs("Email")
			FDate = oRs("CreateDT")
		Case "K" 
			KakaoLogin = "Y"
			KUID = oRs("SNSID")
			KEmail = oRs("Email")
			KDate = oRs("CreateDT")
		Case "G" 
			GoogleLogin = "Y"
			GUID = oRs("SNSID")
			GEmail = oRs("Email")
			GDate = oRs("CreateDT")
	End Select

	oRs.MoveNext
Loop
oRs.Close
'-----------------------------------------------------------------------------------------'
'회원 기본정보 SELECT END
'-----------------------------------------------------------------------------------------'
%>


<!-- #include virtual="/INC/Header.asp" -->
	<style type="text/css">
		#OrderMenu .selector { margin-bottom: 0; }
		#OrderMenu .selector.is-focus .btn-list:after { background: url("/images/ico/ico_arrow_u2.png")no-repeat; background-size: 100% auto; }
	</style>
    <script type="text/javascript" src="/JS/dev/mypage.js?ver=<%=U_DATE%><%=U_TIME%>"></script>
	<script type="text/javascript" src="/JS/dev/login.js?ver=<%=U_DATE%><%=U_TIME%>"></script>
	<script type="text/javascript">
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

				snsConnection();
			}
		}

		function SnsLoginDel(uid) {
			$.ajax({
				url			 : '/ASP/Mypage/Ajax/MySnsDeleteOk.asp',
				data		 : "UID="+uid,
				async		 : false,
				type		 : 'post',
				dataType	 : 'html',
				success		 : function (data) {	
								var splitData	 = data.split("|||||");
								var result		 = splitData[0];
								var msg			 = splitData[1];

								if (result == "OK") {
									openAlertLayer("alert", "해제 되었습니다.", "closePop('alertPop', '');location.reload();", "");
									return;
								}
								else if (result == "FAIL") {
									openAlertLayer("alert", msg, "closePop('alertPop', '');", "");
									return;
								}
								else {
									openAlertLayer("alert", "오류로 인하여 해제되지 않았습니다.<br />다시 확인 하여 주세요.", "closePop('alertPop', '');", "");
									return;
								}
				},
				error		 : function (data) {
								openAlertLayer("alert", "처리중 오류가 발생하였습니다.", "closePop('alertPop', '');", "");
				}
			});
		}
	</script>

<%TopSubMenuTitle = "회원정보"%>
<!-- #include virtual="/INC/TopSub.asp" -->


    <!-- Main -->
    <main id="container" class="container">
        <div class="sub_content">

            <div class="wrap-mypage">
				<div style="height:8px"></div>


				
                <div id="OrderMenu" class="ly-title accordion">
                    <div class="selector">
	                    <button type="button" class="btn-list clickEvt" data-target="OrderMenu">SNS 계정설정</button>
					</div>
					<div class="option my-recode">
						<ul>
							<li><a href="/ASP/Mypage/MyMemberShip.asp">나의 멤버십</a></li>
							<li><a href="/ASP/Mypage/AddressList.asp">배송지관리</a></li>
							<li><a href="javascript:common_PopOpen('DimDepth1','MyInfoModify');">나의 정보 수정</a></li>
							<li><a href="/ASP/Mypage/SnsList.asp">SNS 계정설정</a></li>
						</ul>
					</div>
                </div>



                <div class="mypage-membership">
                    <section id="contentList_4" class="accord-mypage">

                        <div class="ly-content1" id="getMySnsList">


                            <div class="sns-login">
                                <%IF NaverLogin="Y" THEN%>
                                <div class="sns naver">
                                    <div class="logo">
                                        <span class="hidden">NAVER</span>
                                        <div class="mypage">
                                            <span class="badge">연결</span>
                                        </div>
                                    </div>
                                    <div class="btn">
                                        <p><%=NEmail%></p>
                                        <button type="button" class="disconnect" onclick="SnsLoginDel('<%=NUID%>');">해제</button>
                                    </div>
                                </div>
								<%ELSE%>
                                <div class="sns naver">
                                    <div class="logo">
                                        <span class="hidden">NAVER</span>
									</div>
                                    <div class="btn">
                                        <button type="button" class="connect" onclick="pop_NaverLogin();">연결 하기</button>
                                    </div>
                                </div>
								<%END IF%>

                                <%IF FacebookLogin="Y" THEN%>
                                <div class="sns facebook">
                                    <div class="logo">
                                        <span class="hidden">facebook</span>
                                        <div class="mypage">
                                            <span class="badge">연결</span>
                                        </div>
                                    </div>
                                    <div class="btn">
                                        <p><%=FEmail%></p>
                                        <button type="button" class="disconnect" onclick="SnsLoginDel('<%=FUID%>');">해제</button>
                                    </div>
                                </div>
								<%ELSE%>
                                <div class="sns facebook">
                                    <div class="logo"><span class="hidden">facebook</span></div>
                                    <div class="btn">
                                        <button type="button" class="connect" onclick="pop_FacebookLogin();">연결 하기</button>
                                    </div>
                                </div>
								<%END IF%>

                                <%IF KakaoLogin="Y" THEN%>
                                <div class="sns kakao">
                                    <div class="logo">
                                        <span class="hidden">kakao</span>
                                        <div class="mypage">
                                            <span class="badge">연결</span>
                                        </div>
                                    </div>
                                    <div class="btn">
                                        <p><%=KEmail%></p>
                                        <button type="button" class="disconnect" onclick="SnsLoginDel('<%=KUID%>');">해제</button>
                                    </div>
                                </div>
								<%ELSE%>
                                <div class="sns kakao">
                                    <div class="logo"><span class="hidden">kakao</span></div>
                                    <div class="btn">
                                        <button type="button" class="connect" onclick="pop_KakaoLogin();">연결 하기</button>
                                    </div>
                                </div>
								<%END IF%>

                                <%IF GoogleLogin="Y" THEN%>
                                <!--<div class="sns google">
                                    <div class="logo">
                                        <span class="hidden">google</span>
                                        <div class="mypage">
                                            <span class="badge">연결</span>
                                        </div>
                                    </div>
                                    <div class="btn">
                                        <p><%'=GEmail%></p>
                                        <button type="button" class="disconnect" onclick="SnsLoginDel('<%'=GUID%>');">해제</button>
                                    </div>
                                </div>-->
								<%ELSE%>
                                <!--<div class="sns google">
                                    <div class="logo"><span class="hidden">google</span></div>
                                    <div class="btn">
                                        <button type="button" class="connect" onclick="pop_GoogleLogin();">연결 하기</button>
                                    </div>
                                </div>-->
								<%END IF%>
                            </div>
                            <div class="inf-type1">
                                <p class="tit">알려드립니다.</p>
                                <ul>
                                    <li class="bullet-ty1">SNS계정과 연동하여 간편하게 로그인 할 수 있는 서비스 입니다.</li>
                                    <li class="bullet-ty1">SNS계정은 중복하여 이용하실 수 없습니다.</li>
                                    <li class="bullet-ty1">계정 연결 해제 시는 기존 아이디와 패스워드로 로그인하여 이용할 수 있습니다.</li>
								</ul>
                            </div>



                        </div>
                    </section>
                </div>
            </div>
        </div>
    </main>


	<!-- SNS계정 연결 공통 시작 -->
	<!-- SNS계정 로그인 Form -->
 	<form name="SimpleLoginForm" id="SimpleLoginForm" method="post">
		<input type="hidden" name="UID">
		<input type="hidden" name="Email">
		<input type="hidden" name="KName">
		<input type="hidden" name="SNSKind">
	</form>


<!-- #include virtual="/INC/FooterNoBNB.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->


<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>
