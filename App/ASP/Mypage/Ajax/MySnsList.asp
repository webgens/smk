<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'index.asp - 마이페이지
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
PageCode4 = "01"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->
<!-- #include virtual="/Common/CheckID_Ajax.asp" -->

<%

'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn											'# ADODB Connection 개체
DIM oRs												'# ADODB Recordset 개체
DIM oCmd											'# ADODB Command 개체

DIM i
DIM j
DiM x
DIM y

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절


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

'/****************************************************************************************/
'회원 기본정보 SELECT START
'-----------------------------------------------------------------------------------------------------------'
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
'-----------------------------------------------------------------------------------------------------------'
'회원 기본정보 SELECT END
'-----------------------------------------------------------------------------------------------------------'

Response.Write "OK|||||"
%>


                            <div class="h-line">
                                <h2 class="h-level4">SNS로그인 설정</h2>
                            </div>
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
                                <div class="sns google">
                                    <div class="logo">
                                        <span class="hidden">google</span>
                                        <div class="mypage">
                                            <span class="badge">연결</span>
                                        </div>
                                    </div>
                                    <div class="btn">
                                        <p><%=GEmail%></p>
                                        <button type="button" class="disconnect" onclick="SnsLoginDel('<%=GUID%>');">해제</button>
                                    </div>
                                </div>
								<%ELSE%>
                                <div class="sns google">
                                    <div class="logo"><span class="hidden">google</span></div>
                                    <div class="btn">
                                        <button type="button" class="connect" onclick="pop_GoogleLogin();">연결 하기</button>
                                    </div>
                                </div>
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


<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>