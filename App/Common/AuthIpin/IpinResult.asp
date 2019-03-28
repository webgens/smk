<%@ Language=VBScript codepage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'IpinResult.asp - IPIN 본인인증 결과
'Date		: 2018.11.06
'Update		: 
'*****************************************************************************************'
	
'# 페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include Virtual = "/ADO/ADODBCommon.asp" -->
<!-- #include Virtual = "/Common/Common.asp" -->
<!-- #include Virtual = "/Common/SetInfo.asp" -->


<!-- #include virtual="/INC/Header.asp" -->

<%
'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn							'# ADODB Connection 개체
DIM oRs								'# ADODB Recordset 개체
DIM oCmd							'# ADODB Command 개체

DIM wQuery							'# WHERE 절
DIM sQuery							'# SORT 절

DIM sSiteCode
DIM sSitePw
DIM sResponseData
DIM sCPRequest
DIM IPIN_DLL
DIM clsIPINDll
DIM iRtn
DIM sRtnMsg
DIM sVNumber
DIM sName
DIM sDupInfo
DIM sAgeCode
DIM sGenderCode
DIM sBirthDate
DIM sNationalInfo
DIM sCPRequestNum
DIM sReserved1
DIM sReserved2
DIM sReserved3
DIM sForeigner

DIM Age

DIM SFlag : SFlag = "F"
DIM sErrCode

DIM SMode
DIM JoinType
DIM InName
DIM InUserID

DIM sErrType

DIM DB_MemberNum
DIM DB_UserID
DIM DB_Name
DIM DB_HP
DIM DB_GroupCode
DIM DB_SDupInfo
DIM DB_ParentSDupInfo
DIM DB_FTFlag


DIM MemberNum
DIM SaveID
DIM NewAgreementFlag
DIM ProgID
DIM OrderFlag


DIM CouponName
DIM CouponCount : CouponCount = 0

DIM StartDT
DIM EndDT

DIM Message
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


SMode				 = Request("param_r1")			'# 인증목적 : MemberJoin(회원가입) / SearchID(아이디찾기) / SearchPwd(비밀번호찾기) / DormancyRelease(휴면계정해제)
JoinType			 = Request("param_r2")

IF SMode = "" THEN
		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=본인 인증 목적 값이 없습니다.<br />다시 시도하여 주십시오.&Script=APP_PopupHistoryBack();"
		Response.End
END IF


IF SMode = "SearchPwd" THEN
		InName			 = TRIM(Decrypt(Request.Cookies("SW_NAME")))
		InUserID		 = TRIM(Decrypt(Request.Cookies("SW_USERID")))

		IF InName = "" OR InUserID = "" THEN
				Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=입력하신 이름, 아이디 정보가 없습니다.<br />다시 시도하여 주십시오.&Script=APP_PopupHistoryBack();"
				Response.End
		END IF
END IF
	

sSiteCode		= IPIN_H_ID				'NICE로부터 부여받은 사이트 코드
sSitePw			= IPIN_H_PWD			'NICE로부터 부여받은 사이트 패스워드
		
'사용자 정보 및 CP 요청번호를 암호화한 데이타입니다.
sResponseData	 = Request("enc_data")
    
'CP 요청번호 : ipin_main.asp 에서 세션 처리한 데이타
sCPRequest		 = Request.Cookies("CPREQUEST")
    
'컴포넌트 객체 생성
SET clsIPINDll	 = Server.CreateObject("IPINClient.Kisinfo")
iRtn = clsIPINDll.fnResponse(sSiteCode, sSitePw, sResponseData)
	
'Method 결과값에 따른 처리사항
IF (iRtn = 1) THEN
		
		sVNumber			= clsIPINDll.bstrVNumber				'가상주민번호 (13자리이며, 숫자 또는 문자 포함)
		sName				= clsIPINDll.bstrName					'이름
		sDupInfo			= clsIPINDll.bstrDupInfo				'중복가입 확인값 (DI - 64 byte 고유값)
		sAgeCode			= clsIPINDll.bstrAgeCode				'연령대 코드 (개발 가이드 참조)
		sGenderCode			= clsIPINDll.bstrGenderCode				'성별 코드 (개발 가이드 참조)
		sBirthDate			= clsIPINDll.bstrBirthDate				'생년월일 (YYYYMMDD)
		sNationalInfo		= clsIPINDll.bstrNationalInfo			'내/외국인 정보 (개발 가이드 참조)
		sCPRequestNum		= clsIPINDll.bstrCPRequestNUM			'CP 요청번호

		SET oConn		 = ConnectionOpen()							'# 커넥션 생성

		'# 정보 입력
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Admin_EShop_sDusInfo_Log_Insert"

				.Parameters.Append .CreateParameter("@sDusInfo",	 adVarChar,	 adParamInput,   255, sDupInfo)
				.Parameters.Append .CreateParameter("@sName",		 adVarChar,	 adParamInput,    50, sName)
				
				.Execute, , adExecuteNoRecords
		END WITH
		SET oCmd = Nothing

		oConn.Close
		SET oConn = Nothing


		'# 회원가입
		IF SMode = "MemberJoin" THEN



				Response.Cookies("JoinType")			 = ""
				Response.Cookies("AuthType")			 = ""
				Response.Cookies("SDupInfo")			 = ""
				Response.Cookies("Name")				 = ""
				Response.Cookies("Birth")				 = ""
				Response.Cookies("Gender")				 = ""
				Response.Cookies("NationalInfo")		 = ""
				Response.Cookies("Mobile")				 = ""
				Response.Cookies("ParentSDupInfo")		 = ""
				Response.Cookies("ParentName")			 = ""
				Response.Cookies("ParentBirth")			 = ""
				Response.Cookies("ParentGender")		 = ""
				Response.Cookies("ParentNationalInfo")	 = ""
				Response.Cookies("ParentMobile")		 = ""



				Age = CDbl(U_DATE) - CDbl(sBirthDate)

				'# 14세 이상
				IF JoinType = "U" THEN

						IF Age < 140000 THEN
%>
								<script type="text/javascript">
									APP_PopupHistoryBack_JoinAuth_Result("UNDER14", "<%=sName%>님은 만 14세 미만입니다.<br />만 14세 미만 회원가입으로 보호자 동의를 받아 주십시오.");
								</script>
<%
								Response.End
						END IF


						SET oConn		 = ConnectionOpen()							'# 커넥션 생성
						SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
	

						'# 회원정보
						SET oCmd = Server.CreateObject("ADODB.Command")
						WITH oCmd
								.ActiveConnection	 = oConn
								.CommandType		 = adCmdStoredProc
								.CommandText		 = "USP_Front_EShop_Member_Select_By_SDupInfo_Check"

								.Parameters.Append .CreateParameter("@SDupInfo",	 adVarChar,	 adParamInput, 64, sDupInfo)
						END WITH
						oRs.CursorLocation = adUseClient
						oRs.Open oCmd, , adOpenStatic, adLockReadOnly
						SET oCmd = Nothing



						IF NOT oRs.EOF THEN

								oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
%>
								<script type="text/javascript">
									APP_PopupHistoryBack_JoinAuth_Result("MEMBER", "<%=sName%>님은 이미 회원으로 가입되어 있습니다.<br />아이디/비밀번호 찾기를 이용해 주십시오.");
								</script>
<%
								Response.End

						ELSE
	
								Response.Cookies("JOIN_TYPE")		 = ""
								Response.Cookies("JoinType")		 = Encrypt(JoinType)
								Response.Cookies("AuthType")		 = Encrypt("I")
								Response.Cookies("SDupInfo")		 = Encrypt(sDupInfo)
								Response.Cookies("Name")			 = Encrypt(sName)
								Response.Cookies("Birth")			 = Encrypt(sBirthDate)
								Response.Cookies("Gender")			 = Encrypt(sGenderCode)
								'# Response.Cookies("NationalInfo") = Encrypt(sNationalInfo)
								Response.Cookies("Mobile")			 = Encrypt("")

						END IF
						oRs.Close

						SET oRs = Nothing : oConn.Close : SET oConn = Nothing
%>
						<script type="text/javascript">
							APP_PopupHistoryBack_JoinAuth_Result("OK", "");
						</script>
<%
						Response.End


				'# 만14세 미만
				ELSE

						IF Age < 200000 THEN
%>
								<script type="text/javascript">
									APP_PopupHistoryBack_JoinAuth_Result("UNDER20", "만 20세 이상만 보호자 인증을 받을 수 있습니다.");
								</script>
<%
								Response.End
						END IF

						Response.Cookies("JOIN_TYPE")			 = ""
						Response.Cookies("JoinType")			 = Encrypt(JoinType)
						Response.Cookies("AuthType")			 = Encrypt("I")
						Response.Cookies("ParentSDupInfo")		 = Encrypt(sDupInfo)
						Response.Cookies("ParentName")			 = Encrypt(sName)
						Response.Cookies("ParentBirth")			 = Encrypt(sBirthDate)
						Response.Cookies("ParentGender")		 = Encrypt(sGenderCode)
						'# Response.Cookies("ParentNationalInfo")	 = Encrypt(sNationalInfo)
						Response.Cookies("ParentMobile")		 = Encrypt("")
%>
						<script type="text/javascript">
							APP_PopupHistoryBack_JoinAuth_Result("OK", "");
						</script>
<%
						Response.End

				END IF




		'# 아이디찾기
		ELSEIF SMode = "SearchID" THEN



				SET oConn		 = ConnectionOpen()							'# 커넥션 생성
				SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성


				'# 회원정보
				SET oCmd = Server.CreateObject("ADODB.Command")
				WITH oCmd
						.ActiveConnection	 = oConn
						.CommandType		 = adCmdStoredProc
						.CommandText		 = "USP_Front_EShop_Member_Select_By_SDupInfo_Check"

						.Parameters.Append .CreateParameter("@SDupInfo",	 adVarChar,	 adParamInput, 64, sDupInfo)
				END WITH
				oRs.CursorLocation = adUseClient
				oRs.Open oCmd, , adOpenStatic, adLockReadOnly
				SET oCmd = Nothing

				IF NOT oRs.EOF THEN
						DB_UserID		 = oRs("UserID")
						DB_Name			 = oRs("Name")
						DB_HP			 = REPLACE(oRs("HP"), "-", "")
				ELSE
						oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
%>
						<script type="text/javascript">
							APP_PopupHistoryBack_ID_IP_Result("NOTEXISTS", "일치하는 회원이 없습니다.");
						</script>
<%
						Response.End
				END IF
				oRs.Close

				SET oRs = Nothing : oConn.Close : SET oConn = Nothing
%>
				<script type="text/javascript">
					APP_PopupHistoryBack_ID_IP_Result("OK", "<%=sName%>님의 아이디는 <%=MaskUserID(DB_UserID)%> 입니다.");
				</script>
<%
				Response.End


		'# 비밀번호찾기
		ELSEIF SMode = "SearchPwd" THEN


				Response.Cookies("SW_NAME")		 = ""
				Response.Cookies("SW_USERID")	 = ""
	

				SET oConn		 = ConnectionOpen()							'# 커넥션 생성
				SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성


				'# 회원정보
				SET oCmd = Server.CreateObject("ADODB.Command")
				WITH oCmd
						.ActiveConnection	 = oConn
						.CommandType		 = adCmdStoredProc
						.CommandText		 = "USP_Front_EShop_Member_Select_By_SDupInfo_Check"

						.Parameters.Append .CreateParameter("@SDupInfo",	 adVarChar,	 adParamInput, 64, sDupInfo)
				END WITH
				oRs.CursorLocation = adUseClient
				oRs.Open oCmd, , adOpenStatic, adLockReadOnly
				SET oCmd = Nothing

				IF NOT oRs.EOF THEN
						DB_MemberNum	 = oRs("MemberNum")
						DB_UserID		 = oRs("UserID")
						DB_Name			 = oRs("Name")
						DB_HP			 = REPLACE(oRs("HP"), "-", "")
				ELSE
						oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
%>
						<script type="text/javascript">
							APP_PopupHistoryBack_PW_HP_Result("NOTEXISTS", "일치하는 회원이 없습니다.");
						</script>
<%
						Response.End
				END IF
				oRs.Close

				IF InName <> DB_Name THEN
						SET oRs = Nothing : oConn.Close : SET oConn = Nothing
%>
						<script type="text/javascript">
							APP_PopupHistoryBack_PW_HP_Result("NOTEXISTS", "이름이 일치하는 않습니다.");
						</script>
<%
						Response.End
				END IF

				IF InUserID <> DB_UserID THEN
						SET oRs = Nothing : oConn.Close : SET oConn = Nothing
%>
						<script type="text/javascript">
							APP_PopupHistoryBack_PW_HP_Result("NOTEXISTS", "아이디가 일치하는 않습니다.");
						</script>
<%
						Response.End
				END IF

				SET oRs = Nothing : oConn.Close : SET oConn = Nothing

				Response.Cookies("PW_MemberNum") = DB_MemberNum
%>
				<script type="text/javascript">
					APP_PopupHistoryBack_PW_HP_Result("OK", "<%=DB_UserID%>");
				</script>
<%
				Response.End


		'# 휴면계정해제
		ELSEIF SMode = "DormancyRelease" THEN

	
				MemberNum		 = TRIM(Decrypt(Request.Cookies("TEMP_UNUM")))
				SaveID			 = TRIM(Decrypt(Request.Cookies("TEMP_UIDSAVE")))
				OrderFlag		 = TRIM(Decrypt(Request.Cookies("TEMP_ORDERFLAG")))
				ProgID			 = TRIM(Decrypt(Request.Cookies("TEMP_PROGID")))
				IF ProgID		 = "" THEN ProgID = "/"

				SET oConn		 = ConnectionOpen()							'# 커넥션 생성
				SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성




				'# 휴면계정정보
				SET oCmd = Server.CreateObject("ADODB.Command")
				WITH oCmd
						.ActiveConnection	 = oConn
						.CommandType		 = adCmdStoredProc
						.CommandText		 = "USP_Front_EShop_Member_Dormancy_Select_By_MemberNum"

						.Parameters.Append .CreateParameter("@MemberNum",	 adInteger,	 adParamInput,   , MemberNum)
				END WITH
				oRs.CursorLocation = adUseClient
				oRs.Open oCmd, , adOpenStatic, adLockReadOnly
				SET oCmd = Nothing

				IF NOT oRs.EOF THEN
						DB_Name				= oRs("Name")
						DB_HP				= REPLACE(oRs("HP"), "-", "")
						DB_SDupInfo			= oRs("SDupInfo")
						DB_ParentSDupInfo	= oRs("ParentSDupInfo")
						DB_GroupCode		= oRs("GroupCode")
						DB_FTFlag			= oRs("FTFlag")
						IF ISNULL(DB_SDupInfo)			THEN DB_SDupInfo = ""
						IF ISNULL(DB_ParentSDupInfo)	THEN DB_ParentSDupInfo = ""
						IF ISNULL(DB_FTFlag)			THEN DB_FTFlag = "N"
				ELSE
						oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
%>
						<script type="text/javascript">
							APP_PopupHistoryBack_DOR_Result("DOR_NOTEXISTS", "<%=sName%>의 휴면계정 정보가 없습니다.<br />다시 로그인하여 주십시오.");
						</script>
<%
						Response.End
				END IF
				oRs.Close

				IF DB_Name <> sName THEN
						SET oRs = Nothing : oConn.Close : SET oConn = Nothing
%>
						<script type="text/javascript">
							APP_PopupHistoryBack_DOR_Result("NOTMATCH", "<%=sName%>님과 회원정보의 이름이 일치하지 않습니다.");
						</script>
<%
						Response.End
				END IF

				'# 아이핀은 핸드폰번호 없음
				'# IF DB_HP <> sMobileNo THEN
				'# 		SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				'# 
				'# 		Call AlertMessage(sName & "님의 핸드폰번호와 회원정보의 핸드폰번호가 일치하지 않습니다.", "self.close();")
				'# 		Response.End
				'# END IF

				IF (DB_SDupInfo <> sDupInfo AND DB_FTFlag = "N") OR (DB_ParentSDupInfo  <> sDupInfo AND DB_FTFlag = "Y") THEN
						SET oRs = Nothing : oConn.Close : SET oConn = Nothing
%>
						<script type="text/javascript">
							APP_PopupHistoryBack_DOR_Result("NOTMATCH", "<%=sName%>님의 인증값과<br />회원정보의 인증값이 일치하지 않습니다.");
						</script>
<%
						Response.End
				END IF
		


				'# TRANSACTION START
				oConn.BeginTrans
	

				'# 휴면계정해제 정보 입력
				SET oCmd = Server.CreateObject("ADODB.Command")
				WITH oCmd
						.ActiveConnection	 = oConn
						.CommandType		 = adCmdStoredProc
						.CommandText		 = "USP_Front_EShop_Member_Dormancy_Log_Insert"

						.Parameters.Append .CreateParameter("@MemberNum",	 adInteger,	 adParamInput,   , MemberNum)
						.Parameters.Append .CreateParameter("@Location",	 adChar,	 adParamInput,  1, "P")
						.Parameters.Append .CreateParameter("@CreateIP",	 adVarChar,	 adParamInput, 15, U_IP)
				
						.Execute, , adExecuteNoRecords
				END WITH
				SET oCmd = Nothing


				'# 회원 테이블에 공백 컬럼 휴면계정에 있는 데이터로 업데이트
				SET oCmd = Server.CreateObject("ADODB.Command")
				WITH oCmd
						.ActiveConnection	 = oConn
						.CommandType		 = adCmdStoredProc
						.CommandText		 = "USP_Front_EShop_Member_Update_For_Dormancy_Release"

						.Parameters.Append .CreateParameter("@MemberNum",	 adInteger, adParamInput,   , MemberNum)
						.Parameters.Append .CreateParameter("@UpdateID",	 adVarChar, adParamInput, 20, MemberNum)
						.Parameters.Append .CreateParameter("@UpdateIP",	 adVarChar, adParamInput, 15, U_IP)
				
						.Execute, , adExecuteNoRecords
				END WITH
				SET oCmd = Nothing


				'# 휴면계정정보 삭제
				SET oCmd = Server.CreateObject("ADODB.Command")
				WITH oCmd
						.ActiveConnection	 = oConn
						.CommandType		 = adCmdStoredProc
						.CommandText		 = "USP_Front_EShop_Member_Dormancy_Delete"

						.Parameters.Append .CreateParameter("@MemberNum",	 adInteger, adParamInput,   , MemberNum)
				
						.Execute, , adExecuteNoRecords
				END WITH
				SET oCmd = Nothing


				'# 휴면계정해제 쿠키
				Response.Cookies("TEMP_DOR")		 = Encrypt("N")

		
				NewAgreementFlag	 = TRIM(Decrypt(Request.Cookies("TEMP_NEW")))



				oConn.CommitTrans



				'# 신규 약관 동의
				IF NewAgreementFlag = "Y" THEN


						'# 배포된 쿠폰 발급 받기 - EShop_Coupon_Member 에 배포되었지만 아직 안받은 쿠폰
						SET oCmd = Server.CreateObject("ADODB.Command")
						WITH oCmd
								.ActiveConnection	 = oConn
								.CommandType		 = adCmdStoredProc
								.CommandText		 = "USP_Front_EShop_Coupon_Member_Select_For_Not_Receive"
	
								.Parameters.Append .CreateParameter("@MemberNum",	 adInteger,	 adParamInput,   , MemberNum)
						END WITH
						oRs.CursorLocation = adUseClient
						oRs.Open oCmd, , adOpenStatic, adLockReadOnly
						SET oCmd = Nothing

						IF NOT oRs.EOF THEN
								CouponName	 = oRs("CouponName")
								CouponCount	 = CouponCount + oRs.RecordCount

								SET oCmd = Server.CreateObject("ADODB.Command")
								WITH oCmd
										.ActiveConnection	 = oConn
										.CommandType		 = adCmdStoredProc
										.CommandText		 = "USP_Front_EShop_Coupon_Member_Update_For_Receive"
	
										.Parameters.Append .CreateParameter("@MemberNum",	 adInteger,	 adParamInput,   , MemberNum)
										.Parameters.Append .CreateParameter("@UpdateID",	 adVarChar,	 adParamInput, 20, MemberNum)
										.Parameters.Append .CreateParameter("@UpdateIP",	 adVarChar,	 adParamInput, 15, U_IP)

										.Execute, , adExecuteNoRecords
								END WITH
								SET oCmd = Nothing
						END IF
						oRs.Close


						'# 배포될 쿠폰 발급 받기 - EShop_Coupon_Member 에 데이터가 없고 EShop_Coupon 에 내가 배포 대상인 쿠폰
						SET oCmd = Server.CreateObject("ADODB.Command")
						WITH oCmd
								.ActiveConnection	 = oConn
								.CommandType		 = adCmdStoredProc
								.CommandText		 = "USP_Front_EShop_Coupon_Select_For_Not_Receive"
	
								.Parameters.Append .CreateParameter("@MemberNum",	 adInteger,	 adParamInput,   , MemberNum)
								.Parameters.Append .CreateParameter("@GroupCode",	 adVarChar,	 adParamInput, 10, DB_GroupCode)
	
						END WITH
						oRs.CursorLocation = adUseClient
						oRs.Open oCmd, , adOpenStatic, adLockReadOnly
						SET oCmd = Nothing

						IF NOT oRs.EOF THEN
								IF CouponName = "" THEN CouponName	 = oRs("CouponName")
								CouponCount	 = CouponCount + oRs.RecordCount


								Do Until oRs.EOF
										IF oRs("UseDateType") = "U" THEN
												StartDT	 = U_DATE & R_HOUR & "0000"
												EndDT	 = "99999999999999"
										ELSEIF oRs("UseDateType") = "P" THEN
												StartDT	 = U_DATE & R_HOUR & "0000"
												EndDT	 = oRs("UseEDate")
										ELSE
												StartDT	 = U_DATE & "000000"
												EndDT	 = REPLACE(DATEADD("d", oRs("UseDay"), Date), "-", "") & "240000"
										END IF

						
										SET oCmd = Server.CreateObject("ADODB.Command")
										WITH oCmd
												.ActiveConnection	 = oConn
												.CommandType		 = adCmdStoredProc
												.CommandText		 = "USP_Front_EShop_Coupon_Member_Insert"
	
												.Parameters.Append .CreateParameter("@MemberNum",	 adInteger,	 adParamInput,   , MemberNum)
												.Parameters.Append .CreateParameter("@CouponIdx",	 adBigInt,	 adParamInput,   , oRs("Idx"))
												.Parameters.Append .CreateParameter("@StartDT",		 adVarChar,	 adParamInput, 14, StartDT)
												.Parameters.Append .CreateParameter("@EndDT",		 adVarChar,	 adParamInput, 14, EndDT)
												.Parameters.Append .CreateParameter("@CreateID",	 adVarChar,	 adParamInput, 20, MemberNum)
												.Parameters.Append .CreateParameter("@CreateIP",	 adVarChar,	 adParamInput, 15, U_IP)

												.Execute, , adExecuteNoRecords
										END WITH
										SET oCmd = Nothing

										oRs.MoveNext
								Loop

						END IF
						oRs.Close



						IF CouponCount > 0 THEN

								Message = ""
								Message = Message & "			<div class=""area-pop"">"
								Message = Message & "				<div class=""alert"">"
								Message = Message & "					<div class=""tit-pop"">"
								Message = Message & "						<p class=""tit"" id=""confirm_title"">SHOEMARKER</p>"
								Message = Message & "						<button id=""confirm_close"" onclick=""closePop('messagePop')"" class=""btn-hide-pop"">닫기</button>"
								Message = Message & "					</div>"
								Message = Message & "					<div class=""container-pop"">"
								Message = Message & "						<div class=""contents"">"
								Message = Message & "							<div class=""ly-cont"">"
								Message = Message & "								<p id=""confirm_content"" class=""t-level4"" style=""text-align:left"">"
	
								Message = Message &									"휴면계정 해제 처리 되었습니다.<br>"
								Message = Message &									DB_Name & "님께<br>"
																					IF CouponCount = 1 THEN
								Message = Message &											"[" & CouponName & "] 쿠폰이 발급 되었습니다."
																					ELSE
								Message = Message &											"[" & CouponName & "] 외 " & CouponCount - 1 & "장이 발급 되었습니다."
																					END IF

								Message = Message & "								</p>"
								Message = Message & "							</div>"
								Message = Message & "						</div>"
								Message = Message & "						<div class=""btns"">"
								IF OrderFlag <> "Y" THEN
								Message = Message & "							<button type=""button"" id=""message_btn1"" onclick=""APP_TopGoUrl('/ASP/Mypage/CouponList.asp');"" class=""button ty-black"">쿠폰 확인</button>"
								END IF
								Message = Message & "							<button type=""button"" id=""message_btn2"" onclick=""APP_TopGoUrl('" & ProgID & "');"" class=""button ty-red"">확인</button>"
								Message = Message & "						</div>"
								Message = Message & "					</div>"
								Message = Message & "				</div>"
								Message = Message & "			</div>"
	
								IF CouponCount = 1 THEN
										Message = DB_Name & "///" & "[" & CouponName & "] 쿠폰이 발급 되었습니다." & "///" & ProgID & "///" & OrderFlag
								ELSE
										Message = DB_Name & "///" & "[" & CouponName & "] 외 " & CouponCount - 1 & "장이 발급 되었습니다." & "///" & ProgID & "///" & OrderFlag
								END IF
						END IF



						'# 로그인 정보 입력
						SET oCmd = Server.CreateObject("ADODB.Command")
						WITH oCmd
								.ActiveConnection	 = oConn
								.CommandType		 = adCmdStoredProc
								.CommandText		 = "USP_Front_EShop_Member_Login_Insert"
	
								.Parameters.Append .CreateParameter("@MemberNum",	 adInteger,	 adParamInput,   , MemberNum)
								.Parameters.Append .CreateParameter("@Location",	 adChar,	 adParamInput,  1, "P")
								.Parameters.Append .CreateParameter("@LoginIP",		 adVarChar,	 adParamInput, 15, U_IP)

								.Execute, , adExecuteNoRecords
						END WITH
						SET oCmd = Nothing



	
						SET oRs = Nothing : oConn.Close : SET oConn = Nothing


						Response.Cookies("UIP")				 = Request.Cookies("TEMP_UIP")
						Response.Cookies("UMFLAG")			 = Request.Cookies("TEMP_MFLAG")
						Response.Cookies("UNUM")			 = Request.Cookies("TEMP_UNUM")
						Response.Cookies("UID")				 = Request.Cookies("TEMP_UID")
						Response.Cookies("UNAME")			 = Request.Cookies("TEMP_UNAME")
						Response.Cookies("UEFLAG")			 = Request.Cookies("TEMP_EFLAG")
						Response.Cookies("UETYPE")			 = Request.Cookies("TEMP_ETYPE")
						Response.Cookies("UGROUP")			 = Request.Cookies("TEMP_UGROUP")


						'# 아이디 저장
						IF SaveID = "Y" THEN
								Response.Cookies("SMEM_ID")			 = Request.Cookies("TEMP_UID")
								Response.Cookies("SMEM_ID").Expires	 = Date() + 1000
						END IF

	
						Response.Cookies("TEMP_DOR")		 = ""
						Response.Cookies("TEMP_NEW")		 = ""
						Response.Cookies("TEMP_MFLAG")		 = ""
						Response.Cookies("TEMP_PROGID")		 = ""
						Response.Cookies("TEMP_UIDSAVE")	 = ""
						Response.Cookies("TEMP_UIP")		 = ""
						Response.Cookies("TEMP_UNUM")		 = ""
						Response.Cookies("TEMP_UID")		 = ""
						Response.Cookies("TEMP_UNAME")		 = ""
						Response.Cookies("TEMP_EFLAG")		 = ""
						Response.Cookies("TEMP_ETYPE")		 = ""
						Response.Cookies("TEMP_UGROUP")		 = ""


	
						IF Message <> "" THEN
%>
								<form name="form">
									<textarea name="message" style="width:0;height:0;"><%=Message%></textarea>
								</form>
								<script type="text/javascript">
									APP_PopupHistoryBack_DOR_Result("OK2", "<%=Message%>");
								</script>
<%
						ELSE
%>
								<script type="text/javascript">
									APP_PopupHistoryBack_DOR_Result("OK", "휴면계정 해제 처리 되었습니다.");
								</script>
<%
						END IF
						Response.End

				ELSE
	
						SET oRs = Nothing : oConn.Close : SET oConn = Nothing
%>
						<script type="text/javascript">
							APP_PopupHistoryBack_DOR_Result("NEWAGREEMENT", "휴면계정 해제 처리 되었습니다.<br />신규 약관에 동의하여 주십시오.");
						</script>
<%
						Response.End

				END IF


		'# SNS 간편로그인 정회원 전환 (아이디 중복체크는 입력폼 페이지에서 처리)
		ELSEIF SMode = "JoinChgMem" THEN


				Response.Cookies("JoinType")			 = ""
				Response.Cookies("AuthType")			 = ""
				Response.Cookies("SDupInfo")			 = ""
				Response.Cookies("Name")				 = ""
				Response.Cookies("Birth")				 = ""
				Response.Cookies("Gender")				 = ""
				Response.Cookies("NationalInfo")		 = ""
				Response.Cookies("Mobile")				 = ""
				Response.Cookies("ParentSDupInfo")		 = ""
				Response.Cookies("ParentName")			 = ""
				Response.Cookies("ParentBirth")			 = ""
				Response.Cookies("ParentGender")		 = ""
				Response.Cookies("ParentNationalInfo")	 = ""
				Response.Cookies("ParentMobile")		 = ""



				Age = CDbl(U_DATE) - CDbl(sBirthDate)

				'# 14세 이상
				IF JoinType = "U" THEN

						IF Age < 140000 THEN
%>
								<script type="text/javascript">
									APP_PopupHistoryBack_JoinAuth_Result("UNDER14", "<%=sName%>님은 만 14세 미만입니다.<br />만 14세 미만 회원가입으로 보호자 동의를 받아 주십시오.");
								</script>
<%
								Response.End
						END IF


						SET oConn		 = ConnectionOpen()							'# 커넥션 생성
						SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
	

						'# 회원정보
						SET oCmd = Server.CreateObject("ADODB.Command")
						WITH oCmd
								.ActiveConnection	 = oConn
								.CommandType		 = adCmdStoredProc
								.CommandText		 = "USP_Front_EShop_Member_Select_By_SDupInfo_Check"

								.Parameters.Append .CreateParameter("@SDupInfo",	 adVarChar,	 adParamInput, 64, sDupInfo)
						END WITH
						oRs.CursorLocation = adUseClient
						oRs.Open oCmd, , adOpenStatic, adLockReadOnly
						SET oCmd = Nothing



						IF NOT oRs.EOF THEN

								oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
								Response.Cookies("JoinType")		 = Encrypt(JoinType)
								Response.Cookies("SDupInfo")		 = Encrypt(sDupInfo)
%>
								<script type="text/javascript">
									APP_PopupHistoryBack_JoinAuth_Result("MEMBER", "<%=sName%>님은 이미 회원으로 가입되어 있습니다.<br />정회원 정보로 통합하시겠습니까?");
								</script>
<%
								Response.End

						ELSE
	
								Response.Cookies("JOIN_TYPE")		 = ""
								Response.Cookies("JoinType")		 = Encrypt(JoinType)
								Response.Cookies("AuthType")		 = Encrypt("I")
								Response.Cookies("SDupInfo")		 = Encrypt(sDupInfo)
								Response.Cookies("Name")			 = Encrypt(sName)
								Response.Cookies("Birth")			 = Encrypt(sBirthDate)
								Response.Cookies("Gender")			 = Encrypt(sGenderCode)
								'# Response.Cookies("NationalInfo") = Encrypt(sNationalInfo)
								Response.Cookies("Mobile")			 = Encrypt("")

						END IF
						oRs.Close
	
						SET oRs = Nothing : oConn.Close : SET oConn = Nothing
%>
						<script type="text/javascript">
							APP_PopupHistoryBack_JoinAuth_Result("OK", "");
						</script>
<%
						Response.End


				'# 만14세 미만
				ELSE

						IF Age < 200000 THEN
%>
								<script type="text/javascript">
									APP_PopupHistoryBack_JoinAuth_Result("UNDER20", "만 20세 이상만 보호자 인증을 받을 수 있습니다.");
								</script>
<%
								Response.End
						END IF
	
'만14세 미만 정회원전환은 부보 DI값 중복 체크 안함
'#						SET oConn		 = ConnectionOpen()							'# 커넥션 생성
'#						SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
'#	
'#
'#						'# 회원정보(보호자 DI값 확인)
'#						SET oCmd = Server.CreateObject("ADODB.Command")
'#						WITH oCmd
'#								.ActiveConnection	 = oConn
'#								.CommandType		 = adCmdStoredProc
'#								.CommandText		 = "USP_Front_EShop_Member_Select_By_ParentSDupInfo_Check"
'#
'#								.Parameters.Append .CreateParameter("@SDupInfo",	 adVarChar,	 adParamInput, 64, sDupInfo)
'#						END WITH
'#						oRs.CursorLocation = adUseClient
'#						oRs.Open oCmd, , adOpenStatic, adLockReadOnly
'#						SET oCmd = Nothing
'#
'#
'#
'#						IF NOT oRs.EOF THEN
'#
'#								oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
'#								Response.Cookies("JoinType")		 = Encrypt(JoinType)
'#								Response.Cookies("ParentSDupInfo")	 = Encrypt(sDupInfo)
%>
								<script type="text/javascript">
									//APP_PopupHistoryBack_JoinAuth_Result("MEMBER", "<%=sName%>님은 이미 회원으로 가입되어 있습니다.<br />정회원 정보로 통합하시겠습니까?");
								</script>
<%
'#								Response.End
'#
'#						ELSE

								Response.Cookies("JOIN_TYPE")			 = ""
								Response.Cookies("JoinType")			 = Encrypt(JoinType)
								Response.Cookies("AuthType")			 = Encrypt("I")
								Response.Cookies("ParentSDupInfo")		 = Encrypt(sDupInfo)
								Response.Cookies("ParentName")			 = Encrypt(sName)
								Response.Cookies("ParentBirth")			 = Encrypt(sBirthDate)
								Response.Cookies("ParentGender")		 = Encrypt(sGenderCode)
								'# Response.Cookies("ParentNationalInfo")	 = Encrypt(sNationalInfo)
								Response.Cookies("ParentMobile")		 = Encrypt("")
'#						END IF
'#						oRs.Close
'#	
'#						SET oRs = Nothing : oConn.Close : SET oConn = Nothing
%>
						<script type="text/javascript">
							APP_PopupHistoryBack_JoinAuth_Result("OK", "");
						</script>
<%
						Response.End

				END IF


		END IF



ELSEIF (iRtn = -9) THEN
		sErrCode = "입력값 오류 : fnResponse 또는 fnResponseExt 함수 처리시, 필요한 파라미터값의 정보를 정확하게 입력해 주시기 바랍니다."

		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=본인인증에 실패하였습니다.<br />본인인증 실패 사유 : 요청정보 암호화 오류[-9]&Script=APP_PopupHistoryBack();"
		Response.End
ELSEIF (iRtn = -12) THEN
		sErrCode = "CP 비밀번호 불일치 : IPIN 서비스 사이트 패스워드를 확인해 주시기 바랍니다."

		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=CP 비밀번호 불일치 : <br />IPIN 서비스 사이트 패스워드를 확인해 주시기 바랍니다.&Script=APP_PopupHistoryBack();"
		Response.End
ELSEIF (iRtn = -13) THEN
		sErrCode = "CP 요청번호 불일치 : 세션에 넣은 sCPRequest 데이타를 확인해 주시기 바랍니다."
	
		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=CP 비밀번호 불일치 : <br />세션에 넣은 sCPRequest 데이타를 확인해 주시기 바랍니다.&Script=APP_PopupHistoryBack();"
		Response.End
ELSE
		sErrCode = "iRtn 값 확인 후, 한국신용평가정보 개발 담당자에게 문의해 주세요."
	
		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=iRtn 값 확인 후, 한국신용평가정보 개발 담당자에게 문의해 주세요.&Script=APP_PopupHistoryBack();"
		Response.End
END IF
	
SET clsIPINDll = NOTHING
%>