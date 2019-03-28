<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'LoginOk.asp - 로그인 처리
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
PageCode3 = "02"
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/md5.asp" -->

<%
'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn						'# ADODB Connection 개체
DIM oRs							'# ADODB Recordset 개체
DIM oCmd						'# ADODB Command 개체

DIM j
DIM x
DIM y

DIM wQuery						'# WHERE 절
DIM sQuery						'# SORT 절


DIM ProgID
DIM OrderFlag
DIM UserID
DIM Pwd
DIM PwdMd5
DIM PHPCrypt
DIM PwdEnc
DIM SaveID

DIM SavedCookieID				'# 저장 아이디


DIM DB_MemberNum
DIM DB_GroupCode
DIM DB_Name
DIM DB_EmployeeFlag
DIM DB_EmployeeType
DIM DB_DelFlag
DIM DB_DormancyFlag
DIM DB_Pwd
DIM DB_OldPwd
DIM DB_MemberFlag
DIM DB_GroupName
DIM DB_Birth
DIM DB_Sex

Dim Age
Dim Gender

DIM OldPwdFlag : OldPwdFlag = "N"
DIM NewAgreementFlag : NewAgreementFlag = "N"

DIM CouponName
DIM CouponCount : CouponCount = 0

DIM StartDT
DIM EndDT

DIM Message
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


'# 주문작성 로그인 페이지에서 사용되는 쿠키
Response.Cookies("LN_ORDER")		 = ""
Response.Cookies("LN_PROGID")		 = ""


ProgID			 = Request("ProgID")
IF ProgID		 = "" THEN ProgID = "/"
OrderFlag		 = sqlFilter(Request("OrderFlag"))
UserID			 = sqlFilter(Request("UserID"))
Pwd				 = sqlFilter(Request("Pwd"))
SaveID			 = sqlFilter(Request("saveid"))


IF UserID = "" THEN
		Response.Write "FAIL|||||아이디를 입력하여 주십시오."
		Response.End
END IF
IF Pwd = "" THEN
		Response.Write "FAIL|||||비밀번호를 입력하여 주십시오."
		Response.End
END IF


'# MD5 비밀번호 암호화
PwdMd5			 = md5(LCase(Pwd))

'# PHP Crypt 비밀번호 암호화
SET PHPCrypt = Server.CreateObject("PHP.Crypt")
PwdEnc		 = PHPCrypt.Crypt("35e80f121fcae9fdb4d9a4d342e04f76", Pwd)
SET PHPCrypt = nothing



SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성



'# 회원 정보 - 아이디 찾기
SET oCmd = SErver.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Member_Select_By_UserID"

		.Parameters.Append .CreateParameter("@UserID", adVarChar, adParamInput, 30, UserID)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		DB_MemberNum		 = oRs("MemberNum")
		DB_GroupCode		 = oRs("GroupCode")
		DB_Name				 = oRs("Name")
		DB_EmployeeFlag		 = oRs("EmployeeFlag")
		DB_EmployeeType		 = oRs("EmployeeType")
		DB_DelFlag			 = oRs("DelFlag")
		DB_DormancyFlag		 = oRs("DormancyFlag")
		DB_Pwd				 = oRs("Pwd")
		DB_OldPwd			 = oRs("OldPwd")
		DB_MemberFlag		 = oRs("MemberFlag")
		DB_GroupName		 = oRs("GroupName")
		DB_Birth			 = oRs("Birth")
		DB_Sex				 = oRs("Sex")
ELSE
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "ID|||||"
		Response.End
END IF
oRs.Close

'# 탈퇴회원
IF DB_DelFlag = "Y" THEN
		SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "ID|||||"
		Response.End
END IF
'# SNS 전환여부
IF DB_MemberFlag = "N" AND Request.Cookies("SNS_UID")="" THEN
		SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "ID|||||"
		Response.End
END IF

IF TRIM(DB_Pwd) <> "" THEN
		IF DB_Pwd <> PwdMd5 THEN
				SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "PWD|||||"
				Response.End
		END IF
ELSE
		IF DB_OldPwd <> PwdEnc THEN
				SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "PWD|||||"
				Response.End
		ELSE
				OldPwdFlag = "Y"
		END IF
END IF

'# 비밀번호가 md5가 아닌 회원 정보 업데이트
IF OldPwdFlag = "Y" THEN

		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Member_Update_For_Md5_Pwd"
	
				.Parameters.Append .CreateParameter("@MemberNum",	 adInteger, adParamInput,   , DB_MemberNum)
				.Parameters.Append .CreateParameter("@Pwd",			 adVarChar, adParamInput, 50, PwdMd5)

				.Execute, , adExecuteNoRecords
		END WITH
		SET oCmd = Nothing
		
END IF



'# 신규 약관 동의 여부
SET oCmd = SErver.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Member_NewAgreement_Select_By_MemberNum"
	
		.Parameters.Append .CreateParameter("@MemberNum",	 adInteger, adParamInput,   , DB_MemberNum)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		NewAgreementFlag = "Y"
ELSE
		NewAgreementFlag = "N"
END IF
oRs.Close



'# SNS 회원 로그인 연결처리(휴먼회원 아닐때 연결처리)
IF Request.Cookies("SNS_UID")<>"" AND Request.Cookies("SNS_Kind")<>"" AND DB_DormancyFlag <> "Y" THEN
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Member_SNS_Insert"

				.Parameters.Append .CreateParameter("@MemberNum",	adInteger,	 adParamInput,    , DB_MemberNum)
				.Parameters.Append .CreateParameter("@SNSKind",		adVarChar,	 adParamInput,  30, Decrypt(Request.Cookies("SNS_Kind")))
				.Parameters.Append .CreateParameter("@SnsID",		adVarChar,	 adParamInput,  50, Decrypt(Request.Cookies("SNS_UID")))
				.Parameters.Append .CreateParameter("@Email",		adVarChar,	 adParamInput,  50, Decrypt(Request.Cookies("SNS_Email")))
				.Parameters.Append .CreateParameter("@CreateIP",	adVarChar,	 adParamInput,  50, U_IP)

				.Execute, , adExecuteNoRecords
		END WITH
		SET oCmd = Nothing
		'# SNS 회원정보 초기화
		Response.Cookies("SNS_UID")		= ""
		Response.Cookies("SNS_Kind")		= ""
		Response.Cookies("SNS_Email")	= ""
		Response.Cookies("SNS_KName")	= ""
		Response.Cookies("SNS_UserID")	= ""
		Response.Cookies("SNS_UNUM")		= ""
END IF

	

	
'# 휴면계정 OR 신규약관 미동의 일 경우
IF DB_DormancyFlag = "Y" OR NewAgreementFlag = "N" THEN

		SET oRs = Nothing : oConn.Close : SET oConn = Nothing
	
		Response.Cookies("TEMP_DOR")		 = Encrypt(DB_DormancyFlag)
		Response.Cookies("TEMP_NEW")		 = Encrypt(NewAgreementFlag)
		Response.Cookies("TEMP_MFLAG")		 = Encrypt(DB_MemberFlag)
		Response.Cookies("TEMP_PROGID")		 = Encrypt(ProgID)
		Response.Cookies("TEMP_UIDSAVE")	 = Encrypt(SaveID)
		Response.Cookies("TEMP_UIP")		 = Encrypt(U_IP)
		Response.Cookies("TEMP_UNUM")		 = Encrypt(DB_MemberNum)
		Response.Cookies("TEMP_UID")		 = Encrypt(UserID)
		Response.Cookies("TEMP_UNAME")		 = Encrypt(DB_Name)
		Response.Cookies("TEMP_EFLAG")		 = Encrypt(DB_EmployeeFlag)
		Response.Cookies("TEMP_ETYPE")		 = Encrypt(DB_EmployeeType)
		Response.Cookies("TEMP_UGROUP")		 = Encrypt(DB_GroupCode)
		Response.Cookies("TEMP_ORDERFLAG")	 = Encrypt(OrderFlag)

		IF DB_DormancyFlag = "Y" THEN
				Response.Write "DORMANCY|||||"
		ELSEIF DB_DormancyFlag = "N" AND NewAgreementFlag = "N" THEN
				Response.Write "NEWAGREE|||||"
		END IF
		Response.End

ELSE

		'# 배포된 쿠폰 발급 받기 - EShop_Coupon_Member 에 배포되었지만 아직 안받은 쿠폰
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Coupon_Member_Select_For_Not_Receive"
	
				.Parameters.Append .CreateParameter("@MemberNum",	 adInteger,	 adParamInput,   , DB_MemberNum)
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
	
						.Parameters.Append .CreateParameter("@MemberNum",	 adInteger,	 adParamInput,   , DB_MemberNum)
						.Parameters.Append .CreateParameter("@UpdateID",	 adVarChar,	 adParamInput, 20, DB_MemberNum)
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
	
				.Parameters.Append .CreateParameter("@MemberNum",	 adInteger,	 adParamInput,   , DB_MemberNum)
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
	
								.Parameters.Append .CreateParameter("@MemberNum",	 adInteger,	 adParamInput,   , DB_MemberNum)
								.Parameters.Append .CreateParameter("@CouponIdx",	 adBigInt,	 adParamInput,   , oRs("Idx"))
								.Parameters.Append .CreateParameter("@StartDT",		 adVarChar,	 adParamInput, 14, StartDT)
								.Parameters.Append .CreateParameter("@EndDT",		 adVarChar,	 adParamInput, 14, EndDT)
								.Parameters.Append .CreateParameter("@CreateID",	 adVarChar,	 adParamInput, 20, DB_MemberNum)
								.Parameters.Append .CreateParameter("@CreateIP",	 adVarChar,	 adParamInput, 15, U_IP)

								.Execute, , adExecuteNoRecords
						END WITH
						SET oCmd = Nothing

						oRs.MoveNext
				Loop

		END IF
		oRs.Close







		'# 아이디 저장
		IF SaveID = "Y" THEN
				Response.Cookies("SMEM_ID")			 = Encrypt(UserID)
				Response.Cookies("SMEM_ID").Expires	 = Date() + 1000
		ELSE
				Response.Cookies("SMEM_ID")			 = ""
				Response.Cookies("SMEM_ID").Expires	 = Date()-1
		END IF


		'# 로그인 정보 입력
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Member_Login_Insert"
	
				.Parameters.Append .CreateParameter("@MemberNum",	 adInteger,	 adParamInput,   , DB_MemberNum)
				.Parameters.Append .CreateParameter("@Location",	 adChar,	 adParamInput,  1, "A")
				.Parameters.Append .CreateParameter("@LoginIP",		 adVarChar,	 adParamInput, 15, U_IP)

				.Execute, , adExecuteNoRecords
		END WITH
		SET oCmd = Nothing
	
		SET oRs = Nothing : oConn.Close : SET oConn = Nothing

		Response.Cookies("UIP")				 = Encrypt(U_IP)
		Response.Cookies("UMFLAG")			 = Encrypt(DB_MemberFlag)
		Response.Cookies("UNUM")			 = Encrypt(DB_MemberNum)
		Response.Cookies("UID")				 = Encrypt(UserID)
		Response.Cookies("UNAME")			 = Encrypt(DB_Name)
		Response.Cookies("UEFLAG")			 = Encrypt(DB_EmployeeFlag)
		Response.Cookies("UETYPE")			 = Encrypt(DB_EmployeeType)
		Response.Cookies("UGROUP")			 = Encrypt(DB_GroupCode)


		IF CouponCount > 0 THEN
	
				Message = ""
	
				Message = Message & "			<div class=""area-dim"" style=""z-index:101""></div>"&vbLf
				Message = Message & "			<div class=""area-pop"">"&vbLf
				Message = Message & "				<div class=""alert"">"&vbLf
				Message = Message & "					<div class=""tit-pop"">"&vbLf
				Message = Message & "						<p class=""tit"" id=""confirm_title"">SHOEMARKER</p>"&vbLf
				Message = Message & "						<button id=""confirm_close"" onclick=""closePop('messagePop');location.href='" & ProgID & "';"" class=""btn-hide-pop"">닫기</button>"&vbLf
				Message = Message & "					</div>"
				Message = Message & "					<div class=""container-pop"">"&vbLf
				Message = Message & "						<div class=""contents"">"&vbLf
				Message = Message & "							<div class=""ly-cont"">"&vbLf
				Message = Message & "								<p id=""confirm_content"" class=""t-level4"" style=""text-align:left"">"&vbLf

				Message = Message &									DB_Name & "님께<br>"
																	IF CouponCount = 1 THEN
				Message = Message &											"[" & CouponName & "] 쿠폰이 발급 되었습니다."
																	ELSE
				Message = Message &											"[" & CouponName & "] 외 " & CouponCount - 1 & "장이 발급 되었습니다."
																	END IF

				Message = Message & "								</p>"&vbLf
				Message = Message & "							</div>"&vbLf
				Message = Message & "						</div>"&vbLf
				Message = Message & "						<div class=""btns"">"&vbLf
				IF OrderFlag <> "Y" THEN
				Message = Message & "							<button type=""button"" id=""message_btn1"" onclick=""move_MyCouponList()"" class=""button ty-black"">쿠폰 확인</button>"
				END IF
				Message = Message & "							<button type=""button"" id=""message_btn2"" onclick=""closePop('messagePop');location.href='" & ProgID & "';"" class=""button ty-red"">확인</button>"
				Message = Message & "						</div>"&vbLf
				Message = Message & "					</div>"&vbLf
				Message = Message & "				</div>"&vbLf
				Message = Message & "			</div>"&vbLf

		END IF


		If DB_Birth = "" Then
			Age = 0
		Else
			If Len(DB_Birth) > 4 Then
				Age = Cint(Year(Date)) - Cint(Left(DB_Birth, 4))
			Else
				Age = 0
			End If
		End If

		If DB_Sex = "M" Then
			Gender = "man"
		ElseIf DB_Sex = "F" Then
			Gender = "woman"
		Else
			Gender = ""
		End If
	
		Response.Write "OK|||||" & Message & "|||||" & Age & "|||||" & Gender

		Response.End

END IF
%>