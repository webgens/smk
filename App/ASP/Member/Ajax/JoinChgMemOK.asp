<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'JoinForm.asp - 회원 가입(정회원 전환) - 내용입력
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
PageCode3 = "04"
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/md5.asp" -->

<%
IF U_ID <> "" AND U_MFLAG = "Y" THEN
		Response.Write "FAIL|||||이미 로그인 되어 있습니다."
		Response.End
END IF


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


DIM AgreeChk1
DIM AgreeChk2
DIM AgreeChk3
DIM AgreeChk4
DIM Agreement
DIM ThirdPartyFlag
DIM SMSFlag
DIM EmailFlag

DIM JoinType					'# 14세 구분 (U:14세이상 / D:14세미만)
DIM AuthType

DIM SDupInfo
DIM ParentSDupInfo

DIM UserID
DIM Pwd
DIM Name
DIM Birth
DIM ChkBirth
DIM Sex
DIM HP
DIM Email
DIM ZipCode
DIM Addr1
DIM Addr2
DIM Area

DIM ParentName
DIM ParentBirth
DIM ChkParentBirth
DIM ParentHP
DIM ParentEmail

Dim PwdMd5
Dim FTFlag
	
DIM CouponIdx
DIM UseDateType
DIM UseSDate
DIM UseEDate
DIM UseDay
DIM StartDT
DIM EndDT

DIM MemberNum
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


AgreeChk1			 = TRIM(sqlFilter(Request("AgreeChk1")))
AgreeChk2			 = TRIM(sqlFilter(Request("AgreeChk2")))
AgreeChk3			 = TRIM(sqlFilter(Request("AgreeChk3")))
AgreeChk4			 = TRIM(sqlFilter(Request("AgreeChk4")))
ThirdPartyFlag		 = TRIM(sqlFilter(Request("ThirdPartyFlag")))
SMSFlag				 = TRIM(sqlFilter(Request("SMSFlag")))
EmailFlag			 = TRIM(sqlFilter(Request("EMailFlag")))


JoinType			 = TRIM(sqlFilter(Request("JoinType")))
AuthType			 = TRIM(Decrypt(Request.Cookies("AuthType")))
	
SDupInfo			 = TRIM(Decrypt(Request.Cookies("SDupInfo")))
ParentSDupInfo		 = TRIM(Decrypt(Request.Cookies("ParentSDupInfo")))


Userid				 = LCASE(TRIM(sqlFilter(Request("Userid"))))
Pwd					 = LCASE(TRIM(sqlFilter(Request("Pwd"))))
Name				 = TRIM(sqlFilter(Request("Name")))
Birth				 = TRIM(sqlFilter(Request("Birth")))
Sex					 = TRIM(sqlFilter(Request("Sex")))
HP					 = TRIM(sqlFilter(Request("HP1"))) & "-" & TRIM(sqlFilter(Request("HP2"))) & "-" & TRIM(sqlFilter(Request("HP3")))
Email				 = TRIM(sqlFilter(Request("Email")))
ZipCode				 = TRIM(sqlFilter(Request("ZipCode")))
Addr1				 = TRIM(sqlFilter(Request("Addr1")))
Addr2				 = TRIM(sqlFilter(Request("Addr2")))
Area				 = LEFT(Addr1, 2)

ParentName			 = TRIM(sqlFilter(Request("ParentName")))
ParentBirth			 = TRIM(sqlFilter(Request("ParentBirth")))
ParentHP			 = TRIM(sqlFilter(Request("PHP1"))) & "-" & TRIM(sqlFilter(Request("PHP2"))) & "-" & TRIM(sqlFilter(Request("PHP3")))
ParentEmail			 = TRIM(sqlFilter(Request("ParentEmail")))


FTFlag				 = "N"
IF JoinType			 = "D"	 THEN FTFlag		 = "Y"
IF ThirdPartyFlag	 = ""	 THEN ThirdPartyFlag = "N"
IF SMSFlag			 = ""	 THEN SMSFlag		 = "N"
IF EMailFlag		 = ""	 THEN EmailFlag		 = "N"

IF AgreeChk1		 = ""	 THEN AgreeChk1		 = "N"
IF AgreeChk2		 = ""	 THEN AgreeChk2		 = "N"
IF AgreeChk3		 = ""	 THEN AgreeChk3		 = "N"
IF AgreeChk4		 = ""	 THEN AgreeChk4		 = "N"
Agreement			 = AgreeChk1 & "|" & AgreeChk2 & "|" & AgreeChk3 & "|" & AgreeChk4 & "|" & ThirdPartyFlag & "|" & SMSFlag & "|" & EMailFlag


IF JoinType = "" THEN
		Response.Write "FAIL|||||만14세 구분 정보가 없습니다."
		Response.End
END IF
	
IF AgreeChk1 = "N"  THEN
		Response.Write "FAIL|||||쇼핑몰 이용약관에 동의 하셔야 됩니다."
		Response.End
END IF

IF AgreeChk2 = "N"  THEN
		Response.Write "FAIL|||||개인정보 이용 및 수집에 대해 동의 하셔야 됩니다."
		Response.End
END IF

IF JoinType = "U" AND SDupInfo = "" THEN
		Response.Write "FAIL|||||본인인증 정보가 없습니다."
		Response.End
END IF

IF JoinType = "D" AND ParentsDupInfo = "" THEN
		Response.Write "FAIL|||||보호자 본인인증 정보가 없습니다."
		Response.End
END IF

IF TRIM(UserID) = "" THEN
		Response.Write "FAIL|||||아이디 정보가 없습니다."
		Response.End
END IF

IF TRIM(Pwd) = "" THEN
		Response.Write "FAIL|||||비밀번호 정보가 없습니다."
		Response.End
END IF

IF TRIM(Pwd) = TRIM(UserID) THEN
		Response.Write "FAIL|||||아이디와 동일한 비밀번호를 사용하실 수 없습니다."
		Response.End
END IF

IF chk_SameChr(Pwd, 4) = False THEN
		Response.Write "FAIL|||||4자리 이상의 동일한 문자는 사용이 불가합니다."
		Response.End
END IF


IF TRIM(Name) = "" THEN
		Response.Write "FAIL|||||이름 정보가 없습니다."
		Response.End
END IF

IF TRIM(Birth) = "" THEN
		Response.Write "FAIL|||||생년월일 정보가 없습니다."
		Response.End
END IF

IF LEN(TRIM(Birth)) <> 8 OR ISNUMERIC(TRIM(Birth)) = False THEN
		Response.Write "FAIL|||||생년월일 정보가 유효하지 않습니다."
		Response.End
END IF

ChkBirth = LEFT(Birth, 4) & "-" & MID(Birth, 5, 2) & "-" & MID(Birth, 7, 2)
IF ISDATE(ChkBirth) = False THEN
		Response.Write "FAIL|||||생년월일 정보가 유효하지 않습니다."
		Response.End
END IF

IF TRIM(Sex) = "" THEN
		Response.Write "FAIL|||||성별 정보가 없습니다."
		Response.End
END IF

IF TRIM(HP) = "" OR TRIM(HP) = "--" THEN
		Response.Write "FAIL|||||휴대폰 정보가 없습니다."
		Response.End
END IF

IF TRIM(Email) = "" OR TRIM(Email) = "@" THEN
		Response.Write "FAIL|||||이메일 정보가 없습니다."
		Response.End
END IF


IF JoinType = "D" THEN
		IF TRIM(ParentName) = "" THEN
				Response.Write "FAIL|||||보호자 이름 정보가 없습니다."
				Response.End
		END IF

		IF TRIM(ParentBirth) = "" THEN
				Response.Write "FAIL|||||보호자 생년월일 정보가 없습니다."
				Response.End
		END IF

		IF LEN(TRIM(ParentBirth)) <> 8 OR ISNUMERIC(TRIM(ParentBirth)) = False THEN
				Response.Write "FAIL|||||보호자 생년월일 정보가 유효하지 않습니다."
				Response.End
		END IF

		ChkParentBirth = LEFT(ParentBirth, 4) & "-" & MID(ParentBirth, 5, 2) & "-" & MID(ParentBirth, 7, 2)
		IF ISDATE(ChkParentBirth) = False THEN
				Response.Write "FAIL|||||보호자 생년월일 정보가 유효하지 않습니다."
				Response.End
		END IF

		IF TRIM(ParentHP) = "" OR TRIM(ParentHP) = "--" THEN
				Response.Write "FAIL|||||보호자 휴대폰 정보가 없습니다."
				Response.End
		END IF

		IF TRIM(ParentEmail) = "" OR TRIM(ParentEmail) = "@" THEN
				Response.Write "FAIL|||||보호자 이메일 정보가 없습니다."
				Response.End
		END IF
ELSE
		ParentEmail = ""
END IF
		


'# MD5 비밀번호 암호화
PwdMd5				 = md5(LCase(Pwd))








SET oConn			 = ConnectionOpen()							'# 커넥션 생성
SET oRs				 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성



	


'# DI값 체크 - 만14세 이상 일 경우만
IF JoinType = "U" THEN
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Member_Select_By_SDupInfo_Check"

				.Parameters.Append .CreateParameter("@SDupInfo", adVarChar, adParamInput, 64, SDupInfo)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		SET oCmd = Nothing

		IF NOT oRs.EOF THEN
				oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "DIDUP|||||회원으로 가입 되어 계십니다. 아이디/비밀번호 찾기를 이용해 주십시오."
				Response.End
		END IF
		oRs.Close
END IF




'# 아이디 중복 체크
SET oCmd = Server.CreateObject("ADODB.Command")
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
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
	
		Response.Write "IDDUP|||||" & UserID & "는 사용할 수 없는 아이디 입니다."
		Response.End
END IF
oRs.Close




	


oConn.BeginTrans


	



'# 정회원가입 수정
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Member_Update_For_Member"
	
		.Parameters.Append .CreateParameter("@MemberNum",		 adInteger,	 adParamInput,    , U_NUM)
		.Parameters.Append .CreateParameter("@UserID",			 adVarChar,	 adParamInput,  30, UserID)
		.Parameters.Append .CreateParameter("@Pwd",				 adVarChar,	 adParamInput,  50, PwdMd5)
		.Parameters.Append .CreateParameter("@Name",			 adVarChar,	 adParamInput,  50, Name)
		.Parameters.Append .CreateParameter("@Birth",			 adVarChar,	 adParamInput,   8, Birth)
		.Parameters.Append .CreateParameter("@Sex",				 adChar,	 adParamInput,   1, Sex)
		.Parameters.Append .CreateParameter("@ZipCode",			 adVarChar,	 adParamInput,   6, ZipCode)
		.Parameters.Append .CreateParameter("@Address1",		 adVarChar,	 adParamInput, 400, Addr1)
		.Parameters.Append .CreateParameter("@Address2",		 adVarChar,	 adParamInput, 400, Addr2)
		.Parameters.Append .CreateParameter("@Area",			 adVarChar,	 adParamInput,  10, Area)
		.Parameters.Append .CreateParameter("@HP",				 adVarChar,	 adParamInput,  14, HP)
		.Parameters.Append .CreateParameter("@Email",			 adVarChar,	 adParamInput,  50, Email)
		.Parameters.Append .CreateParameter("@SmsFlag",			 adChar,	 adParamInput,   1, SMSFlag)
		.Parameters.Append .CreateParameter("@EmailFlag",		 adChar,	 adParamInput,   1, EmailFlag)
		.Parameters.Append .CreateParameter("@AuthType",		 adChar,	 adParamInput,   1, AuthType)
		.Parameters.Append .CreateParameter("@FTFlag",			 adChar,	 adParamInput,   1, FTFlag)
		.Parameters.Append .CreateParameter("@sDupInfo",		 adVarChar,	 adParamInput,  64, SDupInfo)
		.Parameters.Append .CreateParameter("@ParentSDupInfo",	 adVarChar,	 adParamInput,  64, ParentSDupInfo)
		.Parameters.Append .CreateParameter("@ParentName",		 adVarChar,	 adParamInput,  20, ParentName)
		.Parameters.Append .CreateParameter("@ParentBirth",		 adVarChar,	 adParamInput,   8, ParentBirth)
		.Parameters.Append .CreateParameter("@ParentHP",		 adVarChar,	 adParamInput,  14, ParentHP)
		.Parameters.Append .CreateParameter("@ParentEmail",		 adVarChar,	 adParamInput,  50, ParentEmail)
		.Parameters.Append .CreateParameter("@UpdateID",		 adVarChar,	 adParamInput,  20, U_NUM)
		.Parameters.Append .CreateParameter("@UpdateIP",		 adVarChar,	 adParamInput,  15, U_IP)

		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing



'# 회원가입 축하 쿠폰
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Coupon_Select_For_PC_Member_Join"
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		CouponIdx	 = oRs("Idx")
		UseDateType	 = oRs("UseDateType")
		UseSDate	 = oRs("UseSDate")
		UseEDate	 = oRs("UseEDate")
		UseDay		 = oRs("UseDay")

		'# 쿠폰 사용기간이 무제한일 경우
		IF UseDateType = "U" THEN
				StartDT	 = U_DATE & "000000"
				EndDT	 = "99999999999999"
		'# 쿠폰 사용기간이 기간으로 정해져 있을 경우
		ELSEIF UseDateType = "P" THEN
				StartDT	 = UseSDate
				EndDT	 = UseEDate
		'# 쿠폰 사용기간이 일자만큼 사용할 수 있을 경우
		ELSEIF UseDateType = "D" THEN
				StartDT	 = U_DATE & "000000"
				EndDT	 = REPLACE(DateAdd("d", UseDay, Date), "-", "") & "240000"
		END IF


		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Coupon_Member_Insert_For_Member_Join"

				.Parameters.Append .CreateParameter("@MemberNum",		 adInteger,	 adParamInput,   , U_NUM)
				.Parameters.Append .CreateParameter("@CouponIdx",		 adInteger,	 adParamInput,   , CouponIdx)
				.Parameters.Append .CreateParameter("@StartDT",			 adVarChar,	 adParamInput, 14, StartDT)
				.Parameters.Append .CreateParameter("@EndDT",			 adVarChar,	 adParamInput, 14, EndDT)
				.Parameters.Append .CreateParameter("@CreateID",		 adVarChar,	 adParamInput, 20, U_NUM)
				.Parameters.Append .CreateParameter("@CreateIP",		 adVarChar,	 adParamInput, 15, U_IP)

				.Execute, , adExecuteNoRecords
		END WITH
		SET oCmd = Nothing
END IF
oRs.Close




'# 신규 약관 동의 입력
SET oCmd = SErver.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Member_NewAgreement_Insert"
	
		.Parameters.Append .CreateParameter("@MemberNum",	 adInteger, adParamInput,   , U_NUM)
		.Parameters.Append .CreateParameter("@Agreement",	 adVarChar, adParamInput, 20, Agreement)
		.Parameters.Append .CreateParameter("@CreateIP",	 adVarChar,	adParamInput, 15, U_IP)

		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing








'# 로그아웃 처리
Response.Cookies("UIP").Expires			 = Now - 1000
Response.Cookies("UMFLAG").Expires		 = Now - 1000
Response.Cookies("UNUM").Expires		 = Now - 1000
Response.Cookies("UID").Expires			 = Now - 1000
Response.Cookies("UNAME").Expires		 = Now - 1000
Response.Cookies("UMFLAG").Expires		 = Now - 1000
Response.Cookies("UETYPE").Expires		 = Now - 1000
Response.Cookies("UGROUP").Expires		 = Now - 1000








IF Err.number <> "0" THEN
		oConn.RollbackTrans
		SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||처리 중 오류가 발생하였습니다."
		Response.End
ELSE
		oConn.CommitTrans
END IF





SET oRs = Nothing
oConn.Close
SET oConn = Nothing

Response.Write "OK|||||"
%>