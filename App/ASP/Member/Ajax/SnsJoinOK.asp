<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'SnsJoinForm.asp - SNS 간편로그인 회원 가입 - 내용입력
'Date		: 2018.12.10
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
PageCode3 = "03"
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/md5.asp" -->

<%
IF U_ID <> "" THEN
		Response.Write "LOGINED|||||이미 로그인 되어 있습니다."
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
DIM snsEmail

DIM SNS_UID
DIM SNS_Email
DIM SNS_KName
DIM SNS_Kind
DIM SNS_ProgID

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
snsEmail			 = TRIM(sqlFilter(Request("snsEmail")))

SNS_UID				 = Decrypt(Request.Cookies("SNS_UID"))
SNS_Email			 = Decrypt(Request.Cookies("SNS_Email"))
SNS_Kind			 = Decrypt(Request.Cookies("SNS_Kind"))
SNS_KName			 = TRIM(Decrypt(Request.Cookies("SNS_KName")))
SNS_ProgID			 = Decrypt(Request.Cookies("SNS_ProgID"))


IF SNS_KName		 = "" THEN SNS_KName = "SNS회원("& SNS_Kind &")"
IF SNS_ProgID		 = "" THEN SNS_ProgID		 = "/"
IF ThirdPartyFlag	 = "" THEN ThirdPartyFlag = "N"
IF SMSFlag			 = "" THEN SMSFlag	 = "N"
IF EMailFlag		 = "" THEN EmailFlag	 = "N"

IF AgreeChk1		 = "" THEN AgreeChk1 = "N"
IF AgreeChk2		 = "" THEN AgreeChk2 = "N"
IF AgreeChk3		 = "" THEN AgreeChk3 = "N"
IF AgreeChk4		 = "" THEN AgreeChk4 = "N"

Agreement			 = AgreeChk1 & "|" & AgreeChk2 & "|" & AgreeChk3 & "|" & AgreeChk4 & "|" & ThirdPartyFlag & "|" & SMSFlag & "|" & EMailFlag


	
'# IF AgreeChk1 = "" OR AgreeChk2 = "" OR AgreeChk3 = "" OR AgreeChk4 = "" THEN
IF AgreeChk2 = "" THEN
		Response.Write "FAIL|||||개인정보 이용 및 수집에 대한 동의 정보가 없습니다."
		Response.End
END IF

IF TRIM(SNS_UID) = "" THEN
		Response.Write "FAIL|||||SNS아이디 정보가 없습니다."
		Response.End
END IF

IF TRIM(SNS_KName) = "" THEN
		Response.Write "FAIL|||||SNS성명 정보가 없습니다."
		Response.End
END IF

IF TRIM(SNS_Kind) = "" THEN
		Response.Write "FAIL|||||SNS구분 정보가 없습니다."
		Response.End
END IF

IF TRIM(snsEmail) = "" THEN
		Response.Write "FAIL|||||이메일 정보가 없습니다."
		Response.End
END IF
		



SET oConn			 = ConnectionOpen()							'# 커넥션 생성
SET oRs				 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성



'# SNS 중복체크
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Member_SNS_Select_By_SDupInfo_Check"

		.Parameters.Append .CreateParameter("@SNSKind", adChar,		adParamInput,	 1, SNS_Kind)
		.Parameters.Append .CreateParameter("@SNSID"  , adVarChar,	adParamInput,	50, SNS_UID)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		IF oRs("U_NUM") <> "" AND oRs("MemberFlag") = "Y" THEN
				Response.Write "DIDUP|||||이미 정회원으로 가입된 정보가 있습니다. 아이디/비밀번호 찾기를 이용해 주십시오."
		ELSEIF oRs("U_NUM") <> "" AND oRs("MemberFlag") = "N" THEN
				Response.Write "DIDUP|||||이미 SNS회원으로 등록된 정보가 있습니다."
		END IF
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.End
END IF
oRs.Close




'# 아이디 중복 체크
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Member_Select_By_UserID"

		.Parameters.Append .CreateParameter("@UserID", adVarChar, adParamInput, 30, SNS_UID)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
	
		Response.Write "IDDUP|||||" & SNS_UID & "는 사용할 수 없는 아이디 입니다."
		Response.End
END IF
oRs.Close





	


oConn.BeginTrans


'# SNS 회원가입 입력
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Member_SNS_Join_Insert"

		.Parameters.Append .CreateParameter("@UserID",			 adVarChar,	 adParamInput, 30, SNS_UID)
		.Parameters.Append .CreateParameter("@Name",			 adVarChar,	 adParamInput, 50, SNS_KName)
		.Parameters.Append .CreateParameter("@Email",			 adVarChar,	 adParamInput, 50, snsEmail)
		.Parameters.Append .CreateParameter("@SNSKind",			 adVarChar,	 adParamInput, 30, SNS_Kind)
		.Parameters.Append .CreateParameter("@SmsFlag",			 adChar,	 adParamInput,  1, SMSFlag)
		.Parameters.Append .CreateParameter("@EmailFlag",		 adChar,	 adParamInput,  1, EmailFlag)
		.Parameters.Append .CreateParameter("@JoinLocation",	 adChar,	 adParamInput,  1, "A")
		.Parameters.Append .CreateParameter("@CreateIP",		 adVarChar,	 adParamInput, 15, U_IP)

		.Parameters.Append .CreateParameter("@MemberNum",		 adInteger,	 adParamOutput,  , 0)
		
		.Execute, , adExecuteNoRecords
		MemberNum = .Parameters("@MemberNum").Value
END WITH
SET oCmd = Nothing





'# SNS정보 입력
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Member_SNS_Insert"

		.Parameters.Append .CreateParameter("@MemberNum",	adInteger,	 adParamInput,    , MemberNum)
		.Parameters.Append .CreateParameter("@SNSKind",		adVarChar,	 adParamInput,  30, SNS_Kind)
		.Parameters.Append .CreateParameter("@SnsID",		adVarChar,	 adParamInput,  50, SNS_UID)
		.Parameters.Append .CreateParameter("@Email",		adVarChar,	 adParamInput,  50, snsEmail)
		.Parameters.Append .CreateParameter("@CreateIP",	adVarChar,	 adParamInput,  50, U_IP)

		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing


IF Err.number = 0 THEN
		oConn.CommitTrans
ELSE
		oConn.RollbackTrans
		SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||입력처리 중 오류가 발생하였습니다."
		Response.End
END IF




'# 로그인 정보 입력
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Member_Login_Insert"
	
		.Parameters.Append .CreateParameter("@MemberNum",	 adInteger,	 adParamInput,   , MemberNum)
		.Parameters.Append .CreateParameter("@Location",	 adChar,	 adParamInput,  1, "A")
		.Parameters.Append .CreateParameter("@LoginIP",		 adVarChar,	 adParamInput, 15, U_IP)

		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing



Response.Cookies("UIP")				 = Encrypt(U_IP)
Response.Cookies("UMFLAG")			 = Encrypt("N")
Response.Cookies("UNUM")			 = Encrypt(MemberNum)
Response.Cookies("UID")				 = Encrypt(SNS_UID)
Response.Cookies("UNAME")			 = Encrypt(SNS_KName)
Response.Cookies("UEFLAG")			 = Encrypt("N")
Response.Cookies("UETYPE")			 = Encrypt("P")
Response.Cookies("UGROUP")			 = Encrypt(1000)
Response.Cookies("UPOINTRATE")		 = Encrypt(0)
Response.Cookies("USNSKIND")		 = Encrypt(SNS_Kind)

'# SNS 회원정보 초기화
Response.Cookies("SNS_UID")					 = ""
Response.Cookies("SNS_Kind")					 = ""
Response.Cookies("SNS_Email")				 = ""
Response.Cookies("SNS_KName")				 = ""
Response.Cookies("SNS_UserID")				 = ""
Response.Cookies("SNS_UNUM")					 = ""








SET oRs = Nothing
oConn.Close
SET oConn = Nothing




	
Response.Write "OK|||||" & SNS_ProgID
%>