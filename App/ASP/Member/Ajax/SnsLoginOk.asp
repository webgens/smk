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
PageCode3 = "00"
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
DIM UserID
DIM UID
DIM Email
DIM SNSKind



DIM DB_MemberNum
DIM DB_GroupCode
DIM DB_Name
DIM DB_EmployeeFlag
DIM DB_EmployeeType
DIM DB_DelFlag
DIM DB_DormancyFlag
DIM DB_MemberFlag
DIM DB_PointRate

DIM NewAgreementFlag : NewAgreementFlag = "N"
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


ProgID			 = Request("ProgID")
IF ProgID		 = "" THEN ProgID = "/"
DB_MemberNum	 = sqlFilter(Decrypt(Request.Cookies("SNS_UNUM")))
UserID			 = sqlFilter(Decrypt(Request.Cookies("SNS_UserID")))
UID				 = sqlFilter(Request("UID"))
Email			 = sqlFilter(Request("Email"))
SNSKind			 = sqlFilter(Request("SNSKind"))


IF DB_MemberNum = "" THEN
		Response.Write "FAIL|||||회원번호 정보가 없습니다."
		Response.End
END IF
IF UID = "" THEN
		Response.Write "FAIL|||||SNS 계정ID가 없습니다."
		Response.End
END IF
IF SNSKind = "" THEN
		Response.Write "FAIL|||||SNS 구분정보가 없습니다."
		Response.End
END IF




SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성


'# 회원 정보 - 아이디 찾기
SET oCmd = SErver.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Member_Select_By_SnsInfo"

		.Parameters.Append .CreateParameter("@MemberNum",	 adInteger,	 adParamInput,   , DB_MemberNum)
		.Parameters.Append .CreateParameter("@SNSID",		 adVarChar,	 adParamInput, 50, UID)
		.Parameters.Append .CreateParameter("@SNSKind",		 adChar,	 adParamInput,  1, SNSKind)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		DB_GroupCode	 = oRs("GroupCode")
		DB_Name			 = oRs("Name")
		DB_EmployeeFlag	 = oRs("EmployeeFlag")
		DB_EmployeeType	 = oRs("EmployeeType")
		DB_DelFlag		 = oRs("DelFlag")
		DB_DormancyFlag	 = oRs("DormancyFlag")
		DB_MemberFlag	 = oRs("MemberFlag")
		DB_PointRate	 = oRs("PointRate")
ELSE
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL_LOGIN|||||일치하는 회원정보가 없습니다."
		Response.End
END IF
oRs.Close


'# 휴면계정
IF DB_DormancyFlag = "Y" THEN

		SET oRs = Nothing : oConn.Close : SET oConn = Nothing
	
		Response.Cookies("TEMP_DOR")		 = Encrypt(DB_DormancyFlag)
		Response.Cookies("TEMP_NEW")		 = Encrypt(NewAgreementFlag)
		Response.Cookies("TEMP_MFLAG")		 = Encrypt(DB_MemberFlag)
		Response.Cookies("TEMP_PROGID")		 = Encrypt(ProgID)
		Response.Cookies("TEMP_UIP")		 = Encrypt(U_IP)
		Response.Cookies("TEMP_UNUM")		 = Encrypt(DB_MemberNum)
		Response.Cookies("TEMP_UID")		 = Encrypt(UserID)
		Response.Cookies("TEMP_UNAME")		 = Encrypt(DB_Name)
		Response.Cookies("TEMP_EFLAG")		 = Encrypt(DB_EmployeeFlag)
		Response.Cookies("TEMP_ETYPE")		 = Encrypt(DB_EmployeeType)
		Response.Cookies("TEMP_UGROUP")		 = Encrypt(DB_GroupCode)
		Response.Cookies("TEMP_UPOINTRATE")	 = Encrypt(DB_PointRate)

		Response.Write "DORMANCY|||||휴먼회원입니다.<br />휴먼회원 해제 후 이용하세요."
		Response.End

END IF



'# 정회원 일 경우 신규 약관 동의 여부 처리
IF DB_MemberFlag = "Y" THEN

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
	
		IF NewAgreementFlag = "N" THEN
				Response.Write "NEWAGREE|||||"
				Response.End
		END IF

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
Response.Cookies("UPOINTRATE")		 = Encrypt(DB_PointRate)
Response.Cookies("USNSKIND")		 = Encrypt(SNSKind)
'# SNS 회원정보 초기화
Response.Cookies("SNS_UID")		= ""
Response.Cookies("SNS_Kind")		= ""
Response.Cookies("SNS_Email")	= ""
Response.Cookies("SNS_KName")	= ""
Response.Cookies("SNS_UserID")	= ""
Response.Cookies("SNS_UNUM")		= ""

Response.Write "OK|||||"
Response.End



SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>