<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'MySnsConnection.asp - SNS 계정정보 연결(정회원)
'Date		: 2018.12.05
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
<!-- #include virtual="/Common/CheckID_Ajax.asp" -->

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
	
Dim UID
Dim Email
Dim SNSKind
Dim	MemberFlag
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


UID				 = sqlFilter(Request("UID"))
Email			 = sqlFilter(Request("Email"))
SNSKind			 = sqlFilter(Request("SNSKind"))

IF U_NUM = "" THEN
		Response.Write "FAIL|||||회원번호가 없습니다."
		Response.End
END IF
IF UID = "" THEN
		Response.Write "FAIL|||||SNS계정 ID가 없습니다."
		Response.End
END IF
IF SNSKind = "" THEN
		Response.Write "FAIL|||||SNS 구분코드가 없습니다."
		Response.End
END IF


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성


'등록된 회원인지 체크 (정회원이 SNS 계정연결시 기존 정보 있으면 계정연결 안내)
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Member_SNS_Select_By_SDupInfo_Check"

		.Parameters.Append .CreateParameter("@SNSKind", adChar,		adParamInput,	 1, SNSKind)
		.Parameters.Append .CreateParameter("@SNSID"  , adVarChar,	adParamInput,	50, UID)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		IF oRs("U_NUM") <> "" AND oRs("MemberFlag") = "Y" THEN
				Response.Write "FAIL|||||이미 정회원으로 가입된 정보가 있습니다.<br>아이디/비밀번호 찾기를 이용해 주십시오."
		ELSEIF oRs("U_NUM") <> "" AND oRs("MemberFlag") = "N" THEN
				Response.Cookies("SNS_MEMNUM") = oRs("U_NUM")
				Response.Write "DIDUP|||||이미 SNS회원으로 등록된 정보가 있습니다.<br>현재 로그인 중인 회원정보로 통합 하시겠습니까?"
		END IF
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.End
END IF
oRs.Close


'# SNS계정 연결처리
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Member_SNS_Insert"
	
		.Parameters.Append .CreateParameter("@MemberNum",	 adBigInt,	 adParamInput,    , U_NUM)
		.Parameters.Append .CreateParameter("@SNSKind",		 adChar,	 adParamInput,   1, SNSKind)
		.Parameters.Append .CreateParameter("@SnsID",		 adVarChar,	 adParamInput,  50, UID)
		.Parameters.Append .CreateParameter("@Email",		 adVarChar,	 adParamInput,  50, Email)
		.Parameters.Append .CreateParameter("@CreateIP",	 adVarChar,	 adParamInput,  15, U_IP)

		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing

Response.Write "OK|||||"


SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>