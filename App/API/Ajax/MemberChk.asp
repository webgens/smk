<%@ Language=VBScript codepage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************/
'LoginOk.asp - 로그인 처리
'Date		: 2015.12.14
'Update	: 
'Writer		: Jongho Lee
'/****************************************************************************************/

'//페이지 응답헤더 설정------------------------------------------------------
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'Response.CharSet = "euc-kr"
'//-------------------------------------------------------------------------------
%>

<!-- #include Virtual = "/ADO/ADODBCommon.asp" -->
<!-- #include Virtual = "/Common/Common.asp" -->
<!-- #include Virtual = "/Common/SetInfo.asp" -->
<!-- #include Virtual = "/Common/md5.asp" -->

<%
'/****************************************************************************************/
'변수 선언 START
'-----------------------------------------------------------------------------------------------------------'
DIM oConn						'# ADODB Connection 개체
DIM oRs							'# ADODB Recordset 개체
DIM oRs1							'# ADODB Recordset 개체
DIM oCmd						'# ADODB Command 개체

Dim MemberNum
Dim UID
Dim Email
Dim KName
Dim SNSKind

Dim UserID
Dim DelFlag
Dim SNSChangeFlag
Dim MemberFlag
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'

UID			 = sqlFilter(Request("UID"))
Email		 = sqlFilter(Request("Email"))
KName		 = sqlFilter(Request("KName"))
SNSKind		 = sqlFilter(Request("SNSKind"))

SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성

IF U_NUM = "" THEN
	'등록된 회원인지 체크
	Set oCmd = Server.CreateObject("ADODB.Command")
	WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Front_EShop_Member_SNS_Select_By_SNSID"
		.Parameters.Append .CreateParameter("@SNSKind", adVarChar, adParamInput, 1, SNSKind)
		.Parameters.Append .CreateParameter("@SNSID", adVarChar, adParamInput, 255, UID)
	END WITH
	oRs.CursorLocation = adUseClient
	oRs.Open oCmd, , adOpenStatic, adLockReadOnly
	Set oCmd = Nothing

	If oRs.eof Then
		Response.Cookies("SNS_UID")		= Encrypt(UID)
		Response.Cookies("SNS_Email")	= Encrypt(Email)
		Response.Cookies("SNS_KName")	= Encrypt(KName)
		Response.Cookies("SNS_Kind")		= Encrypt(SNSKind)

		'Response.Write "FAIL_JOIN|||||일치하는 회원 정보가 없습니다. [01] 계정연결 및 간편로그인 이동(연결 SNS계정없음.)|||||/ASP/Member/Join.asp"
		Response.Write "FAIL_JOIN||||||||||/ASP/Member/snsGate.asp"
		Response.End
	ELSE
		MemberNum = oRs("MemberNum")
	End If
	oRs.Close
ELSE
		MemberNum = U_NUM
END IF

'# 회원정보 체크
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_Member_Select_By_MemberNum"
	
		.Parameters.Append .CreateParameter("@MemberNum", adInteger, adParamInput,  , MemberNum)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		UserID					= oRs("UserID")
		DelFlag					= oRs("DelFlag")
		SNSChangeFlag			= oRs("SNSChangeFlag")
		MemberFlag				= oRs("MemberFlag")
		Response.Cookies("SNS_UserID")	= Encrypt(UserID)
		Response.Cookies("SNS_UNUM")		= Encrypt(MemberNum)
ELSE
		Response.Cookies("SNS_UID")		= Encrypt(UID)
		Response.Cookies("SNS_Email")	= Encrypt(Email)
		Response.Cookies("SNS_KName")	= Encrypt(KName)
		Response.Cookies("SNS_Kind")		= Encrypt(SNSKind)

		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		'Response.Write "FAIL||일치하는 회원 정보가 없습니다. [02] 계정연결 및 간편로그인 이동(회원정보 없음)||"
		Response.Write "FAIL_JOIN||||||||||/ASP/Member/snsGate.asp"
		Response.End
END IF
oRs.Close


IF DelFlag = "Y" THEN
		SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||회원님은 탈퇴하신 정보가 있습니다. 재가입하여 주시기 바랍니다.|||||/ASP/Member/Join.asp"
		Response.End
END IF


'# IF MemberFlag = "N" THEN	'// 간편로그인 사용자(회원 미전환)
'# 		SET oRs = Nothing : oConn.Close : SET oConn = Nothing
'# 		Response.Write "FAIL_JOIN|||||간편로그인은 더 이상 지원하지 않습니다. 간편로그인 정보로 정회원 가입 전환 하시겠습니까?|||||/ASP/Member/Join.asp"
'# 		Response.End
'# END IF

Response.Write "OK||||||||||"


Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>
