<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'MemSnsCombineOk.asp - 정회원 SNS계정 연결 시 기존SNS계정 통합처리
'Date		: 2018.12.18
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
PageCode4 = "02"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
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

Dim SNSID
Dim SNSNUM
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


SNSID		 = sqlFilter(Request("UID"))
SNSNUM		 = sqlFilter(Request.Cookies("SNS_MEMNUM"))

IF TRIM(SNSID) = "" OR TRIM(SNSNUM) = "" THEN
		Response.Write "FAIL|||||통합할 SNS 계정정보가 없습니다."
		Response.End
END IF


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성

'# 회원전환(SNS회원에서 정회원 전환)
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Member_SNS_Combine"

		.Parameters.Append .CreateParameter("@UserID",		adVarChar,	 adParamInput, 30, SNSID)
		.Parameters.Append .CreateParameter("@MemberNum",	adBigInt,	 adParamInput,   , SNSNUM)
		.Parameters.Append .CreateParameter("@CombineID",	adVarChar,	 adParamInput, 30, U_ID)
		.Parameters.Append .CreateParameter("@CombineNum",	adBigInt,	 adParamInput,   , U_NUM)
		.Parameters.Append .CreateParameter("@CreateIP",	adVarChar,	 adParamInput, 15, U_IP)

		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing



SET oRs = Nothing
oConn.Close
SET oConn = Nothing


Response.Cookies("SNS_MEMNUM") = ""

Response.Write "OK|||||계정 통합처리가 완료되었습니다."
%>