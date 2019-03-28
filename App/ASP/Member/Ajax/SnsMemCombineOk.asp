<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'JoinForm.asp - SNS회원 정회원으로 통합
'Date		: 2018.12.20
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


DIM JoinType
DIM CombineID
DIM CombineNum
DIM SDupInfo
DIM ParentSDupInfo
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

JoinType			= LCASE(TRIM(sqlFilter(Request("JoinType"))))
CombineID			= LCASE(TRIM(sqlFilter(Request("CombineID"))))
SDupInfo			= LCASE(TRIM(sqlFilter(Request("SDupInfo"))))
ParentSDupInfo		= LCASE(TRIM(sqlFilter(Request("ParentSDupInfo"))))


IF CombineID = "" THEN
		Response.Write "FAIL|||||통합ID 정보가 없습니다."
		Response.End
END IF

IF JoinType = "U" AND SDupInfo = "" THEN
		Response.Write "FAIL|||||본인인증 정보가 없습니다."
		Response.End
END IF

IF JoinType = "D" AND ParentSDupInfo = "" THEN
		Response.Write "FAIL|||||보호자 본인인증 정보가 없습니다."
		Response.End
END IF

IF TRIM(U_ID) = "" THEN
		Response.Write "FAIL|||||아이디 정보가 없습니다."
		Response.End
END IF



SET oConn			 = ConnectionOpen()							'# 커넥션 생성
SET oRs				 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성


'# 통합 회원번호(정회원)
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Member_Select_By_UserID"

		.Parameters.Append .CreateParameter("@UserID", adVarChar, adParamInput, 30, CombineID)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		CombineNum = oRs("MemberNum")
ELSE
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||회원정보가 일치하지 않습니다."
		Response.End
END IF
oRs.Close



oConn.BeginTrans



'# 회원전환(SNS > 정회원)
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Member_SNS_Combine"

		.Parameters.Append .CreateParameter("@UserID",		adVarChar,	 adParamInput, 30, U_ID)
		.Parameters.Append .CreateParameter("@MemberNum",	adBigInt,	 adParamInput,   , U_NUM)
		.Parameters.Append .CreateParameter("@CombineID",	adVarChar,	 adParamInput, 30, CombineID)
		.Parameters.Append .CreateParameter("@CombineNum",	adInteger,	 adParamInput,   , CombineNum)
		.Parameters.Append .CreateParameter("@CreateIP",	adVarChar,	 adParamInput, 15, U_IP)

		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing



oConn.CommitTrans





SET oRs = Nothing
oConn.Close
SET oConn = Nothing


'# 로그아웃 처리
Response.Cookies("UIP").Expires			 = Now - 1000
Response.Cookies("UMFLAG").Expires		 = Now - 1000
Response.Cookies("UNUM").Expires		 = Now - 1000
Response.Cookies("UID").Expires			 = Now - 1000
Response.Cookies("UNAME").Expires		 = Now - 1000
Response.Cookies("UETYPE").Expires		 = Now - 1000
Response.Cookies("UETYPE").Expires		 = Now - 1000
Response.Cookies("UGROUP").Expires		 = Now - 1000



Response.Write "OK|||||계정 통합처리가 완료되었습니다."
%>