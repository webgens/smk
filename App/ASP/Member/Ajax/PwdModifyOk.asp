<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'PwdModifyOk.asp - 비밀번호 수정 처리
'Date		: 2018.11.27
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
PageCode3 = "03"
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


DIM MemberNum
DIM Pwd
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'
	

MemberNum		 = Request.Cookies("PW_MemberNum")
Pwd				 = sqlFilter(Request("Pwd"))

IF MemberNum = "" THEN
		Response.Write "FAIL|||||회원 정보가 없습니다. 비밀번호 찾기를 다시 실행하여 주십시오."
		Response.End
END IF
IF Pwd = "" THEN
		Response.Write "FAIL|||||비밀번호 수정 정보가 없습니다. 다시 입력하여 주십시오."
		Response.End
END IF


Pwd				 = md5(LCase(Pwd))
	


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성




SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_Member_Update_For_Pwd"

		.Parameters.Append .CreateParameter("@MemberNum",	 adInteger,	 adParamInput,    , MemberNum)
		.Parameters.Append .CreateParameter("@Pwd",			 adVarChar,	 adParamInput,  50, Pwd)
		.Parameters.Append .CreateParameter("@UpdateID",	 adVarChar,	 adParamInput,  20, MemberNum)
		.Parameters.Append .CreateParameter("@UpdateIP",	 adVarChar,	 adParamInput,  15, U_IP)

		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing




SET oRs = Nothing
oConn.Close
SET oConn = Nothing


Response.Write "OK|||||"
%>