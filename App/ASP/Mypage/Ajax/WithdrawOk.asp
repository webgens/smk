<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'WithdrawOk.asp - 회원 탈퇴처리
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
PageCode1 = "05"
PageCode2 = "05"
PageCode3 = "03"
PageCode4 = "03"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include Virtual = "/Common/CheckID_Ajax.asp" -->
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

DIM wdReason
DIM MemberNum
DIM Pwd
DIM PwdMd5
DIM PHPCrypt
DIM PwdEnc
DIM DB_Pwd
DIM DB_OldPwd
DIM DelFlag
DIM MemberFlag
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'
	
wdReason		 = sqlFilter(Request("wdReason"))
Pwd				 = sqlFilter(Request("Pwd"))


IF Pwd = "" THEN
		Response.Write "FAIL|||||비밀번호 정보가 없습니다. 다시 입력하여 주십시오."
		Response.End
END IF
	

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

		.Parameters.Append .CreateParameter("@UserID", adVarChar, adParamInput, 30, U_ID)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		MemberNum			 = oRs("MemberNum")
		DB_Pwd				 = oRs("Pwd")
		DB_OldPwd			 = oRs("OldPwd")
		DelFlag				 = oRs("DelFlag")
		MemberFlag			 = oRs("MemberFlag")
ELSE
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "ID|||||회원정보가 없습니다."
		Response.End
END IF
oRs.Close

IF TRIM(DB_Pwd) <> "" THEN
		IF DB_Pwd <> PwdMd5 THEN
				SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "PWD|||||회원정보가 일치하지 않습니다.[01]"
				Response.End
		END IF
ELSE
		IF DB_OldPwd <> PwdEnc THEN
				SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "PWD|||||회원정보가 일치하지 않습니다.[02]"
				Response.End
		END IF
END IF
IF DelFlag="Y" THEN
	SET oRs = Nothing : oConn.Close : SET oConn = Nothing
	Response.Write "PWD|||||이미 탈퇴한 정보가 있습니다."
	Response.End
END IF
IF MemberFlag="N" THEN
	SET oRs = Nothing : oConn.Close : SET oConn = Nothing
	Response.Write "PWD|||||탈퇴할 수 있는 자격이 아닙니다."
	Response.End
END IF


oConn.BeginTrans

SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_Member_Update_For_Withdraw"

		.Parameters.Append .CreateParameter("@MemberNum",	 adInteger,	 adParamInput,    , MemberNum)
		.Parameters.Append .CreateParameter("@WReason",		 adChar,	 adParamInput,  2 , wdReason)
		.Parameters.Append .CreateParameter("@UpdateID",	 adVarChar,	 adParamInput,  20, MemberNum)
		.Parameters.Append .CreateParameter("@UpdateIP",	 adVarChar,	 adParamInput,  15, U_IP)

		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing

IF Err.Number = 0 THEN
		oConn.CommitTrans
ELSE
		oConn.RollbackTrans

		SET oRs = Nothing : oConn.Close : SET oConn = Nothing

		Response.Write "PWD|||||탈퇴 처리중 오류가 발생하였습니다."
		Response.End
END IF


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




Response.Write "OK||||||||||"&U_ID
%>