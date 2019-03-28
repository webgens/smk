<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'MyInfoModifyPwChkOk.asp - 나의정보 수정 비밀번호 확인처리
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
PageCode3 = "03"
PageCode4 = "01"
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

DIM MemberNum
DIM Pwd
DIM PwdMd5
DIM PHPCrypt
DIM PwdEnc
DIM DB_Pwd
DIM DB_OldPwd
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'
	
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
ELSE
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||회원정보가 일치하지 않습니다."
		Response.End
END IF
oRs.Close

IF TRIM(DB_Pwd) <> "" THEN
		IF DB_Pwd <> PwdMd5 THEN
				SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "FAIL|||||현재 비밀번호가 일치하지 않습니다. [01]"
				Response.End
		END IF
ELSE
		IF DB_OldPwd <> PwdEnc THEN
				SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "FAIL|||||현재 비밀번호가 일치하지 않습니다. [02]"
				Response.End
		END IF
END IF



SET oRs = Nothing
oConn.Close
SET oConn = Nothing


Response.Write "OK|||||"
%>