<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'MyInfoModifyOK.asp - 회원정보 수정 - 내용입력
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
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/md5.asp" -->

<%
IF U_ID = "" THEN
		Response.Write "FAIL|||||로그아웃 되어 있습니다."
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


DIM SMSFlag
DIM EmailFlag
DIM FTFlag						'# 14세 구분 (U:14세이상 / D:14세미만)

DIM Pwd
DIM PwdMd5

DIM Name
DIM Birth
DIM Sex
DIM ZipCode
DIM Addr1
DIM Addr2
DIM HP
DIM Email

DIM ParentName
DIM ParentBirth
DIM ParentHP
DIM ParentEmail

DIM ChkBirth
DIM ChkParentBirth
DIM MemberNum

'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

SMSFlag				 = TRIM(sqlFilter(Request("SMSFlag")))
EmailFlag			 = TRIM(sqlFilter(Request("EMailFlag")))
FTFlag				 = TRIM(sqlFilter(Request("FTFlag")))
	
Pwd					 = TRIM(sqlFilter(Request("Pwd")))
PwdMd5				 = md5(LCase(Pwd))

Name				 = TRIM(sqlFilter(Request("Name")))
Birth				 = TRIM(sqlFilter(Request("Birth")))
Sex					 = TRIM(sqlFilter(Request("Sex")))
ZipCode				 = TRIM(sqlFilter(Request("ZipCode")))
Addr1				 = TRIM(sqlFilter(Request("Addr1")))
Addr2				 = TRIM(sqlFilter(Request("Addr2")))

HP					 = ChgTel(TRIM(sqlFilter(Request("HP1"))) & TRIM(sqlFilter(Request("HP23"))))
Email				 = TRIM(sqlFilter(Request("Email")))

ParentName			 = TRIM(sqlFilter(Request("ParentName")))
ParentBirth			 = TRIM(sqlFilter(Request("ParentBirth")))
ParentHP			 = ChgTel(TRIM(sqlFilter(Request("PHP1"))) & TRIM(sqlFilter(Request("PHP2"))))
ParentEmail			 = TRIM(sqlFilter(Request("ParentEmail")))




IF FTFlag = "" THEN
		Response.Write "FAIL|||||만14세 구분 정보가 없습니다."
		Response.End
END IF

IF TRIM(Pwd) = "" THEN
		Response.Write "FAIL|||||비밀번호 정보가 없습니다."
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

IF TRIM(Zipcode) = "" OR TRIM(Addr1) = "" THEN
		Response.Write "FAIL|||||우편번호 정보가 없습니다."
		Response.End
END IF

IF TRIM(Addr2) = "" THEN
		Response.Write "FAIL|||||나머지주소 정보가 없습니다."
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


IF FTFlag = "Y" THEN
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
		





SET oConn			 = ConnectionOpen()							'# 커넥션 생성
SET oRs		= Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성

	



'# 회원정보 수정
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Mobile_EShop_Member_Update"

		.Parameters.Append .CreateParameter("@UserID",			 adVarChar,	 adParamInput, 30, U_ID)
		.Parameters.Append .CreateParameter("@Pwd",				 adVarChar,	 adParamInput, 50, PwdMd5)
		.Parameters.Append .CreateParameter("@Name",			 adVarChar,	 adParamInput, 50, Name)
		.Parameters.Append .CreateParameter("@Birth",			 adVarChar,	 adParamInput,  8, Birth)
		.Parameters.Append .CreateParameter("@Sex",				 adChar,	 adParamInput,  1, Sex)
		.Parameters.Append .CreateParameter("@Zipcode",			 adVarChar,	 adParamInput,  6, Zipcode)
		.Parameters.Append .CreateParameter("@Address1",		 adVarChar,	 adParamInput,400, Addr1)
		.Parameters.Append .CreateParameter("@Address2",		 adVarChar,	 adParamInput,400, Addr2)
		.Parameters.Append .CreateParameter("@HP",				 adVarChar,	 adParamInput, 14, HP)
		.Parameters.Append .CreateParameter("@Email",			 adVarChar,	 adParamInput, 50, Email)
		.Parameters.Append .CreateParameter("@SmsFlag",			 adChar,	 adParamInput,  1, SMSFlag)
		.Parameters.Append .CreateParameter("@EmailFlag",		 adChar,	 adParamInput,  1, EmailFlag)
		.Parameters.Append .CreateParameter("@FTFlag",			 adChar,	 adParamInput,  1, FTFlag)
		.Parameters.Append .CreateParameter("@ParentName",		 adVarChar,	 adParamInput, 20, ParentName)
		.Parameters.Append .CreateParameter("@ParentBirth",		 adVarChar,	 adParamInput,  8, ParentBirth)
		.Parameters.Append .CreateParameter("@ParentHP",		 adVarChar,	 adParamInput, 14, ParentHP)
		.Parameters.Append .CreateParameter("@ParentEmail",		 adVarChar,	 adParamInput, 50, ParentEmail)
		.Parameters.Append .CreateParameter("@UpdateID",		 adVarChar,	 adParamInput, 30, U_NUM)
		.Parameters.Append .CreateParameter("@UpdateIP",		 adVarChar,	 adParamInput, 15, U_IP)


		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing


'# Response.Write "FTFlag				 = " & FTFlag & "<br>"
'# Response.Write "SMSFlag				 = " & SMSFlag & "<br>"
'# Response.Write "EmailFlag			 = " & EMailFlag & "<br>"
'# 
'# 
'# 
'# Response.Write "Name				 = " & Name & "<br>"
'# Response.Write "Birth				 = " & Birth & "<br>"
'# Response.Write "Sex					 = " & Sex & "<br>"
'# Response.Write "HP					 = " & HP & "<br>"
'# Response.Write "Email				 = " & Email & "<br>"
'# 
'# Response.Write "ParentName			 = " & ParentName & "<br>"
'# Response.Write "ParentBirth			 = " & ParentBirth & "<br>"
'# Response.Write "ParentHP			 = " & ParentHP & "<br>"
'# Response.Write "ParentEmail			 = " & ParentEmail & "<br>"
'# 
'# 
'# 	
'# Response.Write "OK|||||"
'# Response.End







Set oRs = Nothing
oConn.Close
SET oConn = Nothing





Response.Write "OK|||||정보수정이 완료되었습니다."
%>