<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'FindIDOk.asp - 아이디 찾기 처리
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
PageCode3 = "02"
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


DIM FindIDType
DIM Name
DIM HP
DIM Email

DIM DB_UserID
DIM DB_MemberFlag
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'
	

FindIDType		 = sqlFilter(Request("FindIDType"))
Name			 = sqlFilter(Request("Name"))
HP				 = sqlFilter(Request("HP1")) & "-" & sqlFilter(Request("HP2")) & "-" & sqlFilter(Request("HP3"))
Email			 = sqlFilter(Request("Email"))


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성


IF FindIDType = "mobile" THEN

		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Member_Select_For_FIndID_By_HP"

				.Parameters.Append .CreateParameter("@Name",	 adVarChar, adParamInput, 50, Name)
				.Parameters.Append .CreateParameter("@HP",		 adVarChar, adParamInput, 14, HP)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		SET oCmd = Nothing

		IF NOT oRs.EOF THEN
				DB_UserID		 = oRs("UserID")
				DB_MemberFlag	 = oRs("MemberFlag")

				IF DB_MemberFlag = "Y" THEN
						oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
						Response.Write "OK|||||" & Name & "님의 아이디는 " & MaskUserID(DB_UserID) & " 입니다."
						Response.End
				ELSE
						oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
						Response.Write "NOTEXISTS|||||일치하는 회원이 없습니다."
						Response.End
				END IF
		ELSE
				oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "NOTEXISTS|||||일치하는 회원이 없습니다."
				Response.End
		END IF
		oRs.Close

ELSE

		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Member_Select_For_FIndID_By_Email"

				.Parameters.Append .CreateParameter("@Name",	 adVarChar, adParamInput, 50, Name)
				.Parameters.Append .CreateParameter("@Email",	 adVarChar, adParamInput, 50, Email)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		SET oCmd = Nothing

		IF NOT oRs.EOF THEN
				DB_UserID		 = oRs("UserID")
				DB_MemberFlag	 = oRs("MemberFlag")

				IF DB_MemberFlag = "Y" THEN
						oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
						Response.Write "OK|||||" & Name & "님의 아이디는 " & MaskUserID(DB_UserID) & " 입니다."
						Response.End
				ELSE
						oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
						Response.Write "NOTEXISTS|||||일치하는 회원이 없습니다."
						Response.End
				END IF
		ELSE
				oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "NOTEXISTS|||||일치하는 회원이 없습니다."
				Response.End
		END IF
		oRs.Close

END IF
%>