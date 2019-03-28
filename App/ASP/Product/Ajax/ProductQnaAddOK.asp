<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'ProductAdd.asp - 상품문의 페이지
'Date		: 2019.01.02
'Update	: 
'/****************************************************************************************'

'//페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'//---------------------------------------------------------------------------------------'
%>

<!-- #include Virtual = "/ADO/ADODBCommon.asp" -->
<!-- #include Virtual = "/Common/Common.asp" -->
<!-- #include Virtual = "/Common/SetInfo.asp" -->

<%
'/****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn											'# ADODB Connection 개체
DIM oRs												'# ADODB Recordset 개체
DIM oRs1											'# ADODB Recordset 개체
DIM oCmd											'# ADODB Command 개체

DIM i
DIM j
DiM x
DIM y

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절

DIM ProductCode
Dim ClassName
DIM Title
DIM Contents
DIM SMSReturnFlag
DIM Mobile
Dim Mobile1
Dim Mobile2
DIM EMailReturnFlag
DIM Email
DIM SecretFlag
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


ProductCode				= sqlFilter(Request("ProductCode"))
ClassName				= sqlFilter(Request("ClassName"))
Title					= sqlFilter(Request("Title"))
Contents				= sqlFilter(Request("Contents"))
SMSReturnFlag			= sqlFilter(Request("SMSReturnFlag"))
Mobile1					= sqlFilter(Request("Mobile1"))
Mobile2					= sqlFilter(Request("Mobile2"))
EMailReturnFlag			= sqlFilter(Request("EMailReturnFlag"))
Email					= sqlFilter(Request("Email"))
SecretFlag				= sqlFilter(Request("SecretFlag"))

If Mobile2 = "" Then
	Mobile = ""
Else
	Mobile = Mobile1 & Mobile2
End If

IF U_Num = "" THEN
		Response.Write "FAIL|||||로그인 정보가 없습니다."
		Response.End
END IF

IF ProductCode = "" THEN
		Response.Write "FAIL|||||상품 정보가 없습니다."
		Response.End
END IF

SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성

oConn.BeginTrans

SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Product_QNA_Insert"
		.Parameters.Append .CreateParameter("@MemberNum",	 adInteger, adParamInput, , U_Num)
		.Parameters.Append .CreateParameter("@ProductCode",	 adInteger, adParamInput, , ProductCode)
		.Parameters.Append .CreateParameter("@ClassName",	 adVarChar, adParamInput, 50, ClassName)
		.Parameters.Append .CreateParameter("@EMailReturnFlag",	 adInteger, adParamInput, , EMailReturnFlag)
		.Parameters.Append .CreateParameter("@Email",	 adVarChar, adParamInput, 50, Email)
		.Parameters.Append .CreateParameter("@SMSReturnFlag",	 adInteger, adParamInput, , SMSReturnFlag)
		.Parameters.Append .CreateParameter("@Mobile",	 adVarChar, adParamInput, 20, Mobile)
		.Parameters.Append .CreateParameter("@Title",	 adVarChar, adParamInput, 255, Title)
		.Parameters.Append .CreateParameter("@Contents",	 adLongVarChar, adParamInput,  10000000, Contents)
		.Parameters.Append .CreateParameter("@SecretFlag",	 adInteger, adParamInput, , SecretFlag)
		.Parameters.Append .CreateParameter("@CreateID",	 adVarChar, adParamInput,  20, U_ID)
		.Parameters.Append .CreateParameter("@CreateIP",	 adVarChar, adParamInput,  15, U_IP)
		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing


'-----------------------------------------------------------------------------------------------------------'
'문의 답변 SMS 알림 START
'-----------------------------------------------------------------------------------------------------------'
If SMSReturnFlag = "1" AND Mobile <> "" Then

	Mobile	= Replace(Mobile, "-", "")
	Dim SmsCode : SmsCode = "MEM_PDTQ"
	SET oCmd = Server.CreateObject("ADODB.Command")
	WITH oCmd
			.ActiveConnection	 = oConn
			.CommandType		 = adCmdStoredProc
			.CommandText		 = "USP_Admin_EShop_QNA_SMS_Send"

			.Parameters.Append .CreateParameter("@NAME",		 adVarChar,	 adParamInput,  50, U_NAME)
			.Parameters.Append .CreateParameter("@TITLE",		 adVarChar,	 adParamInput, 255, Title)
			.Parameters.Append .CreateParameter("@HP",			 adVarChar,	 adParamInput,  14, Mobile)
			.Parameters.Append .CreateParameter("@YEAR",		 adVarChar,	 adParamInput,   4, LEFT(DATE(),4) )
			.Parameters.Append .CreateParameter("@MONTH",		 adVarChar,	 adParamInput,   2, MID(DATE(), 6, 2) )
			.Parameters.Append .CreateParameter("@DAY",			 adVarChar,	 adParamInput,   2, MID(DATE(), 9, 2) )
			.Parameters.Append .CreateParameter("@SmsCode",		 adVarChar,	 adParamInput,  20,	SmsCode)

			.Execute, , adExecuteNoRecords
	END WITH
	SET oCmd = Nothing

End If
'-----------------------------------------------------------------------------------------------------------'
'문의 답변 SMS 알림 END
'-----------------------------------------------------------------------------------------------------------'


IF Err.number <> 0 THEN
	oConn.RollbackTrans
	Set oRs = Nothing
	oConn.Close
	Set oConn = Nothing
	
	Response.Write "FAIL|||||정상적으로 상품문의가 등록 되지 않았습니다."
	Response.End
END IF

oConn.CommitTrans

Response.Write "OK|||||"

Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>