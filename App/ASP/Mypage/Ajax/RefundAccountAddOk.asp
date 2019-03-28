<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'RefundAccountAddOk.asp - 마이페이지 > 회원정보 수정 > 환불계좌 등록/수정 처리
'Date		: 2019.01.19
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
<!-- #include virtual="/Common/CheckID_Ajax.asp" -->

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

Dim Idx
Dim BankCode
Dim AccountNum
Dim AccountName
Dim CreateDT

'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


Idx					= sqlFilter(Request("Idx"))
BankCode			= sqlFilter(Request("BankCode"))
AccountNum			= sqlFilter(Request("AccountNum"))
AccountName			= sqlFilter(Request("AccountName"))


IF BankCode = "" OR AccountNum = "" OR AccountName = "" THEN
		Response.Write "FAIL|||||환불계좌 입력정보가 부족합니다."
		Response.End
END IF


SET oConn	= ConnectionOpen()	'//커넥션 생성
SET oRs		= Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성


IF Idx = "" THEN

	'-----------------------------------------------------------------------------------------------------------'	
	'# 환불계좌 등록 Start
	'-----------------------------------------------------------------------------------------------------------'	
	SET oCmd = Server.CreateObject("ADODB.Command")
	WITH oCmd
			.ActiveConnection	 = oConn
			.CommandType		 = adCmdStoredProc
			.CommandText		 = "USP_Front_EShop_Member_RefundAccount_Insert"

			.Parameters.Append .CreateParameter("@MemberNum",			adInteger,	adParamInput,     ,	 U_NUM)
			.Parameters.Append .CreateParameter("@BankCode",			adChar,		adParamInput,    2,	 BankCode)
			.Parameters.Append .CreateParameter("@AccountNum",			adVarChar,	adParamInput,   20,	 AccountNum)
			.Parameters.Append .CreateParameter("@AccountName",			adVarChar,	adParamInput,   50,	 AccountName)
			.Parameters.Append .CreateParameter("@CreateID",			adVarChar,	adParamInput,   20,	 U_ID)
			.Parameters.Append .CreateParameter("@CreateIP",			adVarChar,	adParamInput,   15,	 U_IP)

			.Execute, , adExecuteNoRecords

	END WITH
	SET oCmd = Nothing


	IF Err.Number <> 0 THEN

			oConn.Close
			SET oConn = Nothing

			Response.Write "FAIL|||||환불계좌 등록 중 오류가 발생하였습니다."
			Response.End
	END IF
	'-----------------------------------------------------------------------------------------------------------'	
	'# 상품후기 등록 End
	'-----------------------------------------------------------------------------------------------------------'	

	Response.Write "OK|||||등록되었습니다."
ELSE

	'-----------------------------------------------------------------------------------------------------------'	
	'# 환불계좌 수정 Start
	'-----------------------------------------------------------------------------------------------------------'	
	SET oCmd = Server.CreateObject("ADODB.Command")
	WITH oCmd
			.ActiveConnection	 = oConn
			.CommandType		 = adCmdStoredProc
			.CommandText		 = "USP_Front_EShop_Member_RefundAccount_Update"

			.Parameters.Append .CreateParameter("@Idx",					adInteger,	adParamInput,     ,	 Idx)
			.Parameters.Append .CreateParameter("@MemberNum",			adInteger,	adParamInput,     ,	 U_NUM)
			.Parameters.Append .CreateParameter("@BankCode",			adChar,		adParamInput,    2,	 BankCode)
			.Parameters.Append .CreateParameter("@AccountNum",			adVarChar,	adParamInput,   20,	 AccountNum)
			.Parameters.Append .CreateParameter("@AccountName",			adVarChar,	adParamInput,   50,	 AccountName)
			.Parameters.Append .CreateParameter("@UpdateID",			adVarChar,	adParamInput,   20,	 U_ID)
			.Parameters.Append .CreateParameter("@UpdateIP",			adVarChar,	adParamInput,   15,	 U_IP)

			.Execute, , adExecuteNoRecords

	END WITH
	SET oCmd = Nothing


	IF Err.Number <> 0 THEN

			oConn.Close
			SET oConn = Nothing

			Response.Write "FAIL|||||환불계좌 수정 중 오류가 발생하였습니다."
			Response.End
	END IF
	'-----------------------------------------------------------------------------------------------------------'	
	'# 상품후기 수정 End
	'-----------------------------------------------------------------------------------------------------------'	

	Response.Write "OK|||||수정되었습니다."
END IF






Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>