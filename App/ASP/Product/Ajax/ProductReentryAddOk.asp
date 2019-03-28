<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'ProductReentryAddOk.asp - 상품 재입고 알림 등록
'Date		: 2019.01.09
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
DIM RIdx
Dim SizeCD
DIM Mobile
Dim Mobile1
Dim Mobile2
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'

ProductCode				= sqlFilter(Request("ProductCode"))
RIdx					= sqlFilter(Request("RIdx"))
SizeCD					= sqlFilter(Request("Reentry_SizeCD"))
Mobile1					= sqlFilter(Request("Mobile1"))
Mobile2					= sqlFilter(Request("Mobile2"))

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

IF RIdx = "" THEN
		Response.Write "FAIL|||||재입고 알림 정보가 없습니다."
		Response.End
END IF

IF SizeCD = "" THEN
		Response.Write "FAIL|||||상품 사이즈 정보가 없습니다."
		Response.End
END IF

IF Mobile = "" THEN
		Response.Write "FAIL|||||연락처 정보가 없습니다."
		Response.End
END IF

SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성


SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Product_Reentry_Select_For_Count"

		.Parameters.Append .CreateParameter("@MemberNum",	 adInteger, adParamInput,   , U_Num)
		.Parameters.Append .CreateParameter("@ProductCode",	 adInteger, adParamInput,   , ProductCode)
		.Parameters.Append .CreateParameter("@RIdx",		 adInteger, adParamInput,   , RIdx)
		.Parameters.Append .CreateParameter("@SizeCD",		 adVarChar, adParamInput, 50, SizeCD)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
	IF oRs(0) > 0 THEN
		oRs.Close : Set oRs = Nothing : oConn.Close : Set oConn = Nothing
		Response.Write "FAIL|||||이미 신청한 재입고 알림신청 내용이 존재합니다."
		Response.End
	END IF
END IF
SET oRs = oRs.NextRecordset

IF NOT oRs.EOF THEN
	IF oRs(0) >= 10 THEN
		oRs.Close : Set oRs = Nothing : oConn.Close : Set oConn = Nothing
		Response.Write "FAIL|||||재입고 알림 신청은 최대 10개 상품까지만 등록됩니다."
		Response.End
	END IF
END IF
oRs.Close

oConn.BeginTrans

SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Product_Reentry_Insert"
		.Parameters.Append .CreateParameter("@MemberNum",	 adInteger, adParamInput,    , U_Num)
		.Parameters.Append .CreateParameter("@ProductCode",	 adInteger, adParamInput,    , ProductCode)
		.Parameters.Append .CreateParameter("@RIdx",		 adInteger, adParamInput,    , RIdx)
		.Parameters.Append .CreateParameter("@SizeCD",		 adVarChar, adParamInput,  50, SizeCD)
		.Parameters.Append .CreateParameter("@Mobile",		 adVarChar, adParamInput,  20, Mobile)
		.Parameters.Append .CreateParameter("@CreateID",	 adVarChar, adParamInput,  20, U_ID)
		.Parameters.Append .CreateParameter("@CreateIP",	 adVarChar, adParamInput,  15, U_IP)
		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing

IF Err.number <> 0 THEN
	oConn.RollbackTrans
	Set oRs = Nothing
	oConn.Close
	Set oConn = Nothing
	
	Response.Write "FAIL|||||정상적으로 재입고 알림 신청이 되지 않았습니다."
	Response.End
END IF

oConn.CommitTrans

Response.Write "OK|||||"

Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>