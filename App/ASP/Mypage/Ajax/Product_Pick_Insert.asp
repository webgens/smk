<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'Product_Pick_Insert.asp - 상품 찜하기
'Date		: 2018.12.24
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
DIM oCmd											'# ADODB Command 개체

DIM i
DIM j
DiM x
DIM y

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절

DIM ProductCode
Dim PickCount
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'

IF U_NUM = "" AND N_NAME = "" THEN
		Response.Write "LOGIN|||||로그인 정보가 없습니다. 로그인 페이지로 이동하시겠습니까?"
		Response.End
END IF


ProductCode			= sqlFilter(Request("ProductCode"))

IF ProductCode = "" THEN
		Response.Write "FAIL|||||상품정보가 없습니다."
		Response.End
END IF

SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성


' 이미 찜한 상품인지 체크
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Front_EShop_Product_Pick_Select_By_ProductCode_MemberNum"

		.Parameters.Append .CreateParameter("@MemberNum",	adInteger,	adParamInput, ,		U_NUM)
		.Parameters.Append .CreateParameter("@ProductCode",	adInteger,	adParamInput, ,		ProductCode)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing
If Not oRs.EOF Then
	oRs.Close
	Set oRs = Nothing
	oConn.Close
	Set oConn = Nothing

	Response.Write "FAIL|||||이미 등록한 상품입니다."
	Response.End
End If
oRs.Close

' 찜한 상품이 몇개인지 체크
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Front_EShop_Product_Pick_Select_By_MemberNum"

		.Parameters.Append .CreateParameter("@MemberNum",	adInteger,	adParamInput, ,		U_NUM)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing
PickCount = oRs.RecordCount

If PickCount >= 30 Then
	oRs.Close
	Set oRs = Nothing
	oConn.Close
	Set oConn = Nothing

	Response.Write "FAIL|||||찜한 상품은 최대 30개까지 저장 됩니다."
	Response.End
End If
oRs.Close


SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Product_Pick_Insert"

		.Parameters.Append .CreateParameter("@MemberNum",			adInteger,	adParamInput,     ,	 U_NUM)
		.Parameters.Append .CreateParameter("@ProductCode",			adInteger,	adParamInput,     ,	 ProductCode)
		.Parameters.Append .CreateParameter("@CreateID",			adVarChar,	adParamInput,   20,	 U_NUM)
		.Parameters.Append .CreateParameter("@CreateIP",			adVarChar,	adParamInput,   15,	 U_IP)
		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing


Response.Write "OK|||||"

Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>