<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'CartProductSelectUpdate.asp - 장바구니 상품 선택/해제 처리 페이지
'Date		: 2018.12.27
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

DIM CartIdx
DIM IsSelected
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


CartIdx			= sqlFilter(Request("CartIdx"))
IsSelected		= sqlFilter(Request("Flag"))


IF CartIdx = "" OR IsSelected = "" THEN
		Response.Write "FAIL|||||상품을 선택해 주십시오."
		Response.End
END IF


SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성



'# 장바구니 담기
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Cart_Update_For_IsSelected"

		.Parameters.Append .CreateParameter("@Idx",				adInteger,	adParamInput,    ,	 CartIdx)
		.Parameters.Append .CreateParameter("@IsSelected",		adChar,		adParamInput,   1,	 IsSelected)
		.Parameters.Append .CreateParameter("@UpdateID",		adVarChar,	adParamInput,  20,	 U_NUM)
		.Parameters.Append .CreateParameter("@UpdateIP",		adVarChar,	adParamInput,  15,	 U_IP)

		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing

IF Err.Number <> 0 THEN
		SET oRs = Nothing
		oConn.Close
		SET oConn = Nothing

		Response.Write "FAIL|||||장바구니 선택 처리중 오류가 발생하였습니다."
		Response.End
END IF





Response.Write "OK|||||"


Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>