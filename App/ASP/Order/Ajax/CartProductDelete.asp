<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'CartProductDelete.asp - 장바구니 상품 삭제 처리
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

DIM Flag
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


Flag		= sqlFilter(Request("Flag"))


IF Flag = "" THEN
		Response.Write "FAIL|||||삭제할 상품을 선택해 주십시오."
		Response.End
END IF



SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성



'# 장바구니 삭제
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Cart_Delete"

		.Parameters.Append .CreateParameter("@CartID",			adVarChar,	adParamInput,  20,	 U_CARTID)
		.Parameters.Append .CreateParameter("@UserID",			adVarChar,	adParamInput,  20,	 U_NUM)
		.Parameters.Append .CreateParameter("@Flag",			adVarChar,	adParamInput,  20,	 Flag)

		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing

IF Err.Number <> 0 THEN
		SET oRs = Nothing
		oConn.Close
		SET oConn = Nothing

		Response.Write "FAIL|||||장바구니 삭제 처리중 오류가 발생하였습니다."
		Response.End
END IF





Response.Write "OK|||||"


Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>