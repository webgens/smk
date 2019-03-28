<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'CartCount.asp - GNB 장바구니 상품갯수 가져오기 페이지
'Date		: 2018.11.19
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

DIM CartCount		: CartCount = 0
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성




Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
	.ActiveConnection = oConn
	.CommandType = adCmdStoredProc
	.CommandText = "USP_Front_EShop_Cart_Select_For_CartCount_By_CartID"
	.Parameters.Append .CreateParameter("@CartID",	 adVarChar,	 adParamInput, 20,		 U_CARTID)
	.Parameters.Append .CreateParameter("@UserID",	 adVarChar,	 adParamInput, 20,		 U_NUM)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing
																
IF NOT oRs.EOF THEN
		CartCount	= oRs("CartCount")
END IF
oRs.Close


Response.Write CartCount
%>
<%
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>
