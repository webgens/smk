<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'OrderSheetPickupStoreModifyOk.asp - 주문서 픽업매장정보 등록 처리 페이지
'Date		: 2018.12.30
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


DIM OrderSheetIdx
DIM PickupShopCD
DIM ReceiveName
DIM ReceiveHP
DIM ReceiveHP1
DIM ReceiveHP23
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


OrderSheetIdx		= sqlFilter(Request("OrderSheetIdx"))
PickupShopCD		= sqlFilter(Request("StoreCode"))
ReceiveName			= sqlFilter(Request("PickupReceiveName"))
ReceiveHP1			= sqlFilter(Request("PickupReceiveHP1"))
ReceiveHP23			= sqlFilter(Request("PickupReceiveHP23"))


ReceiveHP			= ChgTel(ReceiveHP1 & ReceiveHP23)

IF OrderSheetIdx = "" OR PickupShopCD = "" THEN
		Response.Write "FAIL|||||입력정보가 부족합니다."
		Response.End
END IF


SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성







'# 픽업매장 적용
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_OrderSheet_Update_For_PickupShopCD"

		.Parameters.Append .CreateParameter("@Idx",					adInteger,	adParamInput,     ,		OrderSheetIdx)
		.Parameters.Append .CreateParameter("@PickupShopCD",		adVarChar,	adParamInput,   10,		PickupShopCD)
		.Parameters.Append .CreateParameter("@ReceiveName",			adVarChar,	adParamInput,   50,		ReceiveName)
		.Parameters.Append .CreateParameter("@ReceiveTel",			adVarChar,	adParamInput,   20,		"")
		.Parameters.Append .CreateParameter("@ReceiveHP",			adVarChar,	adParamInput,   20,		ReceiveHP)
		.Parameters.Append .CreateParameter("@UpdateID",			adVarChar,	adParamInput,   20,		U_NUM)
		.Parameters.Append .CreateParameter("@UpdateIP",			adVarChar,	adParamInput,   15,		U_IP)

		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing

IF Err.Number <> 0 THEN
		SET oRs = Nothing
		oConn.Close
		SET oConn = Nothing

		Response.Write "FAIL|||||픽업매장 적용 처리중 오류가 발생하였습니다."
		Response.End
END IF







Response.Write "OK|||||"


Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>