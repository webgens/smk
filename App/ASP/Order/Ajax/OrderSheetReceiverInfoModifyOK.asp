<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'OrderSheetReceiverInfoModifyOk.asp - 주문서 배송지정보 등록 처리 페이지
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
DIM AddressName
DIM ReceiveName
DIM ReceiveTel
DIM ReceiveTel1
DIM ReceiveTel23
DIM ReceiveHP
DIM ReceiveHP1
DIM ReceiveHP23
DIM ReceiveZipCode
DIM ReceiveAddr1
DIM ReceiveAddr2
DIM MainFlag
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


OrderSheetIdx		= sqlFilter(Request("OrderSheetIdx"))
AddressName			= sqlFilter(Request("AddressName"))
ReceiveName			= sqlFilter(Request("ReceiveName"))
ReceiveTel1			= sqlFilter(Request("ReceiveTel1"))
ReceiveTel23		= sqlFilter(Request("ReceiveTel23"))
ReceiveHP1			= sqlFilter(Request("ReceiveHP1"))
ReceiveHP23			= sqlFilter(Request("ReceiveHP23"))
ReceiveZipCode		= sqlFilter(Request("ReceiveZipCode"))
ReceiveAddr1		= sqlFilter(Request("ReceiveAddr1"))
ReceiveAddr2		= sqlFilter(Request("ReceiveAddr2"))
MainFlag			= sqlFilter(Request("MainFlag"))


IF AddressName	= ""	THEN AddressName	= ReceiveName
IF MainFlag		<> "Y"	THEN MainFlag		= "N"

ReceiveTel			= ChgTel(ReceiveTel1 & ReceiveTel23)
ReceiveHP			= ChgTel(ReceiveHP1 & ReceiveHP23)

IF OrderSheetIdx = "" THEN
		Response.Write "FAIL|||||입력정보가 부족합니다."
		Response.End
END IF


SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성







'# 배송지 적용
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_OrderSheet_Update_For_ReceiverInfo"

		.Parameters.Append .CreateParameter("@Idx",					adInteger,	adParamInput,     ,		OrderSheetIdx)
		.Parameters.Append .CreateParameter("@AddressName",			adVarChar,	adParamInput,  100,		AddressName)
		.Parameters.Append .CreateParameter("@ReceiveName",			adVarChar,	adParamInput,   50,		ReceiveName)
		.Parameters.Append .CreateParameter("@ReceiveTel",			adVarChar,	adParamInput,   20,		ReceiveTel)
		.Parameters.Append .CreateParameter("@ReceiveHP",			adVarChar,	adParamInput,   20,		ReceiveHP)
		.Parameters.Append .CreateParameter("@ReceiveZipCode",		adVarChar,	adParamInput,    7,		ReceiveZipCode)
		.Parameters.Append .CreateParameter("@ReceiveAddr1",		adVarChar,	adParamInput,  200,		ReceiveAddr1)
		.Parameters.Append .CreateParameter("@ReceiveAddr2",		adVarChar,	adParamInput,  200,		ReceiveAddr2)
		.Parameters.Append .CreateParameter("@MainFlag",			adChar,		adParamInput,	 1,		MainFlag)
		.Parameters.Append .CreateParameter("@UpdateID",			adVarChar,	adParamInput,   20,		U_NUM)
		.Parameters.Append .CreateParameter("@UpdateIP",			adVarChar,	adParamInput,   15,		U_IP)

		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing

IF Err.Number <> 0 THEN
		SET oRs = Nothing
		oConn.Close
		SET oConn = Nothing

		Response.Write "FAIL|||||배송지 적용 처리중 오류가 발생하였습니다."
		Response.End
END IF







Response.Write "OK|||||"


Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>