<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'OrderConfirmOk.asp - 주문 구매확정 처리
'Date		: 2019.01.03
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
IF U_NUM = "" AND N_NAME = "" THEN
		Response.Write "LOGIN|||||로그인 정보가 없습니다. 다시 로그인하여 주십시오."
		Response.End
END IF

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

DIM OrderCode
DIM Idx

DIM EscrowFlag
DIM DelvType
DIM OrderStateNM
DIM ProductPoint	: ProductPoint	= 0
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


OrderCode			= sqlFilter(Request("OrderCode"))
Idx					= sqlFilter(Request("Idx"))


IF OrderCode = "" OR Idx = "" THEN
		Response.Write "FAIL|||||구매확정할 입력정보가 부족합니다."
		Response.End
END IF



SET oConn	= ConnectionOpen()	'//커넥션 생성
SET oRs		= Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성




'-----------------------------------------------------------------------------------------------------------'
'# 주문상품 상태 체크 Start
'-----------------------------------------------------------------------------------------------------------'
wQuery = ""
wQuery = wQuery & "WHERE A.IsShowFlag = 'Y' AND A.ProductType = 'P' "
wQuery = wQuery & "AND A.OrderCode = '" & OrderCode & "' "
wQuery = wQuery & "AND A.Idx = " & Idx & " "

sQuery = ""

SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Order_Product_Select_For_Order_Detail"

		.Parameters.Append .CreateParameter("@WQuery",		 adVarchar, adParaminput, 1000	, wQuery)
		.Parameters.Append .CreateParameter("@SQuery",		 adVarchar, adParaminput, 100	, sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		EscrowFlag		= oRs("EscrowFlag")
		DelvType		= oRs("DelvType")
		OrderStateNM	= GetOrderState(oRs("OrderState"), oRs("CancelState1"), oRs("CancelState2"))

		IF InStr(OrderStateNM, "배송중") <= 0 AND InStr(OrderStateNM, "배송완료") <= 0 THEN
				oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "FAIL|||||구매확정할 수 없는 주문상태 입니다."
				Response.End
		END IF

		ProductPoint	= oRs("ProductPoint")
ELSE
		oRs.Close : SET oRs1 = Nothing : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||구매확정할 주문상품이 없습니다.[1]"
		Response.End
END IF
oRs.Close
'-----------------------------------------------------------------------------------------------------------'
'# 주문상품 상태 체크 End
'-----------------------------------------------------------------------------------------------------------'


ON ERROR RESUME NEXT


oConn.BeginTrans



'-----------------------------------------------------------------------------------------------------------'	
'# 주문상태 변경 Start
' 1. 주문상품 상태변경
' 2. 주문상품 변경이력 생성
' 3. 포인트 적립
'-----------------------------------------------------------------------------------------------------------'	
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_Order_Product_Update_For_OrderConfirm"

		.Parameters.Append .CreateParameter("@OrderCode",			adVarChar,	adParamInput, 20,	 OrderCode)
		.Parameters.Append .CreateParameter("@OPIdx",				adInteger,	adParamInput,   ,	 Idx)
		.Parameters.Append .CreateParameter("@UpdateNM",			adVarChar,	adParamInput, 50,	 U_NAME)
		.Parameters.Append .CreateParameter("@UpdateID",			adVarChar,	adParamInput, 20,	 U_NUM)
		.Parameters.Append .CreateParameter("@UpdateIP",			adVarChar,	adParamInput, 15,	 U_IP)

		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing

IF Err.Number <> 0 THEN
		oConn.RollbackTrans

		oRs.Close
		SET oRs = Nothing
		oConn.Close
		SET oConn = Nothing

		Response.Write "FAIL|||||주문 구매확정 처리중 오류가 발생하였습니다.[2]"
		Response.End
END IF
'-----------------------------------------------------------------------------------------------------------'	
'# 주문상태 변경 End
'-----------------------------------------------------------------------------------------------------------'	



oConn.CommitTrans


'-----------------------------------------------------------------------------------------------------------'	
'문자발송 시작
'-----------------------------------------------------------------------------------------------------------'	
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_Order_Sms_Send"

		.Parameters.Append .CreateParameter("@OrderCode",	 adVarChar,	 adParamInput,   20,	 OrderCode)
		.Parameters.Append .CreateParameter("@OPIdx",		 adInteger,	 adParamInput,     ,	 Idx)
		.Parameters.Append .CreateParameter("@SmsCode",		 adVarChar,	 adParamInput,   20,	 "ORD_S700")

		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing
'-----------------------------------------------------------------------------------------------------------'	
'문자발송 끝
'-----------------------------------------------------------------------------------------------------------'	


Response.Write "OK|||||"

Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>