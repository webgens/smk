<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'OrderSheet_ProductAddOk.asp - 주문서 바로구매 등록 페이지
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
DIM oRs1											'# ADODB Recordset 개체
DIM oCmd											'# ADODB Command 개체

DIM i
DIM j
DiM x
DIM y

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절

DIM OrderType
DIM DelvType
DIM PickupShopCD

DIM ProductCode
DIM SizeCD
DIM OrderCnt
DIM SalePriceType
DIM ProductType

DIM ProductName
DIM NewIdx
DIM GroupIdx
DIM StockCnt
DIM CartOrderCnt
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


OrderType		= sqlFilter(Request("OrderType"))
DelvType		= sqlFilter(Request("DelvType"))
PickupShopCD	= sqlFilter(Request("PickupShopCD"))

ProductCode		= sqlFilter(Request("ProductCode"))
SizeCD			= sqlFilter(Request("SizeCD"))
OrderCnt		= sqlFilter(Request("OrderCnt"))
SalePriceType	= sqlFilter(Request("SalePriceType"))
ProductType		= sqlFilter(Request("ProductType"))


IF OrderType	= "" THEN OrderType		= "G"			'# 주문형태 - G:일반주문, R:예약주문
IF DelvType		= "" THEN DelvType		= "P"			'# 배송형태 - P:일반택배, S:매장픽업


IF ProductCode = "" OR SizeCD = "" OR OrderCnt = "" OR SalePriceType = "" OR ProductType = "" THEN
		Response.Write "FAIL|||||상품을 선택해 주십시오."
		Response.End
END IF



SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성




ProductCode		= Split(ProductCode, ",")
SizeCD			= Split(SizeCD, ",")
OrderCnt		= Split(OrderCnt, ",")
SalePriceType	= Split(SalePriceType, ",")
ProductType		= Split(ProductType, ",")


'# Response.Write "FAIL|||||"
'# FOR i = 0 TO UBOUND(ProductCode)
'# Response.Write ProductCode(i) & "|"
'# NEXT
'# Response.End

oConn.BeginTrans


'# 주문서 삭제
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_OrderSheet_Delete_By_CartID"

		.Parameters.Append .CreateParameter("@CartID",				adVarChar,	adParamInput,  20,	 U_CARTID)

		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing

IF Err.Number <> 0 THEN
		oConn.RollbackTrans

		SET oRs = Nothing
		oConn.Close
		SET oConn = Nothing

		Response.Write "FAIL|||||주문서 담기 처리중 오류가 발생하였습니다.[1]"
		Response.End
END IF



FOR i = 0 TO UBOUND(ProductCode)

		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Product_Select_By_ProductCode"

				.Parameters.Append .CreateParameter("@ProductCode", adInteger, adParamInput, , ProductCode(i))
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		SET oCmd = Nothing

		IF NOT oRs.EOF THEN
				ProductName		= oRs("ProductName")

				IF oRs("ReserveFlag") = "Y" THEN
						IF OrderType <> "R" THEN
								oConn.RollbackTrans
								oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
								Response.Write "FAIL|||||" & ProductName & "는 예약구매만 가능한 상품 입니다."
								Response.End
						END IF
				ELSE
						IF OrderType = "R" THEN
								oConn.RollbackTrans
								oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
								Response.Write "FAIL|||||" & ProductName & "는 예약구매를 할 수 없는 상품 입니다."
								Response.End
						END IF
				END IF

				IF DelvType = "S" THEN
						IF oRs("PickupFlag") <> "Y" THEN
								oConn.RollbackTrans
								oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
								Response.Write "FAIL|||||" & ProductName & "는 매장픽업을 할 수 없는 상품 입니다."
								Response.End
						END IF
				END IF
		ELSE
				oConn.RollbackTrans
				oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "FAIL|||||" & ProductName & "는 없는 상품 입니다."
				Response.End
		END IF
		oRs.Close


		'# 재고 체크
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Product_Select_For_Available_By_ProductCode_N_SizeCD"

				.Parameters.Append .CreateParameter("@ProductCode", adInteger, adParamInput,   , ProductCode(i))
				.Parameters.Append .CreateParameter("@SizeCD",		adVarChar, adParamInput, 50, SizeCD(i))
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		SET oCmd = Nothing

		IF NOT oRs.EOF THEN
				StockCnt	= CInt(oRs("StockCnt"))

				IF StockCnt < 1 THEN
						oConn.RollbackTrans

						oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
						Response.Write "FAIL|||||" & ProductName & "[" & SizeCD(i) & "] 상품은 품절된 상품 입니다."
						Response.End

				ELSEIF StockCnt < CInt(OrderCnt(i)) THEN
						oConn.RollbackTrans

						oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
						Response.Write "FAIL|||||" & ProductName & "[" & SizeCD(i) & "] 상품은 선택된 수량보다 재고가 부족합니다."
						Response.End
				END IF

		ELSE
				oConn.RollbackTrans

				oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "FAIL|||||주문할 수 없는 상품 입니다."
				Response.End
		END IF
		oRs.Close


		'# 매장픽업일 경우 픽업매장의 재고 체크
		IF DelvType = "S" AND PickupShopCD <> "" THEN
				SET oCmd = Server.CreateObject("ADODB.Command")
				WITH oCmd
						.ActiveConnection	 = oConn
						.CommandType		 = adCmdStoredProc
						.CommandText		 = "USP_Admin_EShop_Stock_Select_By_Key"

						.Parameters.Append .CreateParameter("@ProductCode", adInteger, adParamInput,   , ProductCode(i))
						.Parameters.Append .CreateParameter("@SizeCD",		adVarChar, adParamInput, 50, SizeCD(i))
						.Parameters.Append .CreateParameter("@ShopCD",		adVarChar, adParamInput, 10, PickupShopCD)
				END WITH
				oRs.CursorLocation = adUseClient
				oRs.Open oCmd, , adOpenStatic, adLockReadOnly
				SET oCmd = Nothing

				IF NOT oRs.EOF THEN
						StockCnt	= CInt(oRs("RestQty"))

						IF StockCnt < 1 THEN
								oConn.RollbackTrans

								oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
								Response.Write "FAIL|||||" & ProductName & "[" & SizeCD(i) & "] 상품은 픽업매장에는 품절된 상품 입니다."
								Response.End

						ELSEIF StockCnt < CInt(OrderCnt(i)) THEN
								oConn.RollbackTrans

								oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
								Response.Write "FAIL|||||" & ProductName & "[" & SizeCD(i) & "] 상품은 픽업매장에는 선택된 수량보다 재고가 부족합니다."
								Response.End
						END IF

				ELSE
						oConn.RollbackTrans

						oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
						Response.Write "FAIL|||||해당 매장에는 픽업할 수 없는 상품 입니다."
						Response.End
				END IF
				oRs.Close
		END IF


		'# 주문수량을 1개씩 나눠서 주문서에 넣는다.
		FOR j = 1 TO CInt(OrderCnt(i))
				'# 1+1상품일 경우 원상품 장바구니 일련번호를 셋팅
				IF ProductType(i) = "O" THEN
						GroupIdx = NewIdx
				ELSE
						GroupIdx = 0
				END IF

				'# 주문서 담기
				SET oCmd = Server.CreateObject("ADODB.Command")
				WITH oCmd
						.ActiveConnection	 = oConn
						.CommandType		 = adCmdStoredProc
						.CommandText		 = "USP_Front_EShop_OrderSheet_Insert"

						.Parameters.Append .CreateParameter("@CartID",				adVarChar,	adParamInput,  20,	 U_CARTID)
						.Parameters.Append .CreateParameter("@ProductCode",			adInteger,	adParamInput,    ,	 ProductCode(i))
						.Parameters.Append .CreateParameter("@SizeCD",				adVarChar,	adParamInput,  50,	 SizeCD(i))
						.Parameters.Append .CreateParameter("@OrderCnt",			adInteger,	adParamInput,    ,	 OrderCnt(i))
						.Parameters.Append .CreateParameter("@GroupIdx",			adInteger,	adParamInput,    ,	 GroupIdx)
						.Parameters.Append .CreateParameter("@SalePriceType",		adChar,		adParamInput,   1,	 SalePriceType(i))
						.Parameters.Append .CreateParameter("@OrderType",			adChar,		adParamInput,   1,	 OrderType)
						.Parameters.Append .CreateParameter("@DelvType",			adChar,		adParamInput,   1,	 DelvType)
						.Parameters.Append .CreateParameter("@Location",			adChar,		adParamInput,   1,	 "M")
						.Parameters.Append .CreateParameter("@CreateID",			adVarChar,	adParamInput,  20,	 U_NUM)
						.Parameters.Append .CreateParameter("@CreateIP",			adVarChar,	adParamInput,  15,	 U_IP)
						.Parameters.Append .CreateParameter("@NewIdx",				adInteger,	adParamOutput)

						.Execute, , adExecuteNoRecords

						NewIdx = .Parameters("@NewIdx").Value
				END WITH
				SET oCmd = Nothing

				IF Err.Number <> 0 THEN
						oConn.RollbackTrans

						oRs.Close
						SET oRs = Nothing
						oConn.Close
						SET oConn = Nothing

						Response.Write "FAIL|||||주문서 담기 처리중 오류가 발생하였습니다.[2]"
						Response.End
				END IF
		NEXT
NEXT


oConn.CommitTrans


Response.Write "OK|||||"


Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>