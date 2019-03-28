<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'OrderSheetUseScashModifyOk.asp - 주문서 슈즈상품권 사용 처리 페이지
'Date		: 2018.12.28
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
<!-- #include Virtual = "/Common/CheckID_Ajax.asp" -->

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
DIM UseScashPrice

DIM SalePrice
DIM OrderPrice
DIM DB_UseCouponPrice
DIM DB_UsePointPrice
DIM DB_UseScashPrice

DIM TotalScash				: TotalScash			= 0
DIM TotalUseScashPrice		: TotalUseScashPrice	= 0
DIM UsableScashPrice
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


OrderSheetIdx		= sqlFilter(Request("OrderSheetIdx"))
UseScashPrice		= sqlFilter(Request("UseScashPrice"))


IF OrderSheetIdx = "" OR UseScashPrice = "" THEN
		Response.Write "FAIL|||||입력정보가 부족합니다."
		Response.End
END IF


SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성



'# 주문서 상품 내역
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_OrderSheet_Select_By_Idx"

		.Parameters.Append .CreateParameter("@CartID",	adVarChar,	adParamInput, 20,		U_CARTID)
		.Parameters.Append .CreateParameter("@Idx",		adInteger,	adParamInput,   ,		OrderSheetIdx)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		IF oRs("SalePriceType") = "2" THEN
				oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "FAIL|||||임직원 판매가 구매시 슈즈상품권을 사용하실 수 없습니다."
				Response.End
		END IF

		IF oRs("SalePriceType") = "2" THEN
				SalePrice			= oRs("EmployeeSalePrice")
		ELSE
				SalePrice			= oRs("SalePrice")
		END IF
		DB_UseCouponPrice	= oRs("UseCouponPrice")
		DB_UsePointPrice	= oRs("UsePointPrice")
		DB_UseScashPrice	= oRs("UseScashPrice")
ELSE
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||선택한 주문상품이 없습니다."
		Response.End
END IF
oRs.Close


'# 보유 슈즈상품권
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_Member_Select_By_MemberNum"

		.Parameters.Append .CreateParameter("@MemberNum",	adInteger,	adParamInput, ,		U_NUM)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		TotalScash		= oRs("Scash")
END IF
oRs.Close


'# 주문서에서 사용한 총 슈즈상품권
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_OrderSheet_Select_For_UseDiscount_By_CartID"

		.Parameters.Append .CreateParameter("@CartID",	adVarChar,	adParamInput, 20,		U_CARTID)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		TotalUseScashPrice	= oRs("TotalUseScashPrice")
END IF
oRs.Close

'# 사용가능 슈즈상품권
UsableScashPrice	= CDbl(TotalScash) - CDbl(TotalUseScashPrice) + CDbl(DB_UseScashPrice)

'# 상품별 결제 최소금액
DIM MinOrderPrice
MinOrderPrice		= CDbl(SalePrice) - CDbl(DB_UseCouponPrice) - CDbl(DB_UsePointPrice) - CDbl(MALL_MIN_ORDERPRICE)
IF UsableScashPrice > MinOrderPrice THEN
		UsableScashPrice	= MinOrderPrice
END IF

IF CDbl(UseScashPrice) > UsableScashPrice THEN
		SET oRs = Nothing
		oConn.Close
		SET oConn = Nothing

		Response.Write "FAIL|||||적용가능한 슈즈상품권보다 많이 사용하실 수 없습니다."
		Response.End
END IF


'# 주문금액 계산(1000원다 적으면 안됨)
OrderPrice	= CDbl(SalePrice) - CDbl(DB_UseCouponPrice) - CDbl(DB_UsePointPrice) - CDbl(UseScashPrice)
IF OrderPrice < CDbl(MALL_MIN_ORDERPRICE) THEN
		SET oRs = Nothing
		oConn.Close
		SET oConn = Nothing

		Response.Write "FAIL|||||상품별 결제금액이 " & FormatNumber(MALL_MIN_ORDERPRICE,0) & "원 이상이 되어야 합니다."
		Response.End
END IF



'# 슈즈상품권 적용
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_OrderSheet_Update_For_UseScashPrice"

		.Parameters.Append .CreateParameter("@Idx",				adInteger,	adParamInput,    ,	 OrderSheetIdx)
		.Parameters.Append .CreateParameter("@UseScashPrice",	adCurrency,	adParamInput,    ,	 UseScashPrice)
		.Parameters.Append .CreateParameter("@UpdateID",		adVarChar,	adParamInput,  20,	 U_NUM)
		.Parameters.Append .CreateParameter("@UpdateIP",		adVarChar,	adParamInput,  15,	 U_IP)

		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing

IF Err.Number <> 0 THEN
		SET oRs = Nothing
		oConn.Close
		SET oConn = Nothing

		Response.Write "FAIL|||||슈즈상품권 적용 처리중 오류가 발생하였습니다."
		Response.End
END IF





Response.Write "OK|||||"


Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>