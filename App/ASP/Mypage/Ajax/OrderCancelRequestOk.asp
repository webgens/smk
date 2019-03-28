<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'OrderCancelRequestOk.asp - 주문취소신청 처리
'Date		: 2019.01.02
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
DIM OPIdx
DIM ReasonType
DIM Memo
DIM RefundBankCode
DIM RefundBankName
DIM RefundAccountNum
DIM RefundAccountName
DIM RefundPhone1
DIM RefundPhone23
DIM RefundPhone

DIM OrderName
DIM PayType
DIM OrderCnt			: OrderCnt			= 0
DIM TotalOrderCnt		: TotalOrderCnt		= 0

DIM OrderPrice			: OrderPrice		= 0
DIM SalePrice			: SalePrice			= 0
DIM DeliveryPrice		: DeliveryPrice		= 0
DIM AddDeliveryPrice	: AddDeliveryPrice	= 0

DIM DiscountPrice		: DiscountPrice		= 0
DIM UseCouponPrice		: UseCouponPrice	= 0
DIM UsePointPrice		: UsePointPrice		= 0
DIM UseScashPrice		: UseScashPrice		= 0

DIM RefundPrice			: RefundPrice		= 0

DIM SetStandardPrice	: SetStandardPrice	= 0
DIM SetDeliveryPrice	: SetDeliveryPrice	= 0

DIM CancelType			: CancelType	= "Cancel"			'# Cancel : 전체취소, PartialCancel : 부분취소
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


OrderCode			= sqlFilter(Request("OrderCode"))
OPIdx				= sqlFilter(Request("OPIdx"))
ReasonType			= sqlFilter(Request("ReasonType"))
Memo				= sqlFilter(Request("Memo"))
RefundBankCode		= sqlFilter(Request("RefundBankCode"))
RefundAccountNum	= sqlFilter(Request("RefundAccountNum"))
RefundAccountName	= sqlFilter(Request("RefundAccountName"))
RefundPhone1		= sqlFilter(Request("RefundPhone1"))
RefundPhone23		= sqlFilter(Request("RefundPhone23"))
RefundPhone			= RefundPhone1 & RefundPhone23


IF OrderCode = "" OR OPIdx = "" THEN
		Response.Write "FAIL|||||취소신청할 입력정보가 부족합니다."
		Response.End
END IF



SET oConn	= ConnectionOpen()	'//커넥션 생성
SET oRs		= Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성
SET oRs1	= Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성


'-----------------------------------------------------------------------------------------------------------'
'# 주문정보 체크 Start
'-----------------------------------------------------------------------------------------------------------'
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Order_Select_For_OrderInfo"

		.Parameters.Append .CreateParameter("@OrderCode",	adVarchar, adParaminput,	20,		OrderCode)
		.Parameters.Append .CreateParameter("@UserID",		adVarchar, adParaminput,	20,		U_NUM)
		.Parameters.Append .CreateParameter("@OrderName",	adVarChar, adParamInput,	50,		N_NAME)
		.Parameters.Append .CreateParameter("@OrderHp",		adVarChar, adParamInput,	20,		N_HP)
		.Parameters.Append .CreateParameter("@OrderEmail",	adVarChar, adParamInput,	50,		N_EMAIL)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		'# 결제완료 상태가 아니면 취소불가
		IF oRs("SettleFlag") <> "Y" THEN
				oRs.Close : SET oRs1 = Nothing : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "FAIL|||||결제완료되지 않은 주문은 주문취소신청할 수 없습니다."
				Response.End
		END IF

		TotalOrderCnt	= oRs("OrderCnt")		'# 본상품 수량
		PayType			= oRs("PayType")
		OrderName		= oRs("OrderName")

ELSE
		oRs.Close : SET oRs1 = Nothing : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||취소신청할 주문내역이 없습니다.[1]"
		Response.End
END IF
oRs.Close
'-----------------------------------------------------------------------------------------------------------'
'# 주문정보 체크 End
'-----------------------------------------------------------------------------------------------------------'


'-----------------------------------------------------------------------------------------------------------'
'# 주문상품 상태 체크 Start
'-----------------------------------------------------------------------------------------------------------'
wQuery = ""
wQuery = wQuery & "WHERE A.IsShowFlag = 'Y' AND A.ProductType IN ('P','O') "
wQuery = wQuery & "AND A.OrderCode = '" & OrderCode & "' "
wQuery = wQuery & "AND A.OPIdx_Group IN (" & OPIdx & ") "

sQuery = "ORDER BY A.OPIdx_Group, A.OPIdx_Org"

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
		Do Until oRs.EOF
				'# 상품준비중 상태가 아니면 취소신청불가
				IF oRs("OrderState") = "4" AND oRs("CancelState1") = "0" AND oRs("CancelState2") = "0" THEN
						'# ERP 주문상태 체크
						SET oCmd = Server.CreateObject("ADODB.Command")
						WITH oCmd
								.ActiveConnection	 = oConn
								.CommandType		 = adCmdStoredProc
								.CommandText		 = "USP_Admin_ERP_ONLINE_ORDER_V_Select_By_ITEMNO"
		
								.Parameters.Append .CreateParameter("@LINKED_SERVER_NAME",	 adVarChar, adParamInput, 20, ERP_LNK_SRV)
								.Parameters.Append .CreateParameter("@TABLE_NAME",			 adVarChar, adParamInput, 50, ERP_OST_TBL)
								.Parameters.Append .CreateParameter("@ITEMNO",				 adVarChar, adParamInput, 40, oRs("Idx"))
						END WITH
						oRs1.CursorLocation = adUseClient
						oRs1.Open oCmd, , adOpenStatic, adLockReadOnly
						SET oCmd = Nothing
		
						IF NOT oRs1.EOF THEN
								'# STATUS : 1-주문, 2-출고준비, 3-출고/반품완료, 4-재고부족, 5-창고이동중, 8-재고부족, 9-취소
								IF oRs1("STATUS") = "3" THEN
										oRs1.Close : oRs.Close : SET oRs1 = Nothing : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
										Response.Write "FAIL|||||주문취소신청할 수 없는 상태의 상품 있습니다.[11]"
										Response.End
								END IF
						END IF
						oRs1.Close

						IF oRs("ProductType") = "P" THEN
								OrderCnt	= OrderCnt		+ oRs("OrderCnt")
						END IF
				ELSE
						oRs.Close : SET oRs1 = Nothing : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
						Response.Write "FAIL|||||주문취소신청할 수 없는 상태의 상품 있습니다.[11]"
						Response.End
				END IF

				oRs.MoveNext
		Loop
ELSE
		oRs.Close : SET oRs1 = Nothing : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||취소할 주문상품이 없습니다.[12]"
		Response.End
END IF
oRs.Close


IF OrderCnt < TotalOrderCnt THEN
		CancelType	= "PartialCancel"
END IF
'-----------------------------------------------------------------------------------------------------------'
'# 주문상품 상태 체크 End
'-----------------------------------------------------------------------------------------------------------'



'-----------------------------------------------------------------------------------------------------------'
'# 환불금액 계산 Start
'-----------------------------------------------------------------------------------------------------------'
wQuery = ""
wQuery = wQuery & "WHERE A.OrderCode = '" & OrderCode & "' "
wQuery = wQuery & "AND A.ProductType = 'P' "
wQuery = wQuery & "AND A.Idx IN (" & OPIdx & ") "


SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Order_Product_Select_For_Vendor_TotalPrice"

		.Parameters.Append .CreateParameter("@WQUERY",		adVarchar, adParaminput,	1000,	wQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		Do Until oRs.EOF
				SalePrice		= SalePrice			+ oRs("SalePrice")
				UseCouponPrice	= UseCouponPrice	+ oRs("UseCouponPrice")
				UsePointPrice	= UsePointPrice		+ oRs("UsePointPrice")
				UseScashPrice	= UseScashPrice		+ oRs("UseScashPrice")
				DiscountPrice	= DiscountPrice		+ UseCouponPrice + UsePointPrice + UseScashPrice

				'# 배송비 설정 가져오기
				SET oCmd = Server.CreateObject("ADODB.Command")
				WITH oCmd
						.ActiveConnection	 = oConn
						.CommandType		 = adCmdStoredProc
						.CommandText		 = "USP_Admin_EShop_Store_Select_By_ShopCD"

						.Parameters.Append .CreateParameter("@ShopCD",		adChar, adParaminput,	6,	oRs("Vendor"))
				END WITH
				oRs1.CursorLocation = adUseClient
				oRs1.Open oCmd, , adOpenStatic, adLockReadOnly
				SET oCmd = Nothing

				IF NOT oRs1.EOF THEN
						SetStandardPrice	= oRs1("StandardPrice")
						SetDeliveryPrice	= oRs1("DeliveryPrice")
				END IF
				oRs1.Close

				'-----------------------------------------------------------------------------------------------------------'
				'# 배송비 환불금액 계산 Start
				'-----------------------------------------------------------------------------------------------------------'
				wQuery = ""
				wQuery = wQuery & "WHERE A.OrderCode = '" & OrderCode & "' "
				wQuery = wQuery & "AND A.ProductType = 'P' "
				wQuery = wQuery & "AND A.DelvType = 'P' "
				wQuery = wQuery & "AND A.OrderState <> 'C' "
				wQuery = wQuery & "AND A.Vendor = '" & oRs("Vendor") & "' "
				wQuery = wQuery & "AND A.Idx NOT IN (" & OPIdx & ") "

				SET oCmd = Server.CreateObject("ADODB.Command")
				WITH oCmd
						.ActiveConnection	 = oConn
						.CommandType		 = adCmdStoredProc
						.CommandText		 = "USP_Front_EShop_Order_Product_Select_For_TotalPrice"

						.Parameters.Append .CreateParameter("@WQUERY",		adVarChar, adParaminput,	1000,	wQuery)
				END WITH
				oRs1.CursorLocation = adUseClient
				oRs1.Open oCmd, , adOpenStatic, adLockReadOnly
				SET oCmd = Nothing

				IF NOT oRs1.EOF THEN
						'# 남아있는 일반택배 주문상품이 없을 경우 결제한 배송비 환불
						IF CDbl(oRs1("OrderPrice")) = 0 THEN
								DeliveryPrice	= CDbl(DeliveryPrice) + CDbl(oRs("DeliveryPrice"))

						'# 남아있는 주문상품 금액이 배송비 기준금액 미만이면 결제금액에서 차감할 추가 배송비 적용
						ELSEIF CDbl(oRs1("OrderPrice")) < CDbl(SetStandardPrice) THEN
								'# 최초 주문시 무료배송이었을 경우 환불금액에서 차감할 배송비 적용
								IF CDbl(oRs("DeliveryPrice")) = 0 THEN
										AddDeliveryPrice	= CDbl(AddDeliveryPrice) + CDbl(SetDeliveryPrice)
								END IF
						END IF
				ELSE
						DeliveryPrice	= CDbl(DeliveryPrice) + CDbl(oRs("DeliveryPrice"))
				END IF
				oRs1.Close
				'-----------------------------------------------------------------------------------------------------------'
				'# 배송비 환불금액 계산 End
				'-----------------------------------------------------------------------------------------------------------'

				oRs.MoveNext
		Loop
END IF
oRs.Close

'# 환불금액 = 상품금액 + 배송비 - 쿠폰할인/포인트사용/슈즈상품권사용 - 추가배송비
RefundPrice		= SalePrice + DeliveryPrice - DiscountPrice - AddDeliveryPrice

IF RefundPrice < 0 THEN
		oConn.RollbackTrans

		Set oRs1 = Nothing : Set oRs = Nothing
		oConn.Close : Set oConn = Nothing

		Response.Write "FAIL|||||환불금액이 부족하여 취소신청하실 수 없습니다. 고객센터에 문의해 주십시오"
		Response.End
END IF
'-----------------------------------------------------------------------------------------------------------'
'# 환불금액 계산 End
'-----------------------------------------------------------------------------------------------------------'


'-----------------------------------------------------------------------------------------------------------'
'# 환불은행 체크 Start
'-----------------------------------------------------------------------------------------------------------'
IF RefundBankCode <> "" THEN
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Admin_EShop_RefundBank_Select_By_BankCode"

				.Parameters.Append .CreateParameter("@BankCode",	adChar, adParaminput,	2,		RefundBankCode)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		SET oCmd = Nothing

		IF NOT oRs.EOF THEN
				RefundBankName		= oRs("BankName")
		END IF
		oRs.Close
END IF
'-----------------------------------------------------------------------------------------------------------'
'# 환불은행 체크 End
'-----------------------------------------------------------------------------------------------------------'


ON ERROR RESUME NEXT



oConn.BeginTrans



'-----------------------------------------------------------------------------------------------------------'	
'주문취소신청 START
'-----------------------------------------------------------------------------------------------------------'	
OPIdx	= Split(OPIdx, ", ")

FOR i = 0 TO UBOUND(OPIdx)
		'-----------------------------------------------------------------------------------------------------------'	
		'주문상품정보 START
		'-----------------------------------------------------------------------------------------------------------'	
		Set oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection = oConn
				.CommandType = adCmdStoredProc
				.CommandText = "USP_Admin_EShop_Order_Product_Select_By_OPIdx_Group"
				.Parameters.Append .CreateParameter("@OPIdx_Group",		adInteger,	adParamInput,	   ,		OPIdx(i))
			
				.Execute, , adExecuteNoRecords
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		Set oCmd = Nothing

		IF NOT oRs.EOF THEN
				Do Until oRs.EOF
						RefundPrice			= oRs("OrderPrice")

						RefundPrice			= CDbl(RefundPrice) + CDbl(DeliveryPrice) - CDbl(AddDeliveryPrice)
						DeliveryPrice		= 0
						IF RefundPrice < 0 THEN
								AddDeliveryPrice	= AddDeliveryPrice + RefundPrice
								RefundPrice			= 0
						ELSE
								AddDeliveryPrice	= 0
						END IF

						'-----------------------------------------------------------------------------------------------------------'	
						'# 주문상태 변경 Start
						'-----------------------------------------------------------------------------------------------------------'	
						SET oCmd = Server.CreateObject("ADODB.Command")
						WITH oCmd
								.ActiveConnection	 = oConn
								.CommandType		 = adCmdStoredProc
								.CommandText		 = "USP_Admin_EShop_Order_Product_Update_For_CRX_OrderStates"

								.Parameters.Append .CreateParameter("@Idx",					adInteger,	adParamInput,   ,	 oRs("Idx"))
								.Parameters.Append .CreateParameter("@OrderState",			adChar,		adParamInput,  1,	 oRs("OrderState"))
								.Parameters.Append .CreateParameter("@CancelState1",		adChar,		adParamInput,  1,	 oRs("CancelState1"))
								.Parameters.Append .CreateParameter("@CancelState2",		adChar,		adParamInput,  1,	 "R")
								.Parameters.Append .CreateParameter("@ReturnStockCDGB",		adVarChar,	adParamInput,  2,	 "")
								.Parameters.Append .CreateParameter("@UpdateID",			adVarChar,	adParamInput, 20,	 U_NUM)
								.Parameters.Append .CreateParameter("@UpdateIP",			adVarChar,	adParamInput, 15,	 U_IP)

								.Execute, , adExecuteNoRecords
						END WITH
						SET oCmd = Nothing

						IF Err.Number <> 0 THEN
								oConn.RollbackTrans

								oRs.Close
								SET oRs1 = Nothing
								SET oRs = Nothing
								oConn.Close
								SET oConn = Nothing

								Response.Write "FAIL|||||주문 상태변경 처리중 오류가 발생하였습니다."
								Response.End
						END IF
						'-----------------------------------------------------------------------------------------------------------'	
						'# 주문상태 변경 End
						'-----------------------------------------------------------------------------------------------------------'	


						'-----------------------------------------------------------------------------------------------------------'	
						'# 주문변경 이력 생성 Start
						'-----------------------------------------------------------------------------------------------------------'	
						SET oCmd = Server.CreateObject("ADODB.Command")
						WITH oCmd
								.ActiveConnection	 = oConn
								.CommandType		 = adCmdStoredProc
								.CommandText		 = "USP_Admin_EShop_Order_Product_Change_History_Insert"

								.Parameters.Append .CreateParameter("@OPIdx",		 adInteger,	 adParamInput,     ,	 oRs("Idx"))
								.Parameters.Append .CreateParameter("@Contents",	 adVarChar,	 adParamInput, 8000,	 "주문취소신청")
								.Parameters.Append .CreateParameter("@CreateNM",	 adVarChar,	 adParamInput,  100,	 OrderName)
								.Parameters.Append .CreateParameter("@CreateID",	 adVarChar,	 adParamInput,   50,	 U_NUM)
								.Parameters.Append .CreateParameter("@CreateIP",	 adVarChar,	 adParamInput,   20,	 U_IP)

								.Execute, , adExecuteNoRecords
						END WITH
						SET oCmd = Nothing

						IF Err.Number <> 0 THEN
								oConn.RollbackTrans

								oRs.Close
								SET oRs1 = Nothing
								SET oRs = Nothing
								oConn.Close
								SET oConn = Nothing

								Response.Write "FAIL|||||주문 변경이력 생성 처리중 오류가 발생하였습니다."
								Response.End
						END IF
						'-----------------------------------------------------------------------------------------------------------'	
						'# 주문변경 이력 생성 End
						'-----------------------------------------------------------------------------------------------------------'	


						'-----------------------------------------------------------------------------------------------------------'	
						'# 주문취소요청 내역 생성 Start
						'-----------------------------------------------------------------------------------------------------------'	
						SET oCmd = Server.CreateObject("ADODB.Command")
						WITH oCmd
								.ActiveConnection	 = oConn
								.CommandType		 = adCmdStoredProc
								.CommandText		 = "USP_Admin_EShop_Order_Product_Cancel_Insert"

								.Parameters.Append .CreateParameter("@OPIdx",			adInteger,	adParamInput,     ,	 oRs("Idx"))
								.Parameters.Append .CreateParameter("@CancelType",		adChar,		adParamInput,    1,	 "C")
								.Parameters.Append .CreateParameter("@OrderCode",		adVarChar,	adParamInput,   20,	 oRs("OrderCode"))
								.Parameters.Append .CreateParameter("@ProductCode",		adInteger,	adParamInput,     ,	 oRs("ProductCode"))
								.Parameters.Append .CreateParameter("@SizeChangeFlag",	adChar,		adParamInput,    1,	 "N")
								.Parameters.Append .CreateParameter("@SizeCD",			adVarChar,	adParamInput,   10,	 oRs("SizeCD"))
								.Parameters.Append .CreateParameter("@OrderCnt",		adInteger,	adParamInput,     ,	 oRs("OrderCnt"))
								.Parameters.Append .CreateParameter("@ReasonType",		adVarChar,	adParamInput,   10,	 ReasonType)
								.Parameters.Append .CreateParameter("@Memo",			adVarChar,	adParamInput,  255,	 Memo)
								.Parameters.Append .CreateParameter("@DelvFee",			adCurrency,	adParamInput,     ,	 0)
								.Parameters.Append .CreateParameter("@DelvFeeType",		adChar,		adParamInput,    1,	 "")
								.Parameters.Append .CreateParameter("@DelvFeeMemo",		adVarChar,	adParamInput,  255,	 "")
								.Parameters.Append .CreateParameter("@ContactName",		adVarChar,	adParamInput,   50,	 OrderName)
								.Parameters.Append .CreateParameter("@ContactHp",		adVarChar,	adParamInput,   50,	 RefundPhone)
								.Parameters.Append .CreateParameter("@RefundPrice",		adCurrency,	adParamInput,     ,	 RefundPrice)
								.Parameters.Append .CreateParameter("@DepositBankCode",	adVarChar,	adParamInput,   10,	 RefundBankCode)
								.Parameters.Append .CreateParameter("@DepositBankName",	adVarChar,	adParamInput,   50,	 RefundBankName)
								.Parameters.Append .CreateParameter("@DepositNumber",	adVarChar,	adParamInput,   50,	 RefundAccountNum)
								.Parameters.Append .CreateParameter("@DepositName",		adVarChar,	adParamInput,   50,	 RefundAccountName)
								.Parameters.Append .CreateParameter("@CreateID",		adVarChar,	adParamInput,   50,	 U_NUM)
								.Parameters.Append .CreateParameter("@CreateIP",		adVarChar,	adParamInput,   15,	 U_IP)

								.Execute, , adExecuteNoRecords
						END WITH
						SET oCmd = Nothing

						IF Err.Number <> 0 THEN
								oConn.RollbackTrans

								oRs.Close
								SET oRs1 = Nothing
								SET oRs = Nothing
								oConn.Close
								SET oConn = Nothing

								Response.Write "FAIL|||||주문취소 요청정보 생성 처리중 오류가 발생하였습니다."
								Response.End
						END IF
						'-----------------------------------------------------------------------------------------------------------'	
						'# 주문취소요청 내역 생성 End
						'-----------------------------------------------------------------------------------------------------------'	

						oRs.MoveNext
				Loop
		ELSE
				oRs.Close : SET oRs1 = Nothing : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "FAIL|||||취소신청한 주문상품이 없습니다[21]"
				Response.End
		END IF
		oRs.Close
		'-----------------------------------------------------------------------------------------------------------'	
		'주문상품정보 End
		'-----------------------------------------------------------------------------------------------------------'	
NEXT
'-----------------------------------------------------------------------------------------------------------'	
'주문취소신청 End
'-----------------------------------------------------------------------------------------------------------'	


oConn.CommitTrans



Response.Write "OK|||||"

SET oRs1 = Nothing
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>