<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'OrderCancelOk.asp - 결제완료건 주문취소 처리
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
<!-- #include virtual = "/API/json_for_asp/aspJSON1.17.asp" -->

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
DIM RefundBankCode
DIM RefundAccountNum
DIM RefundAccountName
DIM RefundPhone1
DIM RefundPhone23
DIM RefundPhone

DIM PayType
DIM EscrowFlag
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
DIM LGD_MID
DIM LGD_TID
DIM LGD_OID
DIM LGD_CANCELREASON
DIM LGD_CANCELREQUESTER
DIM LGD_CANCELREQUESTERIP
DIM LGD_RESPCODE
DIM LGD_RESPMSG
DIM LGD_TIMESTAMP
DIM LGD_PAYTYPE
DIM LGD_RFBANKCODE
DIM LGD_RFACCOUNTNUM
DIM LGD_RFCUSTOMERNAME
DIM LGD_RFPHONE

DIM Result
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


OrderCode			= sqlFilter(Request("OrderCode"))
OPIdx				= sqlFilter(Request("OPIdx"))
RefundBankCode		= sqlFilter(Request("RefundBankCode"))
RefundAccountNum	= sqlFilter(Request("RefundAccountNum"))
RefundAccountName	= sqlFilter(Request("RefundAccountName"))
RefundPhone1		= sqlFilter(Request("RefundPhone1"))
RefundPhone23		= sqlFilter(Request("RefundPhone23"))
RefundPhone			= RefundPhone1 & RefundPhone23


IF OrderCode = "" OR OPIdx = "" THEN
		Response.Write "FAIL|||||취소할 입력정보가 부족합니다."
		Response.End
END IF


'# 결제 오류시 로그 데이터
SUB SettleErrorLogWrite(ByVal orderCode, ByVal cancelFlag, ByVal errCode, ByVal errPage, ByVal errMsg, ByVal errDesc)

		ON ERROR RESUME NEXT

		DIM oErrConn
		DIM oErrCmd

		SET oErrConn	 = ConnectionOpen()

		SET oErrCmd = Server.CreateObject("ADODB.Command")
		WITH oErrCmd
				.ActiveConnection	 = oErrConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Order_Settle_Error_Log_Insert"

				.Parameters.Append .CreateParameter("@OrderCode",	 adInteger,	 adParamInput,     ,	 orderCode)
				.Parameters.Append .CreateParameter("@Location",	 adChar,	 adParamInput,    1,	 "W")
				.Parameters.Append .CreateParameter("@CancelFlag",	 adChar,	 adParamInput,    1,	 cancelFlag)
				.Parameters.Append .CreateParameter("@ErrCode",		 adChar,	 adParamInput,    4,	 errCode)
				.Parameters.Append .CreateParameter("@ErrPage",		 adVarChar,	 adParamInput,   20,	 errPage)
				.Parameters.Append .CreateParameter("@ErrMsg",		 adVarChar,	 adParamInput,  100,	 errMsg)
				.Parameters.Append .CreateParameter("@ErrDesc",		 adVarChar,	 adParamInput, 3000,	 errDesc)
				.Parameters.Append .CreateParameter("@CreateID",	 adVarChar,	 adParamInput,   20,	 U_NUM)
				.Parameters.Append .CreateParameter("@CreateIP",	 adVarChar,	 adParamInput,   15,	 U_IP)

				.Execute, , adExecuteNoRecords
		END WITH
		SET oErrCmd = Nothing

		oErrConn.Close
		SET oErrConn = Nothing
END SUB


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
				Response.Write "FAIL|||||결제완료되지 않은 주문은 주문취소할 수 없습니다."
				Response.End
		END IF

		TotalOrderCnt	= oRs("OrderCnt")		'# 본상품 수량
		PayType			= oRs("PayType")
		EscrowFlag		= oRs("EscrowFlag")
		LGD_TID			= oRs("LGD_TID")

ELSE
		oRs.Close : SET oRs1 = Nothing : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||취소할 주문내역이 없습니다.[1]"
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
				'# 결제완료 상태가 아니면 취소불가
				IF oRs("OrderState") = "3" AND oRs("CancelState1") = "0" AND oRs("CancelState2") = "0" THEN
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
								IF oRs1("STATUS") <> "1" THEN
										oRs1.Close : oRs.Close : SET oRs1 = Nothing : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
										Response.Write "FAIL|||||주문취소할 수 없는 상태의 상품 있습니다.[11]"
										Response.End
								END IF
						END IF
						oRs1.Close

						IF oRs("ProductType") = "P" THEN
								OrderCnt	= OrderCnt		+ oRs("OrderCnt")
						END IF
				ELSE
						oRs.Close : SET oRs1 = Nothing : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
						Response.Write "FAIL|||||주문취소할 수 없는 상태의 상품 있습니다.[12]"
						Response.End
				END IF

				oRs.MoveNext
		Loop
ELSE
		oRs.Close : SET oRs1 = Nothing : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||취소할 주문상품이 없습니다.[13]"
		Response.End
END IF
oRs.Close


IF OrderCnt < TotalOrderCnt THEN
		IF EscrowFlag = "Y" THEN
				SET oRs1 = Nothing : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "FAIL|||||에스크로 적용 주문은 부분취소할 수 없습니다.[14]"
				Response.End
		END IF

		CancelType	= "PartialCancel"
END IF
'-----------------------------------------------------------------------------------------------------------'
'# 주문상품 상태 체크 End
'-----------------------------------------------------------------------------------------------------------'



ON ERROR RESUME NEXT



oConn.BeginTrans



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

										'-----------------------------------------------------------------------------------------------------------'
										' 업체별 추가배송비(EShop_Order_DeliveryPrice) 데이터 입력 START
										'-----------------------------------------------------------------------------------------------------------'
										Set oCmd = Server.CreateObject("ADODB.Command")
										WITH oCmd
												.ActiveConnection = oConn
												.CommandType = adCmdStoredProc
												.CommandText = "USP_Front_EShop_Order_DeliveryPrice_Insert"
												.Parameters.Append .CreateParameter("@OrderCode",			adVarChar,	adParamInput,	 20,	OrderCode)
												.Parameters.Append .CreateParameter("@Vendor",				adVarChar,	adParamInput,	 10,	oRs("Vendor"))
												.Parameters.Append .CreateParameter("@OPCIdx",				adInteger,	adParamInput,	   ,	0)
												.Parameters.Append .CreateParameter("@SettlePrice",			adCurrency,	adParamInput,	   ,	SetDeliveryPrice)
												.Parameters.Append .CreateParameter("@RefundPrice",			adCurrency,	adParamInput,	   ,	0)
												.Parameters.Append .CreateParameter("@PayType",				adChar,		adParamInput,	  1,	PayType)
												.Parameters.Append .CreateParameter("@MemberCouponIdx",		adInteger,	adParamInput,	   ,	0)
												.Parameters.Append .CreateParameter("@Memo",				adVarChar,	adParamInput,	500,	"주문취소시 추가배송비")
												.Parameters.Append .CreateParameter("@CreateID",			adVarChar,	adParamInput,	 50,	U_NUM)
												.Parameters.Append .CreateParameter("@CreateIP",			adVarChar,	adParamInput,	 15,	U_IP)
			
												.Execute, , adExecuteNoRecords
										END WITH
										Set oCmd = Nothing

										IF Err.number <> 0 THEN
												oConn.RollbackTrans

												oRs1.Close : oRs.Close
												Set oRs1 = Nothing : Set oRs = Nothing
												oConn.Close : Set oConn = Nothing

												Response.Write "FAIL|||||주문취소 중 오류가 발생하였습니다.[21]"
												Response.End
										END IF
										'-----------------------------------------------------------------------------------------------------------'
										'업체별 추가배송비(EShop_Order_DeliveryPrice) 데이터 입력 END
										'-----------------------------------------------------------------------------------------------------------'
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

		Response.Write "FAIL|||||환불금액이 부족하여 취소하실 수 없습니다. 고객센터에 문의해 주십시오"
		Response.End
END IF
'-----------------------------------------------------------------------------------------------------------'
'# 환불금액 계산 End
'-----------------------------------------------------------------------------------------------------------'


'-----------------------------------------------------------------------------------------------------------'	
'주문취소 START
'-----------------------------------------------------------------------------------------------------------'	
OPIdx	= Split(OPIdx, ", ")

FOR i = 0 TO UBOUND(OPIdx)
		'-----------------------------------------------------------------------------------------------------------'	
		'결제완료 주문취소  업데이트 START
		'-----------------------------------------------------------------------------------------------------------'	
		'2. 주문 상품 정보 테이블에 주문 상태 정보 Update
		'3. 쿠폰/포인트/슈즈상품권 환원 처리
		'	3-1. 쿠폰 환원 처리 Upudate
		'	3-2. 포인트, 슈즈상품권 환원 처리
		'		3-2-1. 포인트 환원 처리
		'			3-2-1-1. 포인트 사용 삭제 처리
		'			3-2-1-2. 포인트 사용이력 삭제 처리 시작
		'				3-2-1-2-1. 회원포인트 사용이력 삭제
		'				3-2-1-2-2. 회원포인트 사용차감
		'			3-2-1-3. 회원정보 포인트 누적처리
		'		3-2-2. 슈즈상품권 환원 처리 시작
		'			3-2-2-1. 슈즈상품권 사용 삭제 처리
		'			3-2-2-2. 슈즈상품권 사용이력 삭제 처리
		'				3-2-2-2-1. 회원슈즈상품권 사용이력 삭제
		'				3-2-2-2-2. 회원슈즈상품권 사용차감
		'			3-2-2-3. 회원정보 슈즈상품권 누적처리
		'	3-3. 임직원쿠폰 사용 처리 Upudate
		'	3-4. 주문 상품 재고 Upudate
		'-----------------------------------------------------------------------------------------------------------'	
		Set oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection = oConn
				.CommandType = adCmdStoredProc
				.CommandText = "USP_Front_EShop_Order_Update_For_OrderCancel"
				.Parameters.Append .CreateParameter("@OrderCode",			adVarChar,	adParamInput,	 20,		OrderCode)
				.Parameters.Append .CreateParameter("@OPIdx",				adInteger,	adParamInput,	   ,		OPIdx(i))
				.Parameters.Append .CreateParameter("@UpdateNM",			adVarChar,	adParamInput,	100,		U_NAME)
				.Parameters.Append .CreateParameter("@UpdateID",			adVarChar,	adParamInput,	 20,		U_NUM)
				.Parameters.Append .CreateParameter("@UpdateIP",			adVarChar,	adParamInput,	 15,		U_IP)
			
				.Execute, , adExecuteNoRecords
		END WITH
		Set oCmd = Nothing

		IF Err.number <> 0 THEN
				oConn.RollbackTrans

				SET oRs1 = Nothing : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "FAIL|||||주문취소 중 오류가 발생하였습니다.[1]"
				Response.End
		END IF
		'-----------------------------------------------------------------------------------------------------------'	
		'EShop_Order  업데이트 End
		'-----------------------------------------------------------------------------------------------------------'	

NEXT
'-----------------------------------------------------------------------------------------------------------'	
'주문취소 End
'-----------------------------------------------------------------------------------------------------------'	



'-----------------------------------------------------------------------------------------------------------'	
'결제 취소요청 Start
'-----------------------------------------------------------------------------------------------------------'	
IF RefundPrice > 0 AND IsNull(LGD_TID) = false AND LGD_TID <> "" THEN

		'# 네이버페이
		IF PayType = "N" THEN

				'-----------------------------------------------------------------------------------------------------------'	
				'ERP 전송용 I/F 주문 생성 START
				'-----------------------------------------------------------------------------------------------------------'	
				FOR i = 0 TO UBOUND(OPIdx)
						wQuery = ""
						wQuery = wQuery & "WHERE A.IsShowFlag = 'Y' AND A.ProductType IN ('P','O') "
						wQuery = wQuery & "AND A.OrderCode = '" & OrderCode & "' "
						wQuery = wQuery & "AND A.OPIdx_Group = " & OPIdx(i) & " "

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
										'# 주문/결제 생성전송
										SET oCmd = Server.CreateObject("ADODB.Command")
										WITH oCmd
												.ActiveConnection	 = oConn
												.CommandType		 = adCmdStoredProc
												.CommandText		 = "USP_Admin_IF_ONLINE_ORDER_Insert_With_IF_ONLINE_ORDER_APP"

												.Parameters.Append .CreateParameter("@Idx",			 adInteger,	 adParamInput,     ,	 oRs("Idx"))
												.Parameters.Append .CreateParameter("@DOCTYPECD",	 adVarChar,	 adParamInput,   40,	 "CANCEL")
												.Parameters.Append .CreateParameter("@CreateID",	 adVarChar,	 adParamInput,   50,	 U_NUM)
												.Parameters.Append .CreateParameter("@CreateIP",	 adVarChar,	 adParamInput,   20,	 U_IP)

												.Execute, , adExecuteNoRecords
										END WITH
										SET oCmd = Nothing

										IF Err.Number <> 0 THEN
												oConn.RollbackTrans

												oRs.Close
												SET oRs1 = Nothing : SET oRs = Nothing
												oConn.Close
												SET oConn = Nothing

												Response.Write "FAIL|||||주문취소 처리 도중 오류가 발생하였습니다.[31]"
												Response.End
										END IF

										'-----------------------------------------------------------------------------------------------------------'	
										'문자발송 시작
										'-----------------------------------------------------------------------------------------------------------'	
										IF oRs("ProductType") = "P" THEN
												SET oCmd = Server.CreateObject("ADODB.Command")
												WITH oCmd
														.ActiveConnection	 = oConn
														.CommandType		 = adCmdStoredProc
														.CommandText		 = "USP_Admin_EShop_Order_Sms_Send"

														.Parameters.Append .CreateParameter("@OrderCode",	 adVarChar,	 adParamInput,   20,	 oRs("OrderCode"))
														.Parameters.Append .CreateParameter("@OPIdx",		 adInteger,	 adParamInput,     ,	 oRs("Idx"))
														.Parameters.Append .CreateParameter("@SmsCode",		 adVarChar,	 adParamInput,   20,	 "ORD_SC00")

														.Execute, , adExecuteNoRecords
												END WITH
												SET oCmd = Nothing
										END IF
										'-----------------------------------------------------------------------------------------------------------'	
										'문자발송 끝
										'-----------------------------------------------------------------------------------------------------------'	

										IF Err.Number <> 0 THEN
												oConn.RollbackTrans

												oRs.Close
												SET oRs1 = Nothing : SET oRs = Nothing
												oConn.Close
												SET oConn = Nothing

												Response.Write "FAIL|||||주문취소 처리 도중 오류가 발생하였습니다.[32]"
												Response.End
										END IF

										oRs.MoveNext
								Loop 
						End IF
						oRs.Close
				NEXT
				'-----------------------------------------------------------------------------------------------------------'	
				'ERP 전송용 I/F 주문 생성 End
				'-----------------------------------------------------------------------------------------------------------'	


				'-----------------------------------------------------------------------------------------------------------'	
				'네이버페이 결제취소 Start
				'-----------------------------------------------------------------------------------------------------------'	
				DIM HTTP_Object
				DIM ResponseText	: ResponseText	= ""
				DIM Read_Data
				DIM ResultCode
				DIM ResultMsg
				DIM Param

				Param	= ""
				Param	= Param & "paymentId="			& LGD_TID
				Param	= Param & "&cancelAmount="		& RefundPrice
				Param	= Param & "&taxScopeAmount="	& RefundPrice
				Param	= Param & "&taxExScopeAmount=" & "0"
				Param	= Param & "&cancelReason="		& "Order Cancel"
				Param	= Param & "&cancelRequester="	& "1"				'# 1:사용자, 2:관리자

				Set HTTP_Object = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
				With HTTP_Object
						'결제취소는 API 통신 Timeout 을 60초로 지정
						.SetTimeouts 60000, 60000, 60000, 60000
						.Open				"POST",						NAVER_PAY_CANCELURL, False
						.SetRequestHeader	"Content-Type",				"application/x-www-form-urlencoded"
						.SetRequestHeader	"X-Naver-Client-Id",		NAVER_PAY_CLIENTID
						.SetRequestHeader	"X-Naver-Client-Secret",	NAVER_PAY_CLIENTSECRET
						.Send				Param
						.WaitForResponse

						IF .Status = 200 THEN
								ResponseText = .ResponseText
						ELSE
								ResponseText = ""
						END IF
				End With

				IF ResponseText <> "" THEN
						Set Read_Data = New aspJSON
						Read_Data.loadJSON(ResponseText)
						With Read_Data
								ResultCode		= .data("code")
								ResultMsg		= .data("message")
						End With
				End If

				IF ResultCode <> "Success" THEN
						oConn.RollbackTrans

						SET oRs1 = Nothing : SET oRs = Nothing
						oConn.Close
						SET oConn = Nothing

						Response.Write "FAIL|||||주문취소 처리 도중 오류가 발생하였습니다.[33] [" & ResultMsg & "]"
						Response.End
				END IF
				'-----------------------------------------------------------------------------------------------------------'	
				'네이버페이 결제취소 End
				'-----------------------------------------------------------------------------------------------------------'	


				'-----------------------------------------------------------------------------------------------------------'	
				'결제 정보 저장 START
				'-----------------------------------------------------------------------------------------------------------'
				Set oCmd = Server.CreateObject("ADODB.Command")
				WITH oCmd
						.ActiveConnection = oConn
						.CommandType = adCmdStoredProc
						.CommandText = "USP_Front_EShop_Order_Settle_Cancel_Insert"
						.Parameters.Append .CreateParameter("@OrderCode",					adVarChar,	adParamInput,	 20,	OrderCode)
						.Parameters.Append .CreateParameter("@LGD_RESPCODE",				adVarChar,	adParamInput,	  4,	LEFT(ResultCode,4))
						.Parameters.Append .CreateParameter("@LGD_RESPMSG",					adVarChar,	adParamInput,	512,	ResultMsg)
						.Parameters.Append .CreateParameter("@LGD_AMOUNT",					adVarChar,	adParamInput,	 12,	RefundPrice)
						.Parameters.Append .CreateParameter("@LGD_MID",						adVarChar,	adParamInput,	 15,	"")
						.Parameters.Append .CreateParameter("@LGD_TID",						adVarChar,	adParamInput,	 24,	LGD_TID)
						.Parameters.Append .CreateParameter("@LGD_OID",						adVarChar,	adParamInput,	 64,	OrderCode)
						.Parameters.Append .CreateParameter("@LGD_TIMESTAMP",				adVarChar,	adParamInput,	 14,	U_DATE & U_TIME)
						.Parameters.Append .CreateParameter("@LGD_PAYTYPE",					adVarChar,	adParamInput,	  6,	"NPAY")
						.Parameters.Append .CreateParameter("@LGD_RFBANKCODE",				adVarChar,	adParamInput,	  2,	RefundBankCode)
						.Parameters.Append .CreateParameter("@LGD_RFACCOUNTNUM",			adVarChar,	adParamInput,	 20,	RefundAccountNum)
						.Parameters.Append .CreateParameter("@LGD_RFCUSTOMERNAME",			adVarChar,	adParamInput,	 40,	RefundAccountName)
						.Parameters.Append .CreateParameter("@LGD_RFPHONE",					adVarChar,	adParamInput,	 20,	RefundPhone)
						.Parameters.Append .CreateParameter("@CreateID",					adVarChar,	adParamInput,	 50,	U_NUM)
						.Parameters.Append .CreateParameter("@CreateIP",					adVarChar,	adParamInput,	 15,	U_IP)

						.Execute, , adExecuteNoRecords
				END WITH
				Set oCmd = Nothing

				IF Err.number <> 0 THEN
				END IF
				'-----------------------------------------------------------------------------------------------------------'	
				'결제 정보 저장 End
				'-----------------------------------------------------------------------------------------------------------'	


		'# LGU+
		ELSE
				'/*
				' * [결제취소 요청 페이지]
				' *
				' * LG유플러스으로 부터 내려받은 거래번호(LGD_TID)를 가지고 취소 요청을 합니다.(파라미터 전달시 POST를 사용하세요)
				' * (승인시 LG유플러스으로 부터 내려받은 PAYKEY와 혼동하지 마세요.)
				' */

				'CST_PLATFORM         = trim(request("CST_PLATFORM"))        ' LG유플러스 결제서비스 선택(test:테스트, service:서비스)
				'CST_MID              = trim(request("CST_MID"))             ' LG유플러스으로 부터 발급받으신 상점아이디를 입력하세요.
																			' 테스트 아이디는 't'를 제외하고 입력하세요.
				IF PAY_PLATFORM = "test" THEN                               ' 상점아이디(자동생성)
						LGD_MID = "t" & CST_MID
				ELSE
						LGD_MID = CST_MID
				END IF
				'#LGD_TID               = trim(request("LGD_TID"))          ' LG유플러스으로 부터 내려받은 거래번호(LGD_TID)
				LGD_CANCELREASON        = "주문취소"                        ' 취소사유
				IF N_NAME = "" THEN
						LGD_CANCELREQUESTER     = U_NAME                            ' 취소요청자
				ELSE
						LGD_CANCELREQUESTER     = N_NAME                            ' 취소요청자
				END IF
				LGD_CANCELREQUESTERIP   = U_IP                              ' 취소요청IP
    
				SELECT CASE PayType
					CASE "C" : LGD_PAYTYPE = "SC0010"
					CASE "B" : LGD_PAYTYPE = "SC0030"
					CASE "V" : LGD_PAYTYPE = "SC0040"
					CASE "M" : LGD_PAYTYPE = "SC0060"
				END SELECT

				' ※ 중요
				' 환경설정 파일의 경우 반드시 외부에서 접근이 가능한 경로에 두시면 안됩니다.
				' 해당 환경파일이 외부에 노출이 되는 경우 해킹의 위험이 존재하므로 반드시 외부에서 접근이 불가능한 경로에 두시기 바랍니다. 
				' 예) [Window 계열] C:\inetpub\wwwroot\lgdacom -- 절대불가(웹 디렉토리)
				DIM configPath
				configPath = "C:/LGDacom"  'LG유플러스에서 제공한 환경파일("/conf/lgdacom.conf, /conf/mall.conf") 위치 지정.  

				DIM xpay				' 결제요청 API 객체

				Set xpay = CreateObject("XPayClientCOM.XPayClient")
				xpay.Init configPath, PAY_PLATFORM
				xpay.Init_TX(LGD_MID)

				xpay.Set "LGD_TXNAME",				CancelType
				xpay.Set "LGD_TID",					LGD_TID
				xpay.Set "LGD_OID",					OrderCode
				xpay.Set "LGD_CANCELREASON",		LGD_CANCELREASON
				xpay.Set "LGD_CANCELREQUESTER",		LGD_CANCELREQUESTER
				xpay.Set "LGD_CANCELREQUESTERIP",	LGD_CANCELREQUESTERIP

				'# 부분취소일 경우 부분취소할 금액 입력
				IF CancelType = "PartialCancel" THEN
						xpay.Set "LGD_CANCELAMOUNT",			RefundPrice
				END IF
 
				'# 가상계좌일 경우 환불계좌 입력
				IF PayType = "V" THEN
						xpay.Set "LGD_RFBANKCODE",			RefundBankCode
						xpay.Set "LGD_RFACCOUNTNUM",		RefundAccountNum
						xpay.Set "LGD_RFCUSTOMERNAME",		RefundAccountName
						xpay.Set "LGD_RFPHONE",				RefundPhone
				END IF

				'/*
				' * 1. 결제취소 요청 결과처리
				' *
				' * 취소결과 리턴 파라미터는 연동메뉴얼을 참고하시기 바랍니다.
				' *
				' * [[[중요]]] 고객사에서 정상취소 처리해야할 응답코드
				' * 1. 신용카드 : 0000, AV11  
				' * 2. 계좌이체 : 0000, RF00, RF10, RF09, RF15, RF19, RF23, RF25 (환불진행중 응답-> 환불결과코드.xls 참고)
				' * 3. 나머지 결제수단의 경우 0000(성공) 만 취소성공 처리
				' *
				' */

				IF xpay.TX() THEN
						'1)결제취소결과 화면처리(성공,실패 결과 처리를 하시기 바랍니다.)
						'Response.Write("결제취소 요청이 완료되었습니다. <br>")
						'Response.Write("TX Response_code = " & xpay.resCode & "<br>")
						'Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")
						LGD_RESPCODE				 = xpay.Response("LGD_RESPCODE", 0)
						LGD_RESPMSG					 = xpay.Response("LGD_RESPMSG", 0)
						'LGD_TID						 = xpay.Response("LGD_TID", 0)
						'LGD_TIMESTAMP				 = xpay.Response("LGD_TIMESTAMP", 0)
						'LGD_PAYTYPE					 = xpay.Response("LGD_PAYTYPE", 0)
						'LGD_RFBANKCODE				 = xpay.Response("LGD_RFBANKCODE", 0)
						'LGD_RFACCOUNTNUM			 = xpay.Response("LGD_RFACCOUNTNUM", 0)
						'LGD_RFCUSTOMERNAME			 = xpay.Response("LGD_RFCUSTOMERNAME", 0)
						'LGD_RFPHONE					 = xpay.Response("LGD_RFPHONE", 0)

						IF LGD_RESPCODE = "0000" THEN
								'-----------------------------------------------------------------------------------------------------------'	
								'결제 정보 저장 START
								'-----------------------------------------------------------------------------------------------------------'
								Set oCmd = Server.CreateObject("ADODB.Command")
								WITH oCmd
										.ActiveConnection = oConn
										.CommandType = adCmdStoredProc
										.CommandText = "USP_Front_EShop_Order_Settle_Cancel_Insert"
										.Parameters.Append .CreateParameter("@OrderCode",					adVarChar,	adParamInput,	 20,	OrderCode)
										.Parameters.Append .CreateParameter("@LGD_RESPCODE",				adVarChar,	adParamInput,	  4,	LGD_RESPCODE)
										.Parameters.Append .CreateParameter("@LGD_RESPMSG",					adVarChar,	adParamInput,	512,	LGD_RESPMSG)
										.Parameters.Append .CreateParameter("@LGD_AMOUNT",					adVarChar,	adParamInput,	 12,	RefundPrice)
										.Parameters.Append .CreateParameter("@LGD_MID",						adVarChar,	adParamInput,	 15,	LGD_MID)
										.Parameters.Append .CreateParameter("@LGD_TID",						adVarChar,	adParamInput,	 24,	LGD_TID)
										.Parameters.Append .CreateParameter("@LGD_OID",						adVarChar,	adParamInput,	 64,	OrderCode)
										.Parameters.Append .CreateParameter("@LGD_TIMESTAMP",				adVarChar,	adParamInput,	 14,	U_DATE & U_TIME)
										.Parameters.Append .CreateParameter("@LGD_PAYTYPE",					adVarChar,	adParamInput,	  6,	LGD_PAYTYPE)
										.Parameters.Append .CreateParameter("@LGD_RFBANKCODE",				adVarChar,	adParamInput,	  2,	RefundBankCode)
										.Parameters.Append .CreateParameter("@LGD_RFACCOUNTNUM",			adVarChar,	adParamInput,	 20,	RefundAccountNum)
										.Parameters.Append .CreateParameter("@LGD_RFCUSTOMERNAME",			adVarChar,	adParamInput,	 40,	RefundAccountName)
										.Parameters.Append .CreateParameter("@LGD_RFPHONE",					adVarChar,	adParamInput,	 20,	RefundPhone)
										.Parameters.Append .CreateParameter("@CreateID",					adVarChar,	adParamInput,	 50,	U_NUM)
										.Parameters.Append .CreateParameter("@CreateIP",					adVarChar,	adParamInput,	 15,	U_IP)

										.Execute, , adExecuteNoRecords
								END WITH
								Set oCmd = Nothing

								IF Err.number <> 0 THEN
										oConn.RollbackTrans

										SET oRs1 = Nothing : Set oRs = Nothing
										oConn.Close
										Set oConn = Nothing

										xpay.Rollback("상점 DB처리 실패로 인하여 Rollback 처리 [TID:" & LGD_TID & ",MID:" & LGD_MID & ",OID:" & OrderCode & "]")
            		
										IF "0000" = xpay.resCode THEN
												Call SettleErrorLogWrite(LGD_OID, "Y", "PR11", "OrderCancelOk", "EShop_Order_Settle_Cancel 입력 오류 / " & GetPayType(PayType) & " 취소 완료", Err.Description)
												Response.Write "FAIL|||||주문취소 처리 도중 오류가 발생하였습니다.[11]"
												Response.End
										ELSE
												Call SettleErrorLogWrite(LGD_OID, "Y", "PR12", "OrderCancelOk", "EShop_Order_Settle_Cancel 입력 오류 / " & GetPayType(PayType) & " 취소 오류", Err.Description)
												Response.Write "FAIL|||||주문취소 처리 도중 오류가 발생하였습니다.[12]"
												Response.End
										END IF
								END IF
								'-----------------------------------------------------------------------------------------------------------'	
								'결제 정보 저장 End
								'-----------------------------------------------------------------------------------------------------------'	


								'-----------------------------------------------------------------------------------------------------------'	
								'ERP 전송용 I/F 주문 생성 START
								'-----------------------------------------------------------------------------------------------------------'	
								FOR i = 0 TO UBOUND(OPIdx)
										wQuery = ""
										wQuery = wQuery & "WHERE A.IsShowFlag = 'Y' AND A.ProductType IN ('P','O') "
										wQuery = wQuery & "AND A.OrderCode = '" & OrderCode & "' "
										wQuery = wQuery & "AND A.OPIdx_Group = " & OPIdx(i) & " "

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
														'# 주문/결제 생성전송
														SET oCmd = Server.CreateObject("ADODB.Command")
														WITH oCmd
																.ActiveConnection	 = oConn
																.CommandType		 = adCmdStoredProc
																.CommandText		 = "USP_Admin_IF_ONLINE_ORDER_Insert_With_IF_ONLINE_ORDER_APP"

																.Parameters.Append .CreateParameter("@Idx",			 adInteger,	 adParamInput,     ,	 oRs("Idx"))
																.Parameters.Append .CreateParameter("@DOCTYPECD",	 adVarChar,	 adParamInput,   40,	 "CANCEL")
																.Parameters.Append .CreateParameter("@CreateID",	 adVarChar,	 adParamInput,   50,	 U_NUM)
																.Parameters.Append .CreateParameter("@CreateIP",	 adVarChar,	 adParamInput,   20,	 U_IP)

																.Execute, , adExecuteNoRecords
														END WITH
														SET oCmd = Nothing

														IF Err.Number <> 0 THEN
																oConn.RollbackTrans

																oRs.Close
																SET oRs1 = Nothing : SET oRs = Nothing
																oConn.Close
																SET oConn = Nothing

																xpay.Rollback("상점 DB처리 실패로 인하여 Rollback 처리 [TID:" & LGD_TID & ",MID:" & LGD_MID & ",OID:" & OrderCode & "]")
            		
																IF "0000" = xpay.resCode THEN
																		Call SettleErrorLogWrite(LGD_OID, "Y", "PR21", "OrderCancelOk", "ERP I/F 주문 입력 오류 / " & GetPayType(PayType) & " 취소 완료", Err.Description)
																		Response.Write "FAIL|||||주문취소 처리 도중 오류가 발생하였습니다.[21]"
																		Response.End
																ELSE
																		Call SettleErrorLogWrite(LGD_OID, "Y", "PR22", "OrderCancelOk", "ERP I/F 주문 입력 오류 / " & GetPayType(PayType) & " 취소 오류", Err.Description)
																		Response.Write "FAIL|||||주문취소 처리 도중 오류가 발생하였습니다.[22]"
																		Response.End
																END IF
														END IF

														'-----------------------------------------------------------------------------------------------------------'	
														'문자발송 시작
														'-----------------------------------------------------------------------------------------------------------'	
														IF oRs("ProductType") = "P" THEN
																SET oCmd = Server.CreateObject("ADODB.Command")
																WITH oCmd
																		.ActiveConnection	 = oConn
																		.CommandType		 = adCmdStoredProc
																		.CommandText		 = "USP_Admin_EShop_Order_Sms_Send"

																		.Parameters.Append .CreateParameter("@OrderCode",	 adVarChar,	 adParamInput,   20,	 oRs("OrderCode"))
																		.Parameters.Append .CreateParameter("@OPIdx",		 adInteger,	 adParamInput,     ,	 oRs("Idx"))
																		.Parameters.Append .CreateParameter("@SmsCode",		 adVarChar,	 adParamInput,   20,	 "ORD_SC00")

																		.Execute, , adExecuteNoRecords
																END WITH
																SET oCmd = Nothing
														END IF
														'-----------------------------------------------------------------------------------------------------------'	
														'문자발송 끝
														'-----------------------------------------------------------------------------------------------------------'	

														oRs.MoveNext
												Loop 
										End IF
										oRs.Close
								NEXT
								'-----------------------------------------------------------------------------------------------------------'	
								'ERP 전송용 I/F 주문 생성 End
								'-----------------------------------------------------------------------------------------------------------'	


								'-----------------------------------------------------------------------------------------------------------'	
								'문자발송 시작
								'-----------------------------------------------------------------------------------------------------------'	
								'Server.Execute("/Common/SMS/OrderSmsSend.asp")
								'-----------------------------------------------------------------------------------------------------------'	
								'문자발송 끝
								'-----------------------------------------------------------------------------------------------------------'	
						ELSE
								'-----------------------------------------------------------------------------------------------------------'	
								'결제 정보 저장 START
								'-----------------------------------------------------------------------------------------------------------'
								Set oCmd = Server.CreateObject("ADODB.Command")
								WITH oCmd
										.ActiveConnection = oConn
										.CommandType = adCmdStoredProc
										.CommandText = "USP_Front_EShop_Order_Settle_Cancel_Insert"
										.Parameters.Append .CreateParameter("@OrderCode",					adVarChar,	adParamInput,	 20,	OrderCode)
										.Parameters.Append .CreateParameter("@LGD_RESPCODE",				adVarChar,	adParamInput,	  4,	LGD_RESPCODE)
										.Parameters.Append .CreateParameter("@LGD_RESPMSG",					adVarChar,	adParamInput,	512,	LGD_RESPMSG)
										.Parameters.Append .CreateParameter("@LGD_AMOUNT",					adVarChar,	adParamInput,	 12,	RefundPrice)
										.Parameters.Append .CreateParameter("@LGD_MID",						adVarChar,	adParamInput,	 15,	LGD_MID)
										.Parameters.Append .CreateParameter("@LGD_TID",						adVarChar,	adParamInput,	 24,	LGD_TID)
										.Parameters.Append .CreateParameter("@LGD_OID",						adVarChar,	adParamInput,	 64,	OrderCode)
										.Parameters.Append .CreateParameter("@LGD_TIMESTAMP",				adVarChar,	adParamInput,	 14,	U_DATE & U_TIME)
										.Parameters.Append .CreateParameter("@LGD_PAYTYPE",					adVarChar,	adParamInput,	  6,	LGD_PAYTYPE)
										.Parameters.Append .CreateParameter("@LGD_RFBANKCODE",				adVarChar,	adParamInput,	  2,	RefundBankCode)
										.Parameters.Append .CreateParameter("@LGD_RFACCOUNTNUM",			adVarChar,	adParamInput,	 20,	RefundAccountNum)
										.Parameters.Append .CreateParameter("@LGD_RFCUSTOMERNAME",			adVarChar,	adParamInput,	 40,	RefundAccountName)
										.Parameters.Append .CreateParameter("@LGD_RFPHONE",					adVarChar,	adParamInput,	 20,	RefundPhone)
										.Parameters.Append .CreateParameter("@CreateID",					adVarChar,	adParamInput,	 50,	U_NUM)
										.Parameters.Append .CreateParameter("@CreateIP",					adVarChar,	adParamInput,	 15,	U_IP)

										.Execute, , adExecuteNoRecords
								END WITH
								Set oCmd = Nothing
				
								IF Err.number <> 0 THEN
										oConn.RollbackTrans

										SET oRs1 = Nothing : Set oRs = Nothing
										oConn.Close
										Set oConn = Nothing

										Response.Write "FAIL|||||주문취소 처리 도중 오류가 발생하였습니다.[41]"
										Response.End
								END IF
								'-----------------------------------------------------------------------------------------------------------'	
								'결제 정보 저장 End
								'-----------------------------------------------------------------------------------------------------------'	

						END IF
				ELSE
						'2)API 요청 실패 화면처리
						'Response.Write("결제취소 요청이 실패하였습니다. <br>")
						'Response.Write("TX Response_code = " & xpay.resCode & "<br>")
						'Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")

						oConn.RollbackTrans

						SET oRs1 = Nothing : Set oRs = Nothing
						oConn.Close
						Set oConn = Nothing

						Response.Write "FAIL|||||주문취소 처리 도중 오류가 발생하였습니다.[01]"
						Response.End
				END IF
		END IF
END IF
'-----------------------------------------------------------------------------------------------------------'	
'결제취소 End
'-----------------------------------------------------------------------------------------------------------'	


oConn.CommitTrans


'-----------------------------------------------------------------------------------------------------------'	
'# 보증보험 취소 Start
'-----------------------------------------------------------------------------------------------------------'	
'# IF GuaranteeInsurance = "Y" AND GuaranteeInsuranceGubun = "A0" THEN
'# 
'# 		DIM USafeCom
'# 		DIM USafeComResult
'# 		Set USafeCom = CreateObject( "USafeCom.guarantee.1"  )
'# 
'# 		USafeCom.Port = 80
'# 		USafeCom.Url = "gateway.usafe.co.kr"
'# 		USafeCom.CallForm = "/esafe/guartrn.asp"
'# 
'# 		'데이터 64Bit 암호화시 사용
'# 		USafeCom.EncKey = "uclick"
'# 
'# 	
'# 		'///////////////////////////////////////////////////////////////////////////
'# 		USafeCom.gubun 		= "B0"	                         
'# 		USafeCom.mallId		= USAFE_ID
'# 		USafeCom.oId			= OrderCode	' 상점의 주문번호
'# 		'// 테스트를 위해 코딩 end
'# 		'///////////////////////////////////////////////////////////////////////////
'# 
'# 		USafeComResult = USafeCom.cancelInsurance
'# END IF
'-----------------------------------------------------------------------------------------------------------'	
'# 보증보험 취소 End
'-----------------------------------------------------------------------------------------------------------'	



Response.Write "OK|||||"

SET oRs1 = Nothing
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>