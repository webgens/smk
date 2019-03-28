<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'OrderAddOk.asp - 주문정보 생성 처리
'Date		: 2018.12.30
'Update		: 
'*****************************************************************************************'

'# 페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'Response.CharSet = "euc-kr"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include Virtual = "/ADO/ADODBCommon.asp" -->
<!-- #include Virtual = "/Common/Common.asp" -->
<!-- #include Virtual = "/Common/SetInfo.asp" -->
<!-- #include Virtual = "/Common/ProgID1.asp" -->
<!-- #include Virtual = "/Common/md5.asp" -->



<%
IF INSTR(LCASE(HOME_URL), LCASE(Request.ServerVariables("HTTP_HOST")) ) <= 0 THEN
		Response.Write "FAIL|||||잘못된 경로로 접근하셨습니다" & Request.ServerVariables("HTTP_HOST")
		Response.End
END IF




'ON ERROR RESUME NEXT





'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn							'# ADODB Connection 개체
DIM oRs								'# ADODB Recordset 개체
DIM oRs1							'# ADODB Recordset 개체
DIM oRs2							'# ADODB Recordset 개체
DIM oRs3							'# ADODB Recordset 개체
DIM oCmd							'# ADODB Command 개체

'DIM i
DIM j
DIM X

'# 다중배송여부
DIM MultiDelvFlag

'# 주문자 정보
DIM OrderName
DIM OrderTel
DIM OrderTel1
DIM OrderTel23
DIM OrderHp
DIM OrderHp1
DIM OrderHp23
DIM OrderEmail
DIM OrderEmail1
DIM OrderEmail2
DIM OrderZipCode
DIM OrderAddr1
DIM OrderAddr2

'# 배송지 정보
DIM AddressName
DIM ReceiveName
DIM ReceiveTel
DIM ReceiveTel1
DIM ReceiveTel23
DIM ReceiveHp
DIM ReceiveHp1
DIM ReceiveHp23
DIM ReceiveZipCode
DIM ReceiveAddr1
DIM ReceiveAddr2
DIM MainFlag
DIM Memo								'# 주문시메모

'# 결제정보
DIM PayType								'# 결제수단
DIM SettleFlag


DIM OrderCode							'# 주문입력후 생성된 주문코드
DIM OPIdx_Org							'# 주문상품입력후 생성된 주문상품일련번호

'# 주문상품 정보
DIM OrderSheetIdx
DIM TagPrice							'# 개별 상품의 공급가격
DIM SalePrice							'# 개별 상품의 판매가격
DIM DCRate								'# 개별 상품의 판매할인율
DIM OrderCnt							'# 개별 상품의 주문수량
DIM OrderPrice							'# 개별 상품의 주문가격
DIM UseCouponPrice						'# 개별 상품의 쿠폰 할인가
DIM UseScashPrice						'# 개별 상품의 슈즈상품권 할인가
DIM UsePointPrice						'# 개별 상품의 포인트 할인가
DIM DiscountPrice						'# 개별 상품의 할인가
DIM SavePoint							'# 개발 상품의 적립금
DIM DeliveryPrice						'# 배송비

DIM ProductCode
DIM ProductCD
DIM ProductName
DIM ProdCD
DIM ColorCD
DIM SizeCD
DIM ShopCD
DIM WareHouseType
DIM OnlineGB
DIM OutShopCD
DIM OutletGB

DIM EventProdCD(2)
DIM EventProdNM(2)
DIM EventProdQty(2)

DIM ProductEmployeeFlag
DIM ProductEmployeeType
DIM ProductEmployeeNo
DIM ProductEmployeeCardNo

DIM ProductAddressName
DIM ProductReceiveName
DIM ProductReceiveTel
DIM ProductReceiveHP
DIM ProductReceiveZipCode
DIM ProductReceiveAddr1
DIM ProductReceiveAddr2
DIM ProductMainFlag


'# 주문상태
DIM OrderState
DIM CancelState1
DIM CancelState2

DIM DB_TotalOrderPrice		: DB_TotalOrderPrice		= 0

'# 주문 총 결제금액 정보
DIM TotalOrderCnt			: TotalOrderCnt				= 0			'# 전체 상품수량
DIM TotalTagPrice			: TotalTagPrice				= 0			'# 전체 공급가격
DIM TotalSalePrice			: TotalSalePrice			= 0			'# 전체 판매가격
DIM TotalOrderPrice			: TotalOrderPrice			= 0			'# 전체 주문가격(실제 주문 가격)
DIM TotalUseCouponPrice		: TotalUseCouponPrice		= 0			'# 전체 쿠폰 할인가
DIM TotalUseScashPrice		: TotalUseScashPrice		= 0			'# 전체 슈즈상품권 할인가
DIM TotalUsePointPrice		: TotalUsePointPrice		= 0			'# 전체 포인트 할인가
DIM TotalDiscountPrice		: TotalDiscountPrice		= 0			'# 전체 할인가
DIM TotalDeliveryPrice		: TotalDeliveryPrice		= 0			'# 전체 배송비
DIM TotalSavePoint			: TotalSavePoint			= 0			'# 전체 적립금
DIM TotalEmployeeOrderCnt	: TotalEmployeeOrderCnt		= 0			'# 전체 임직원가 구매수량
DIM TotalEmployeeOrderPrice	: TotalEmployeeOrderPrice	= 0			'# 전체 임직원가 구매금액

'# 업체별 총 결제금액 정보
DIM ShopOrderCnt			: ShopOrderCnt				= 0
DIM ShopTagPrice			: ShopTagPrice				= 0
DIM ShopSalePrice			: ShopSalePrice				= 0
DIM ShopOrderPrice			: ShopOrderPrice			= 0
DIM ShopUseCouponPrice		: ShopUseCouponPrice		= 0
DIM ShopUseScashPrice		: ShopUseScashPrice			= 0
DIM ShopUsePointPrice		: ShopUsePointPrice			= 0
DIM ShopDiscountPrice		: ShopDiscountPrice			= 0
DIM ShopDeliveryPrice		: ShopDeliveryPrice			= 0
DIM ShopSavePoint			: ShopSavePoint				= 0

'# 회원 정보
DIM PointRate				: PointRate					= 0			'# 회원의 구매 적립율
DIM MemberScash				: MemberScash				= 0			'# 회원의 보유 슈즈상품권
DIM MemberPoint				: MemberPoint				= 0			'# 회원의 보유 포인트
DIM EmployeeLimit			: EmployeeLimit				= 0			'# 임직원 구매한도

DIM EmployeeFlag			: EmployeeFlag				= "N"		'# 임직원여부
DIM EmployeeType			: EmployeeType				= "P"		'# 임직원구분(P:일반회원/S:슈마커직원/J:JD직원)
DIM EmployeeNo				: EmployeeNo				= ""		'# 임직원번호
DIM EmployeeCardNo			: EmployeeCardNo			= ""		'# 임직원카드번호

'# 쿠폰정보
DIM CouponIdx
DIM CouponName
DIM CouponType
DIM LimitPriceType
DIM LimitPrice
DIM MoneyType
DIM Discount
DIM ApplyPriceType
DIM LimitDiscountFlag
DIM LimitDiscount
DIM DuplicateUseFlag
DIM CouponDcPrice

DIM CouponIdxs
DIM DuplicateUseFlags

'# PG사 결제요청 정보
DIM LGD_CLOSEDATE						'# 결제가능일시(가상계좌 입금마감일시)
DIM LGD_PRODUCTINFO						'# 결제에 넘어가는 구매내역
DIM LGD_BUYERADDRESS
DIM LGD_RECEIVENAME
DIM LGD_RECEIVEZIPCODE
DIM LGD_RECEIVEADDR1
DIM LGD_RECEIVEADDR2
DIM LGD_RECEIVEHP
DIM N

'# 에스크로
DIM LGD_ESCROW_USEYN	: LGD_ESCROW_USEYN	= "N"

'# Usafe보증보험 관련
DIM GuaranteeInsurance
DIM GuaranteeInsuranceAgreement
DIM USafeJumin1
DIM USafeJumin2
DIM USafeEmailFlag
DIM USafeSmsFlag

'# 네이버페이 결제요청 정보
DIM NPay_ProductName		: NPay_ProductName	= ""
DIM NPay_ProductItems		: NPay_ProductItems	= ""

DIM HTTP_USER_AGENT
DIM USER_AGENT
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'



'# 로그인 검사
'# IF U_NUM = "" THEN
'# 		CALL AlertMessage2 ("로그인이 필요한 서비스입니다", "history.back();")
'# 		Response.End
'# END IF

HTTP_USER_AGENT = LCase(Request.ServerVariables("HTTP_USER_AGENT"))
IF InStr(HTTP_USER_AGENT, "android") THEN
		USER_AGENT = "A"
ELSEIF InStr(HTTP_USER_AGENT, "iphone") OR InStr(HTTP_USER_AGENT, "ipad") OR InStr(HTTP_USER_AGENT, "ipod") THEN
		USER_AGENT = "N"
ELSE
		USER_AGENT = "Y"
END IF


MultiDelvFlag						 = Trim(sqlFilter(Request("MultiDelvFlag")))

OrderName							 = Trim(sqlFilter(Request("OrderName")))
OrderTel1							 = Trim(sqlFilter(Request("OrderTel1")))
OrderTel23							 = Trim(sqlFilter(Request("OrderTel23")))
OrderHp1							 = Trim(sqlFilter(Request("OrderHp1")))
OrderHp23							 = Trim(sqlFilter(Request("OrderHp23")))
OrderEmail							 = Trim(sqlFilter(Request("OrderEmail")))
'# OrderEmail1							 = Trim(sqlFilter(Request("OrderEmail1")))
'# OrderEmail2							 = Trim(sqlFilter(Request("OrderEmail2")))
OrderZipCode						 = Trim(sqlFilter(Request("OrderZipCode")))
OrderAddr1							 = Trim(sqlFilter(Request("OrderAddr1")))
OrderAddr2							 = Trim(sqlFilter(Request("OrderAddr2")))

AddressName							 = Trim(sqlFilter(Request("AddressName")))
ReceiveName							 = Trim(sqlFilter(Request("ReceiveName")))
ReceiveTel1							 = Trim(sqlFilter(Request("ReceiveTel1")))
ReceiveTel23						 = Trim(sqlFilter(Request("ReceiveTel23")))
ReceiveHp1							 = Trim(sqlFilter(Request("ReceiveHp1")))
ReceiveHp23							 = Trim(sqlFilter(Request("ReceiveHp23")))
ReceiveZipCode						 = Trim(sqlFilter(Request("ReceiveZipCode")))
ReceiveAddr1						 = Trim(sqlFilter(Request("ReceiveAddr1")))
ReceiveAddr2						 = Trim(sqlFilter(Request("ReceiveAddr2")))
MainFlag							 = Trim(sqlFilter(Request("MainFlag")))
Memo								 = Trim(sqlFilter(Request("Memo")))
PayType								 = Trim(sqlFilter(Request("PayType")))

IF AddressName		= "" THEN AddressName	= ReceiveName
IF MainFlag			= "" THEN MainFlag		= "N"

'# Usafe보증보험 관련
IF USAFE_FLAG = "Y" THEN
	    LGD_ESCROW_USEYN					 = "N"
		GuaranteeInsurance					 = sqlFilter(Request("GuaranteeInsurance")) 
		GuaranteeInsuranceAgreement			 = sqlFilter(Request("GuaranteeInsuranceAgreement")) 
		USafeJumin1							 = sqlFilter(Request("USafeYear")) & sqlFilter(Request("USafeMonth")) & sqlFilter(Request("USafeDay"))
		USafeJumin2							 = sqlFilter(Request("USafeSex")) 
		USafeEmailFlag						 = "Y" 
		USafeSmsFlag						 = "Y" 
		'# 설정된 결제수단일 경우에만 보증보험을 발급한다
		IF PayType <> USAFE_PAYTYPE	THEN GuaranteeInsurance = "N" 
		IF GuaranteeInsurance = ""	THEN GuaranteeInsurance = "N"
ELSE
		IF PayType = "B" OR PayType = "V" THEN
				LGD_ESCROW_USEYN = "Y"
		ELSE
				LGD_ESCROW_USEYN = "N"
		END IF
		GuaranteeInsurance					 = "N"
		GuaranteeInsuranceAgreement			 = "N"
		USafeJumin1							 = ""
		USafeJumin2							 = ""
		USafeEmailFlag						 = "N" 
		USafeSmsFlag						 = "N" 
END IF

OrderTel							 = ChgTel(OrderTel1 & OrderTel23)
OrderHp								 = ChgTel(OrderHp1 & OrderHp23)
'# OrderEmail							 = OrderEmail1 & "@" & OrderEmail2
ReceiveTel							 = ChgTel(ReceiveTel1 & ReceiveTel23)
ReceiveHp							 = ChgTel(ReceiveHp1 & ReceiveHp23)



'# 결제가능일시(가상계좌 입금마감일시)
DIM CurDate 
CurDate								 = DATEADD("d", MALL_CLOSEDATE, Now)  
LGD_CLOSEDATE						 = YEAR(CurDate) & RIGHT("0" & MONTH(CurDate), 2) & RIGHT("0" & DAY(CurDate), 2) & RIGHT("0" & HOUR(CurDate), 2) & RIGHT("0" & MINUTE(CurDate), 2) & RIGHT("0" & SECOND(CurDate), 2)

IF OrderState						 = "" THEN OrderState	 = "0"
IF CancelState1						 = "" THEN CancelState1	 = "0"
IF CancelState2						 = "" THEN CancelState2	 = "0"
IF SettleFlag						 = "" THEN SettleFlag	 = "N"


'--변수 검사
'Response.write "U_ID : " & U_ID & "<br />"
'Response.write "OrderName : " & OrderName & "<br />"
'Response.write "OrderTel : "& OrderTel & "<br />"
'Response.write "OrderHp : " & OrderHp & "<br />"
'Response.write "OrderEmail : " & OrderEmail & "<br />"
'Response.write "ReceiveName : " & ReceiveName & "<br />"
'Response.write "ReceiveTel : " & ReceiveTel & "<br />"
'Response.write "ReceiveHp : " & ReceiveHp & "<br />"
'Response.write "ReceiveZip1 : " & ReceiveZip1 & "<br />"
'Response.write "ReceiveZip2 : " & ReceiveZip2 & "<br />"
'Response.write "ReceiveAddress1 : " & ReceiveAddress1 & "<br />"
'Response.write "ReceiveAddress2 : " & ReceiveAddress2 & "<br />"
'Response.write "OpenPrice : " & OpenPrice & "<br />"
'Response.write "SalePrice : " & SalePrice & "<br />"
'Response.write "Amount : " & Amount & "<br />"
'Response.write "UseCouponAmount : " & UseCouponAmount & "<br />"
'Response.write "DeliveryPrice : " & DeliveryPrice & "<br />"
'Response.write "PayType : " & PayType & "<br />"
'Response.write "OrderState : " & OrderState & "<br />"
'Response.write "CancelState1  : " & CancelState1 & "<br />"
'Response.write "CancelState2 : " & CancelState2 & "<br />"
'Response.write "ReceiptFlag : " & ReceiptFlag & "<br />"
'Response.write "ReceiptKinds : " & ReceiptKinds & "<br />"
'Response.write "SettleFlag : " & SettleFlag & "<br />"
'Response.write "CouponCode1 : " & CouponCode1 & "<br />"


SET oConn	= ConnectionOpen()	'//커넥션 생성
SET oRs		= Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성
SET oRs1	= Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성
SET oRs2	= Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성
SET oRs3	= Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성



'-----------------------------------------------------------------------------------------------------------'
'주문서 테이블 검색 시작
'-----------------------------------------------------------------------------------------------------------'
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.Commandtype = adCmdStoredProc
		.CommandText = "USP_Front_EShop_OrderSheet_Select_For_OrderCount"

		.Parameters.Append .CreateParameter("@CartID",	adVarChar, adParaminput,		20,		U_CARTID)
End WITH
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing

'//주문서 테이블에 상품이 없을시 데이터 튕겨냄
IF NOT oRs.EOF THEN
		IF CInt(oRs("OrderCount")) = 0 THEN
				oRs.Close
				Set oRs3 = Nothing
				Set oRs2 = Nothing
				Set oRs1 = Nothing
				Set oRs = Nothing
				oConn.Close
				Set oConn = Nothing

				Response.Write "FAIL|||||주문하실 상품데이터가 없습니다. 다시 주문해 주세요."
				Response.End
		END IF

		'# 일반택배 배송이 없으면 에스크로 적용안함
		'# IF CInt(oRs("DelvType_P")) = 0 THEN
		'# 		LGD_ESCROW_USEYN = "N"
		'# END IF
		
ELSE
		oRs.Close
		Set oRs3 = Nothing
		Set oRs2 = Nothing
		Set oRs1 = Nothing
		Set oRs = Nothing
		oConn.Close
		Set oConn = Nothing
	
		Response.Write "FAIL|||||주문하실 상품데이터가 없습니다. 다시 주문해 주세요."
		Response.End
END IF
oRs.Close
'-----------------------------------------------------------------------------------------------------------'
'주문서 테이블 검색 끝
'-----------------------------------------------------------------------------------------------------------'



'-----------------------------------------------------------------------------------------------------------'
'주문서 테이블에 재고가 부족한 상품이 있는지 검색 시작
'-----------------------------------------------------------------------------------------------------------'
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.Commandtype = adCmdStoredProc
		.CommandText = "USP_Front_EShop_OrderSheet_Select_For_StockCheck"

		.Parameters.Append .CreateParameter("@CartID",	adVarChar, adParaminput,		20,		U_CARTID)
End WITH
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing

IF NOT oRs.EOF THEN
		oRs.Close
		Set oRs3 = Nothing
		Set oRs2 = Nothing
		Set oRs1 = Nothing
		Set oRs = Nothing
		oConn.Close
		Set oConn = Nothing
	
		Response.Write "FAIL|||||주문하신 상품중 품절이거나 재고가 부족한 상품이 있습니다."
		Response.End
End IF
oRs.Close
'-----------------------------------------------------------------------------------------------------------'
'주문서 테이블에 재고가 부족한 상품이 있는지 검색 시작
'-----------------------------------------------------------------------------------------------------------'





'-----------------------------------------------------------------------------------------------------------'
'총 주문금액 검색 START
'-----------------------------------------------------------------------------------------------------------'
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Front_EShop_OrderSheet_Select_For_TotalOrderPrice"

		.Parameters.Append .CreateParameter("@CartID",	adVarChar,	adParamInput, 20,	U_CARTID)
END WITH
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing

IF NOT oRs.EOF THEN
		DB_TotalOrderPrice			 = oRs("OrderPrice")

		TotalOrderCnt				 = oRs("OrderCnt")
		TotalTagPrice				 = oRs("TagPrice")
		TotalSalePrice				 = oRs("SalePrice")
		TotalOrderPrice				 = oRs("OrderPrice")
		TotalUseCouponPrice			 = oRs("UseCouponPrice")
		TotalUseScashPrice			 = oRs("UseScashPrice")
		TotalUsePointPrice			 = oRs("UsePointPrice")
		TotalEmployeeOrderCnt		 = oRs("EmployeeOrderCnt")
		TotalEmployeeOrderPrice		 = oRs("EmployeeOrderPrice")
ELSE
		DB_TotalOrderPrice			 = 0

		TotalOrderCnt				 = 0
		TotalTagPrice				 = 0
		TotalSalePrice				 = 0
		TotalOrderPrice				 = 0
		TotalUseCouponPrice			 = 0
		TotalUseScashPrice			 = 0
		TotalUsePointPrice			 = 0
		TotalEmployeeOrderCnt		 = 0
		TotalEmployeeOrderPrice		 = 0
END IF
oRs.Close
'-----------------------------------------------------------------------------------------------------------'
'총 주문금액 검색 END
'-----------------------------------------------------------------------------------------------------------'


'-----------------------------------------------------------------------------------------------------------'
'포인트, 슈즈상품권 유효성 검사 시작
'사용하려는 포인트 합계가 회원이 가지고 있는 포인트/슈즈상품권보다 많으면 롤백시키고 주문취소
'-----------------------------------------------------------------------------------------------------------'
IF U_NUM <> "" THEN
		Set oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection = oConn
				.CommandType = adCmdStoredProc
				.CommandText = "USP_Admin_EShop_Member_Select_By_MemberNum"

				.Parameters.Append .CreateParameter("@MemberNum",	adInteger,	adParamInput, ,		U_NUM)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		Set oCmd = Nothing

		IF NOT oRs.EOF THEN
				PointRate		= oRs("PointRate")

				MemberPoint		= oRs("Point")
				MemberScash		= oRs("Scash")

				EmployeeFlag	= oRs("EmployeeFlag")						'# 임직원여부
				IF EmployeeFlag = "Y" THEN
						EmployeeType	= oRs("EmployeeType")				'# 임직원구분(P:일반회원/S:슈마커직원/J:JD직원)
						EmployeeNo		= oRs("EmployeeNo")					'# 임직원번호
						EmployeeCardNo	= oRs("EmployeeCardNo")				'# 임직원카드번호
				ELSE
						EmployeeType	= "P"								'# 임직원구분(P:일반회원/S:슈마커직원/J:JD직원)
						EmployeeNo		= ""								'# 임직원번호
						EmployeeCardNo	= ""								'# 임직원카드번호
				END IF
		END IF
		oRs.Close
END IF

IF CDbl(TotalUsePointPrice) > 0 AND CDbl(TotalUsePointPrice) > CDbl(MemberPoint) THEN
		Set oRs3 = Nothing
		Set oRs2 = Nothing
		Set oRs1 = Nothing
		Set oRs = Nothing
		oConn.Close
		Set oConn = Nothing
	
		Response.Write "FAIL|||||주문에 적용한 포인트가 현재 보유한 포인트보다 많습니다."
		Response.End
END IF

IF CDbl(TotalUseScashPrice) > 0 AND CDbl(TotalUseScashPrice) > CDbl(MemberScash) THEN
		Set oRs3 = Nothing
		Set oRs2 = Nothing
		Set oRs1 = Nothing
		Set oRs = Nothing
		oConn.Close
		Set oConn = Nothing
	
		Response.Write "FAIL|||||주문에 적용한 슈즈상품권이 현재 보유한 슈즈상품권보다 많습니다."
		Response.End
END IF
'-----------------------------------------------------------------------------------------------------------'
'포인트, 슈즈상품권 유효성 검사 끝
'-----------------------------------------------------------------------------------------------------------'




'-----------------------------------------------------------------------------------------------------------'
'임직원가 구매한도 유효성 검사 시작
'-----------------------------------------------------------------------------------------------------------'
IF EmployeeType = "S" THEN
		'-----------------------------------------------------------------------------------------------------------'
		'슈마커 임직원가 구매금액 유효성 검사 시작
		'-----------------------------------------------------------------------------------------------------------'
		Set oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection = oConn
				.CommandType = adCmdStoredProc
				.CommandText = "USP_Front_ERP_IF_ONLINE_STAFFCARDLIMIT_V_Select_By_EMPNO"

				.Parameters.Append .CreateParameter("@LINKED_SERVER_NAME",	adVarChar,	adParamInput, 20,		ERP_LNK_SRV)
				.Parameters.Append .CreateParameter("@TABLE_NAME",			adVarChar,	adParamInput, 50,		ERP_SCL_TBL)
				.Parameters.Append .CreateParameter("@EMPNO",				adVarChar,	adParamInput, 10,		EmployeeNo)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		Set oCmd = Nothing

		IF NOT oRs.EOF THEN
				EmployeeLimit	= oRs("REMAINAMT")
		ELSE
				EmployeeLimit	= 0
		END IF
		oRs.Close


		IF CDbl(TotalEmployeeOrderPrice) > 0 AND CDbl(TotalEmployeeOrderPrice) > CDbl(EmployeeLimit) THEN
				Set oRs3 = Nothing
				Set oRs2 = Nothing
				Set oRs1 = Nothing
				Set oRs = Nothing
				oConn.Close
				Set oConn = Nothing
	
				Response.Write "FAIL|||||주문에 적용한 임직원가 구매금액이 현재 남아있는 한도금액보다 많습니다."
				Response.End
		END IF
		'-----------------------------------------------------------------------------------------------------------'
		'슈마커 임직원가 구매금액 유효성 검사 끝
		'-----------------------------------------------------------------------------------------------------------'
ELSEIF EmployeeType = "J" THEN
		'-----------------------------------------------------------------------------------------------------------'
		'JD 임직원가 구매수량 유효성 검사 시작
		'-----------------------------------------------------------------------------------------------------------'
		Set oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection = oConn
				.CommandType = adCmdStoredProc
				.CommandText = "USP_Front_EShop_Coupon_Employee_Select_For_Useable_Count"

				.Parameters.Append .CreateParameter("@EmployeeType",	adChar,		adParamInput,  1,		EmployeeType)
				.Parameters.Append .CreateParameter("@EmployeeNo",		adVarChar,	adParamInput, 10,		EmployeeNo)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		Set oCmd = Nothing

		IF NOT oRs.EOF THEN
				EmployeeLimit	= oRs("CouponCount")
		ELSE
				EmployeeLimit	= 0
		END IF
		oRs.Close


		IF CDbl(TotalEmployeeOrderCnt) > 0 AND CDbl(TotalEmployeeOrderCnt) > CDbl(EmployeeLimit) THEN
				Set oRs3 = Nothing
				Set oRs2 = Nothing
				Set oRs1 = Nothing
				Set oRs = Nothing
				oConn.Close
				Set oConn = Nothing
	
				Response.Write "FAIL|||||주문에 적용한 임직원가 구매수량이 현재 남아있는 한도수량보다 많습니다."
				Response.End
		END IF
		'-----------------------------------------------------------------------------------------------------------'
		'JD 임직원가 구매수량 유효성 검사 끝
		'-----------------------------------------------------------------------------------------------------------'
END IF
'-----------------------------------------------------------------------------------------------------------'
'임직원가 구매한도 유효성 검사 끝
'-----------------------------------------------------------------------------------------------------------'





oConn.BeginTrans






'-----------------------------------------------------------------------------------------------------------'
'EShop_Order 데이터 입력 START
'-----------------------------------------------------------------------------------------------------------'
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Front_EShop_Order_Insert"
		.Parameters.Append .CreateParameter("@UserID",				adVarChar,	adParamInput,	 50,	U_NUM)
		.Parameters.Append .CreateParameter("@CartId",				adVarChar,	adParamInput,	 20,	U_CARTID)
		.Parameters.Append .CreateParameter("@OrderName",			adVarChar,	adParamInput,	 50,	OrderName)
		.Parameters.Append .CreateParameter("@OrderTel",			adVarChar,	adParamInput,	 20,	OrderTel)
		.Parameters.Append .CreateParameter("@OrderHp",				adVarChar,	adParamInput,	 20,	OrderHp)
		.Parameters.Append .CreateParameter("@OrderEmail",			adVarChar,	adParamInput,	 50,	OrderEmail)
		.Parameters.Append .CreateParameter("@OrderZipCode",		adVarChar,	adParamInput,	  7,	OrderZipCode)
		.Parameters.Append .CreateParameter("@OrderAddr1",			adVarChar,	adParamInput,	200,	OrderAddr1)
		.Parameters.Append .CreateParameter("@OrderAddr2",			adVarChar,	adParamInput,	200,	OrderAddr2)
		.Parameters.Append .CreateParameter("@TagPrice",			adCurrency,	adParamInput,	   ,	TagPrice)
		.Parameters.Append .CreateParameter("@SalePrice",			adCurrency,	adParamInput,	   ,	SalePrice)
		.Parameters.Append .CreateParameter("@OrderPrice",			adCurrency,	adParamInput,	   ,	OrderPrice)
		.Parameters.Append .CreateParameter("@UseCouponPrice",		adCurrency,	adParamInput,	   ,	TotalUseCouponPrice)
		.Parameters.Append .CreateParameter("@UseScashPrice",		adCurrency,	adParamInput,	   ,	TotalUseScashPrice)
		.Parameters.Append .CreateParameter("@UsePointPrice",		adCurrency,	adParamInput,	   ,	TotalUsePointPrice)
		.Parameters.Append .CreateParameter("@DeliveryPrice",		adCurrency,	adParamInput,	   ,	DeliveryPrice)
		.Parameters.Append .CreateParameter("@PayType",				adChar,		adParamInput,	  1,	PayType)
		.Parameters.Append .CreateParameter("@ReceiptFlag",			adChar,		adParamInput,	  1,	"N")
		.Parameters.Append .CreateParameter("@ReceiptKinds",		adChar,		adParamInput,	  1,	"")
		.Parameters.Append .CreateParameter("@OrderDate",			adChar,		adParamInput,	  8,	U_DATE)
		.Parameters.Append .CreateParameter("@OrderTime",			adChar,		adParamInput,	  6,	U_TIME)
		.Parameters.Append .CreateParameter("@SettleFlag",			adChar,		adParamInput,	  1,	SettleFlag)
		.Parameters.Append .CreateParameter("@CloseDate",			adVarChar,	adParamInput,	 14,	LGD_CLOSEDATE)
		.Parameters.Append .CreateParameter("@Memo",				adVarChar,	adParamInput,	500,	Memo)
		.Parameters.Append .CreateParameter("@EscrowFlag",			adChar,		adParamInput,	  1,	LGD_ESCROW_USEYN)
		.Parameters.Append .CreateParameter("@GuaranteeInsurance",	adChar,		adParamInput,	  1,	GuaranteeInsurance)
		.Parameters.Append .CreateParameter("@USafeJumin1",			adVarChar,	adParamInput,	 20,	USafeJumin1)
		.Parameters.Append .CreateParameter("@USafeJumin2",			adVarChar,	adParamInput,	 20,	USafeJumin2)
		.Parameters.Append .CreateParameter("@MemberGroupCode",		adInteger,	adParamInput,	   ,	U_GROUP)
		.Parameters.Append .CreateParameter("@Location",			adChar,		adParamInput,	  1,	"A")
		.Parameters.Append .CreateParameter("@CreateID",			adVarChar,	adParamInput,	 20,	U_NUM)
		.Parameters.Append .CreateParameter("@CreateIP",			adVarChar,	adParamInput,	 15,	U_IP)
		.Parameters.Append .CreateParameter("@OrderCode",			adVarChar,	adParamOutput,	 20)
			
		.Execute, , adExecuteNoRecords

		OrderCode = .Parameters("@OrderCode").Value
END WITH
Set oCmd = Nothing

IF Err.number <> 0 THEN
		oConn.RollbackTrans

		Set oRs3 = Nothing
		Set oRs2 = Nothing
		Set oRs1 = Nothing
		Set oRs = Nothing
		oConn.Close
		Set oConn = Nothing
	
		Response.Write "FAIL|||||주문 처리 도중 오류가 발생하였습니다. [10001]"
		Response.End
END IF
'-----------------------------------------------------------------------------------------------------------'
'EShop_Order 데이터 입력 END
'-----------------------------------------------------------------------------------------------------------'



'-----------------------------------------------------------------------------------------------------------'
'주문서 테이블 검색 시작
'-----------------------------------------------------------------------------------------------------------'
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_OrderSheet_Select_For_ShopList_By_CartID"

		.Parameters.Append .CreateParameter("@CartID", adVarChar, adParamInput, 20, U_CARTID)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing


TotalOrderCnt			= 0
TotalTagPrice			= 0
TotalSalePrice			= 0
TotalOrderPrice			= 0
TotalUseCouponPrice		= 0
TotalUsePointPrice		= 0
TotalUseScashPrice		= 0
TotalSavePoint			= 0
TotalDeliveryPrice		= 0


IF NOT oRs.EOF THEN
		N = 0
		Do Until oRs.EOF 

				ShopOrderCnt		= 0
				ShopTagPrice		= 0
				ShopSalePrice		= 0
				ShopOrderPrice		= 0
				ShopUseCouponPrice	= 0
				ShopUsePointPrice	= 0
				ShopUseScashPrice	= 0
				ShopDiscountPrice	= 0
				ShopDeliveryPrice	= 0
				ShopSavePoint		= 0


				'-----------------------------------------------------------------------------------------------------------'
				'배송형태별 업체별 주문서 상품 테이블 검색 시작
				'-----------------------------------------------------------------------------------------------------------'
				SET oCmd = Server.CreateObject("ADODB.Command")
				WITH oCmd
						.ActiveConnection	 = oConn
						.CommandType		 = adCmdStoredProc
						.CommandText		 = "USP_Front_EShop_OrderSheet_Select_By_CartID_DelvType_ShopCD"

						.Parameters.Append .CreateParameter("@CartID",		adVarChar,	adParamInput, 20,	oRs("CartID"))
						.Parameters.Append .CreateParameter("@DelvType",	adChar,		adParamInput,  1,	oRs("DelvType"))
						.Parameters.Append .CreateParameter("@ShopCD",		adChar,		adParamInput,  6,	oRs("ShopCD"))
				END WITH
				oRs1.CursorLocation = adUseClient
				oRs1.Open oCmd, , adOpenStatic, adLockReadOnly
				SET oCmd = Nothing

				IF NOT oRs1.EOF THEN
						Do Until oRs1.EOF

								N = N + 1

								OrderSheetIdx		= oRs1("Idx")
								TagPrice			= oRs1("TagPrice")

								IF oRs1("SalePriceType") = "2" THEN
										SalePrice				= oRs1("EmployeeSalePrice")
										DCRate					= oRs1("EmployeeDCRate")

										ProductEmployeeFlag		= EmployeeFlag
										ProductEmployeeType		= EmployeeType
										ProductEmployeeNo		= EmployeeNo
										ProductEmployeeCardNo	= EmployeeCardNo
								ELSE
										SalePrice				= oRs1("SalePrice")
										DCRate					= oRs1("DCRate")

										ProductEmployeeFlag		= "N"
										ProductEmployeeType		= "P"
										ProductEmployeeNo		= ""
										ProductEmployeeCardNo	= ""
								END IF

								UseCouponPrice			= oRs1("UseCouponPrice")
								UseScashPrice			= oRs1("UseScashPrice")
								UsePointPrice			= oRs1("UsePointPrice")
								DiscountPrice			= UseCouponPrice + UseScashPrice + UsePointPrice

								OrderCnt				= oRs1("OrderCnt")
								OrderPrice				= CDbl(SalePrice) - CDbl(DiscountPrice)
								SavePoint				= Int(OrderPrice * CDbl(PointRate) / 100 + 0.5)

								'# 개별 배송비 계산 (일반택배, 기준금액미만 주문시 배송비 부과)
								IF oRs1("DelvType") = "P" AND SalePrice < CDbl(oRs("StandardPrice")) THEN
										DeliveryPrice	= CDbl(oRs("DeliveryPrice"))
								ELSE
										DeliveryPrice	= 0
								END IF

								'# 단일배송지 이거나 배송지등록안된 일반택배는 메인 배송지 적용
								IF oRs1("DelvType") = "P" AND (MultiDelvFlag = "N" OR IsNull(oRs1("ReceiveName")) OR oRs1("ReceiveName") = "") THEN
										ProductAddressName			= AddressName
										ProductReceiveName			= ReceiveName
										ProductReceiveTel			= ReceiveTel
										ProductReceiveHP			= ReceiveHP
										ProductReceiveZipCode		= ReceiveZipCode
										ProductReceiveAddr1			= ReceiveAddr1
										ProductReceiveAddr2			= ReceiveAddr2
										ProductMainFlag				= MainFlag
								ELSE
										ProductAddressName			= oRs1("AddressName")
										ProductReceiveName			= oRs1("ReceiveName")
										ProductReceiveTel			= oRs1("ReceiveTel")
										ProductReceiveHP			= oRs1("ReceiveHP")
										ProductReceiveZipCode		= oRs1("ReceiveZipCode")
										ProductReceiveAddr1			= oRs1("ReceiveAddr1")
										ProductReceiveAddr2			= oRs1("ReceiveAddr2")
										ProductMainFlag				= oRs1("MainFlag")
								END IF

								'# LG데이콤 PG사에 넘기는 상품명
								'# 에스크로 거래이면 일반택배 거래건으로 셋팅
								'# IF LGD_ESCROW_USEYN = "Y" THEN
								'# 		IF oRs1("DelvType") = "P" AND LGD_PRODUCTINFO = "" THEN
								'# 				LGD_PRODUCTINFO		= oRs1("ProductName")
								'# 				LGD_BUYERADDRESS	= ProductReceiveZipCode & " " & ProductReceiveAddr1 & " " & ProductReceiveAddr2
								'# 				LGD_RECEIVENAME		= ProductReceiveName
								'# 				LGD_RECEIVEZIPCODE	= ProductReceiveZipCode
								'# 				LGD_RECEIVEADDR1	= ProductReceiveAddr1
								'# 				LGD_RECEIVEADDR2	= ProductReceiveAddr2
								'# 				LGD_RECEIVEHP		= ProductReceiveHP
								'# 		END IF

								'# 에스크로 거래가 아니면 첫번째 거래건으로 셋팅
								'# ELSE
										IF N = 1 THEN
												LGD_PRODUCTINFO		= oRs1("ProductName")
												LGD_BUYERADDRESS	= ProductReceiveZipCode & " " & ProductReceiveAddr1 & " " & ProductReceiveAddr2
												LGD_RECEIVENAME		= ProductReceiveName
												LGD_RECEIVEZIPCODE	= ProductReceiveZipCode
												LGD_RECEIVEADDR1	= ProductReceiveAddr1
												LGD_RECEIVEADDR2	= ProductReceiveAddr2
												LGD_RECEIVEHP		= ProductReceiveHP
										END IF
								'# END IF

								'# 네이버페이 결제요청시 전송할 상품리스트
								IF NPay_ProductItems = "" THEN
										NPay_ProductName	= oRs1("ProductName")
										NPay_ProductItems	= "{""categoryType"": ""PRODUCT"", ""categoryId"": ""GENERAL"", ""uid"": """ & oRs1("ProductCD") & """, ""name"": """ & oRs1("ProductName") & """, ""payReferrer"": ""PARTNER_DIRECT"", ""count"": " & oRs1("OrderCnt") & "}"
								ELSE
										NPay_ProductItems	= NPay_ProductItems & ", " & "{""categoryType"": ""PRODUCT"", ""categoryId"": ""GENERAL"", ""uid"": """ & oRs1("ProductCD") & """, ""name"": """ & oRs1("ProductName") & """, ""payReferrer"": ""PARTNER_DIRECT"", ""count"": " & oRs1("OrderCnt") & "}"
								END IF


								'-----------------------------------------------------------------------------------------------------------'
								'재고 체크 START
								'-----------------------------------------------------------------------------------------------------------'
								'# 매장픽업
								IF oRs1("DelvType") = "S" THEN
										ShopCD			= oRs1("PickupShopCD")
										'# 재고보유 여부 체크
										Set oCmd = Server.CreateObject("ADODB.Command")
										WITH oCmd
												.ActiveConnection = oConn
												.CommandType = adCmdStoredProc
												.CommandText = "USP_Admin_EShop_Stock_Select_By_Key"

												.Parameters.Append .CreateParameter("@ProductCode",		 adInteger,	 adParamInput,   ,	 oRs1("ProductCode"))
												.Parameters.Append .CreateParameter("@SizeCD",			 adVarChar,	 adParamInput, 50,	 oRs1("SizeCD"))
												.Parameters.Append .CreateParameter("@ShopCD",			 adVarChar,	 adParamInput, 10,	 ShopCD)
										END WITH
										oRs2.CursorLocation = adUseClient
										oRs2.Open oCmd, , adOpenStatic, adLockReadOnly
										Set oCmd = Nothing

										IF NOT oRs2.EOF THEN
												IF oRs2("UseFlag") <> "Y" OR oRs2("SizeCDUseFlag") <> "Y" OR oRs2("SoldOut") = "Y" OR oRs2("RestQty") <= 0 THEN
														oConn.RollbackTrans

														oRs2.Close : oRs1.Close : oRs.Close
														Set oRs3 = Nothing : Set oRs2 = Nothing : Set oRs1 = Nothing : Set oRs = Nothing
														oConn.Close : Set oConn = Nothing
	
														Response.Write "FAIL|||||재고부족으로 주문 처리 도중 오류가 발생하였습니다. [12011]"
														Response.End
												END IF
										ELSE
												oConn.RollbackTrans

												oRs2.Close : oRs1.Close : oRs.Close
												Set oRs3 = Nothing : Set oRs2 = Nothing : Set oRs1 = Nothing : Set oRs = Nothing
												oConn.Close : Set oConn = Nothing
	
												Response.Write "FAIL|||||재고부족으로 주문 처리 도중 오류가 발생하였습니다. [12012]"
												Response.End
										END IF
										oRs2.Close
								'# 일반택배
								ELSE
										'# 재고보유 창고/매장 찾기
										Set oCmd = Server.CreateObject("ADODB.Command")
										WITH oCmd
												.ActiveConnection = oConn
												.CommandType = adCmdStoredProc
												.CommandText = "USP_Admin_EShop_Stock_Select_For_GetShopCD_By_ProductCode_N_SizeCD"

												.Parameters.Append .CreateParameter("@ProductCode",		 adInteger,	 adParamInput,   ,	 oRs1("ProductCode"))
												.Parameters.Append .CreateParameter("@SizeCD",			 adVarChar,	 adParamInput, 50,	 oRs1("SizeCD"))
										END WITH
										oRs2.CursorLocation = adUseClient
										oRs2.Open oCmd, , adOpenStatic, adLockReadOnly
										Set oCmd = Nothing

										IF NOT oRs2.EOF THEN
												ShopCD			= oRs2("ShopCD")
										ELSE
												oConn.RollbackTrans

												oRs2.Close : oRs1.Close : oRs.Close
												Set oRs3 = Nothing : Set oRs2 = Nothing : Set oRs1 = Nothing : Set oRs = Nothing
												oConn.Close : Set oConn = Nothing
	
												Response.Write "FAIL|||||재고부족으로 주문 처리 도중 오류가 발생하였습니다. [12013]"
												Response.End
										END IF
										oRs2.Close
								END IF

								WareHouseType	= GetWareHouseType(ShopCD)
								'-----------------------------------------------------------------------------------------------------------'
								'재고 체크 END
								'-----------------------------------------------------------------------------------------------------------'

								'-----------------------------------------------------------------------------------------------------------'
								'업체정보 체크 START
								'-----------------------------------------------------------------------------------------------------------'
								Set oCmd = Server.CreateObject("ADODB.Command")
								WITH oCmd
										.ActiveConnection = oConn
										.CommandType = adCmdStoredProc
										.CommandText = "USP_Admin_EShop_Store_Select_By_ShopCD"
										.Parameters.Append .CreateParameter("@ShopCD",		 adChar,	 adParamInput, 6,	 oRs1("ShopCD"))
								END WITH
								oRs2.CursorLocation = adUseClient
								oRs2.Open oCmd, , adOpenStatic, adLockReadOnly
								Set oCmd = Nothing

								IF NOT oRs2.EOF THEN
										'# 입점몰 여부 셋팅
										IF oRs2("OutShopFlag") = "Y" AND oRs2("ShopCD") <> "006740" THEN
												OnlineGB	= "I"
												OutShopCD	= "009899"
										ELSE
												OnlineGB	= "S"
												OutShopCD	= ""
										END IF
								ELSE
										oConn.RollbackTrans

										oRs2.Close : oRs1.Close : oRs.Close
										Set oRs3 = Nothing : Set oRs2 = Nothing : Set oRs1 = Nothing : Set oRs = Nothing
										oConn.Close : Set oConn = Nothing
	
										Response.Write "FAIL|||||주문 처리 도중 오류가 발생하였습니다. [12014]"
										Response.End
								END IF
								oRs2.Close
								'-----------------------------------------------------------------------------------------------------------'
								'업체정보 체크 END
								'-----------------------------------------------------------------------------------------------------------'



								'-----------------------------------------------------------------------------------------------------------'
								'사용쿠폰 유효성 체크 시작
								'-----------------------------------------------------------------------------------------------------------'
								IF CDbl(UseCouponPrice) > 0 THEN
										CouponIdxs				= ""
										DuplicateUseFlags		= ""

										'# 상품에 적용된 쿠폰 유효성 체크
										Set oCmd = Server.CreateObject("ADODB.Command")
										WITH oCmd
												.ActiveConnection = oConn
												.CommandType = adCmdStoredProc
												.CommandText = "USP_Front_EShop_OrderSheet_UseCoupon_Select_By_OrderSheetIdx"
												.Parameters.Append .CreateParameter("@MemberNum",		 adInteger,	 adParamInput, ,	 U_NUM)
												.Parameters.Append .CreateParameter("@OrderSheetIdx",	 adInteger,	 adParamInput, ,	 OrderSheetIdx)
										END WITH
										oRs2.CursorLocation = adUseClient
										oRs2.Open oCmd, , adOpenStatic, adLockReadOnly
										Set oCmd = Nothing

										'# 해당 상품에 적용된 쿠폰 배열
										IF NOT oRs2.EOF THEN
												Do Until oRs2.EOF

														'# 사용가능한 쿠폰인지 체크
														Set oCmd = Server.CreateObject("ADODB.Command")
														WITH oCmd
																.ActiveConnection = oConn
																.CommandType = adCmdStoredProc
																.CommandText = "USP_Front_EShop_Coupon_Member_Select_By_Idx"
																.Parameters.Append .CreateParameter("@MemberNum",			adInteger,	 adParamInput,   ,	U_NUM)
																.Parameters.Append .CreateParameter("@MemberCouponIdx",		adInteger,	 adParamInput,   ,	oRs2("MemberCouponIdx"))
														END WITH
														oRs3.Open oCmd, , adOpenStatic, adLockReadOnly
														Set oCmd = Nothing

														IF NOT oRs3.EOF THEN
																CouponName		= oRs3("CouponName")

																IF oRs3("ReceiveFlag") <> "Y" THEN
																		oConn.RollbackTrans

																		oRs3.Close : oRs2.Close : oRs1.Close : oRs.Close
																		Set oRs3 = Nothing : Set oRs2 = Nothing : Set oRs1 = Nothing : Set oRs = Nothing
																		oConn.Close : Set oConn = Nothing
	
																		Response.Write "FAIL|||||[" & CouponName & "] 쿠폰은 배포되지 않은 쿠폰입니다. [13011]"
																		Response.End

																ELSEIF oRs3("UseFlag") = "Y" THEN
																		oConn.RollbackTrans

																		oRs3.Close : oRs2.Close : oRs1.Close : oRs.Close
																		Set oRs3 = Nothing : Set oRs2 = Nothing : Set oRs1 = Nothing : Set oRs = Nothing
																		oConn.Close : Set oConn = Nothing
	
																		Response.Write "FAIL|||||[" & CouponName & "] 쿠폰은 이미 사용된 쿠폰입니다. [13012]"
																		Response.End

																ELSEIF oRs3("CollectFlag") = "Y" THEN
																		oConn.RollbackTrans

																		oRs3.Close : oRs2.Close : oRs1.Close : oRs.Close
																		Set oRs3 = Nothing : Set oRs2 = Nothing : Set oRs1 = Nothing : Set oRs = Nothing
																		oConn.Close : Set oConn = Nothing
	
																		Response.Write "FAIL|||||[" & CouponName & "] 쿠폰은 회수된 쿠폰입니다. [13013]"
																		Response.End

																ELSEIF oRs3("StartDT") > U_DATE & LEFT(U_TIME,4) THEN
																		oConn.RollbackTrans

																		oRs3.Close : oRs2.Close : oRs1.Close : oRs.Close
																		Set oRs3 = Nothing : Set oRs2 = Nothing : Set oRs1 = Nothing : Set oRs = Nothing
																		oConn.Close : Set oConn = Nothing
	
																		Response.Write "FAIL|||||[" & CouponName & "] 쿠폰은 아직 사용하실 수 없습니다. [13014]"
																		Response.End

																ELSEIF oRs3("EndDT") < U_DATE & LEFT(U_TIME,4) THEN
																		oConn.RollbackTrans

																		oRs3.Close : oRs2.Close : oRs1.Close : oRs.Close
																		Set oRs3 = Nothing : Set oRs2 = Nothing : Set oRs1 = Nothing : Set oRs = Nothing
																		oConn.Close : Set oConn = Nothing
	
																		Response.Write "FAIL|||||[" & CouponName & "] 쿠폰은 사용기한이 지났습니다. [13015]"
																		Response.End

																ELSEIF oRs3("AppFlag") <> "Y" THEN
																		oConn.RollbackTrans

																		oRs3.Close : oRs2.Close : oRs1.Close : oRs.Close
																		Set oRs3 = Nothing : Set oRs2 = Nothing : Set oRs1 = Nothing : Set oRs = Nothing
																		oConn.Close : Set oConn = Nothing
	
																		Response.Write "FAIL|||||[" & CouponName & "] 쿠폰은 PC에서 사용불가한 쿠폰입니다. [13016]"
																		Response.End

																ELSEIF oRs3("DeliveryCouponFlag") = "Y" THEN
																		oConn.RollbackTrans

																		oRs3.Close : oRs2.Close : oRs1.Close : oRs.Close
																		Set oRs3 = Nothing : Set oRs2 = Nothing : Set oRs1 = Nothing : Set oRs = Nothing
																		oConn.Close : Set oConn = Nothing
	
																		Response.Write "FAIL|||||[" & CouponName & "] 쿠폰은 주문시 사용불가한 쿠폰입니다. [13017]"
																		Response.End
																END IF
														ELSE
																oConn.RollbackTrans

																oRs3.Close : oRs2.Close : oRs1.Close : oRs.Close
																Set oRs3 = Nothing : Set oRs2 = Nothing : Set oRs1 = Nothing : Set oRs = Nothing
																oConn.Close : Set oConn = Nothing

																Call AlertMessage2("[" & CouponName & "] 쿠폰은 없는 쿠폰입니다. [13019]", "history.back();")
																Response.End
														END IF
														oRs3.Close


														'# 해당 주문중에 다른 상품에 사용한 쿠폰인지 체크
														Set oCmd = Server.CreateObject("ADODB.Command")
														WITH oCmd
																.ActiveConnection = oConn
																.CommandType = adCmdStoredProc
																.CommandText = "USP_Front_EShop_OrderSheet_UseCoupon_Select_For_Used_Check"
																.Parameters.Append .CreateParameter("@CartID",				adVarChar,	 adParamInput, 20,	U_CARTID)
																.Parameters.Append .CreateParameter("@OrderSheetIdx",		adInteger,	 adParamInput,   ,	OrderSheetIdx)
																.Parameters.Append .CreateParameter("@MemberCouponIdx",		adInteger,	 adParamInput,   ,	oRs2("MemberCouponIdx"))
														END WITH
														oRs3.Open oCmd, , adOpenStatic, adLockReadOnly
														Set oCmd = Nothing

														IF NOT oRs3.EOF THEN
																oConn.RollbackTrans

																oRs3.Close : oRs2.Close : oRs1.Close : oRs.Close
																Set oRs3 = Nothing : Set oRs2 = Nothing : Set oRs1 = Nothing : Set oRs = Nothing
																oConn.Close : Set oConn = Nothing
	
																Response.Write "FAIL|||||[" & CouponName & "] 쿠폰은 다른 상품에 적용한 쿠폰입니다. [13021]"
																Response.End
														END IF
														oRs3.Close


														'# 해당상품에 적용가능 여부
														Set oCmd = Server.CreateObject("ADODB.Command")
														WITH oCmd
																.ActiveConnection = oConn
																.CommandType = adCmdStoredProc
																.CommandText = "USP_Front_EShop_Coupon_Member_Select_For_UseCheck"
																.Parameters.Append .CreateParameter("@MemberNum",			adInteger,	 adParamInput,   ,	 U_NUM)
																.Parameters.Append .CreateParameter("@OrderSheetIdx",		adInteger,	 adParamInput,   ,	 OrderSheetIdx)
																.Parameters.Append .CreateParameter("@MemberCouponIdx",		adInteger,	 adParamInput,   ,	 oRs2("MemberCouponIdx"))
														END WITH
														oRs3.Open oCmd, , adOpenStatic, adLockReadOnly
														Set oCmd = Nothing
	

														'# 해당상품에 적용가능한 쿠폰인 경우
														IF NOT oRs3.EOF THEN

																CouponIdx			= oRs3("CouponIdx")
																CouponName			= oRs3("CouponName")
																CouponType			= oRs3("CouponType")
																LimitPriceType		= oRs3("LimitPriceType")
																LimitPrice			= oRs3("LimitPrice")
																MoneyType			= oRs3("MoneyType")
																Discount			= oRs3("Discount")
																ApplyPriceType		= oRs3("ApplyPriceType")
																LimitDiscountFlag	= oRs3("LimitDiscountFlag")
																LimitDiscount		= oRs3("LimitDiscount")
																DuplicateUseFlag	= oRs3("DuplicateUseFlag")

																'# 임직원 쿠폰일 경우 다른 쿠폰과 사용 금지
																IF CouponType = "99" AND oRs2.RecordCount > 1 THEN
																		oConn.RollbackTrans

																		oRs3.Close : oRs2.Close : oRs1.Close : oRs.Close
																		Set oRs3 = Nothing : Set oRs2 = Nothing : Set oRs1 = Nothing : Set oRs = Nothing
																		oConn.Close : Set oConn = Nothing
	
																		Response.Write "FAIL|||||[" & CouponName & "] 쿠폰은 다른 쿠폰과 중복 사용할 수 없습니다. [13031]"
																		Response.End
																END IF

																'# 판매가 제한
																IF LimitPriceType = "W" THEN
																		IF CDbl(SalePrice) < CDbl(LimitPrice) THEN
																				oConn.RollbackTrans

																				oRs3.Close : oRs2.Close : oRs1.Close : oRs.Close
																				Set oRs3 = Nothing : Set oRs2 = Nothing : Set oRs1 = Nothing : Set oRs = Nothing
																				oConn.Close : Set oConn = Nothing
	
																				Response.Write "FAIL|||||[" & CouponName & "] 쿠폰은 " & FormatNumber(LimitPrice,0) & "원이상 상품에만 사용할 수 있습니다. [13032]"
																				Response.End
																		END IF

																'# 할인율 제한
																ELSEIF LimitPriceType = "P" THEN
																		IF CDbl(DCRate) > (100 - CDbl(LimitPrice)) THEN
																				oConn.RollbackTrans

																				oRs3.Close : oRs2.Close : oRs1.Close : oRs.Close
																				Set oRs3 = Nothing : Set oRs2 = Nothing : Set oRs1 = Nothing : Set oRs = Nothing
																				oConn.Close : Set oConn = Nothing
	
																				Response.Write "FAIL|||||[" & CouponName & "] 쿠폰은 " & FormatNumber(100 - LimitPrice,0) & "% 미만 할인된 상품에만 사용할 수 있습니다. [13033]"
																				Response.End
																		END IF
																END IF


																'# 할인금액 계산
																IF MoneyType = "W" THEN
																		CouponDcPrice	 = Discount
																ELSE
																		IF ApplyPriceType = "T" THEN
																				CouponDcPrice	 = Round(TagPrice * CDbl(Discount) / 1000) * 10
																		ELSE
																				CouponDcPrice	 = Round(SalePrice * CDbl(Discount) / 1000) * 10
																		END IF
																END IF

																'# 최대할인금액 적용
																IF LimitDiscountFlag = "Y" AND CDbl(CouponDcPrice) > CDbl(LimitDiscount) THEN
																		CouponDcPrice	 = LimitDiscount
																END IF

																'# 최소 주문금액 체크
																IF CouponDcPrice > (OrderPrice - MALL_MIN_ORDERPRICE) THEN
																		CouponDcPrice	= OrderPrice - MALL_MIN_ORDERPRICE
																END IF

																'# 할인금액이 0보다 작을 경우
																IF CouponDcPrice <= 0 THEN
																		oConn.RollbackTrans

																		oRs3.Close : oRs2.Close : oRs1.Close : oRs.Close
																		Set oRs3 = Nothing : Set oRs2 = Nothing : Set oRs1 = Nothing : Set oRs = Nothing
																		oConn.Close : Set oConn = Nothing
	
																		Response.Write "FAIL|||||[" & CouponName & "] 쿠폰은 쿠폰할인금액이 0 보다 작아 쿠폰을 사용할 수 없습니다. [13034]"
																		Response.End
																END IF

																'# 똑같은 쿠폰이 여러장인지 체크
																'# 상품 상세에서 쿠폰 여러 장 다운 후 여러장을 한 상품에 사용할 경우 제거
																IF InStr(CouponIdxs, "," & CouponIdx & ",") THEN
																		oConn.RollbackTrans

																		oRs3.Close : oRs2.Close : oRs1.Close : oRs.Close
																		Set oRs3 = Nothing : Set oRs2 = Nothing : Set oRs1 = Nothing : Set oRs = Nothing
																		oConn.Close : Set oConn = Nothing
	
																		Response.Write "FAIL|||||[" & CouponName & "] 쿠폰은 하나만 사용 가능합니다. [13035]"
																		Response.End
																END IF


																'# 중복 사용 가능 쿠폰인지 체크
																'# 해당 쿠폰이 쇼핑몰쿠폰 중복 불가능 쿠폰이면, 다른 쇼핑몰쿠폰이 사용되었는지 체크해서 제거
																IF DuplicateUseFlag = "N" AND CouponIdxs <> "" THEN
																		oConn.RollbackTrans

																		oRs3.Close : oRs2.Close : oRs1.Close : oRs.Close
																		Set oRs3 = Nothing : Set oRs2 = Nothing : Set oRs1 = Nothing : Set oRs = Nothing
																		oConn.Close : Set oConn = Nothing
	
																		Response.Write "FAIL|||||[" & CouponName & "] 쿠폰은 다른 쿠폰과 같이 사용 하실 수 없습니다. [13036]"
																		Response.End
																END IF
	

																'# 중복 사용 가능 쿠폰인지 체크
																'# 쇼핑몰 중복 불가능 쿠폰과 같이 사용하는 경우
																IF DuplicateUseFlag = "Y" AND InStr(DuplicateUseFlags, "N") THEN
																		oConn.RollbackTrans

																		oRs3.Close : oRs2.Close : oRs1.Close : oRs.Close
																		Set oRs3 = Nothing : Set oRs2 = Nothing : Set oRs1 = Nothing : Set oRs = Nothing
																		oConn.Close : Set oConn = Nothing
	
																		Response.Write "FAIL|||||[" & CouponName & "] 쿠폰은 단독사용 쿠폰과 같이 사용 하실 수 없습니다. [13037]"
																		Response.End
																END IF


																CouponIdxs			= CouponIdxs		& "," & CouponIdx			& ","
																DuplicateUseFlags	= DuplicateUseFlags & "," & DuplicateUseFlag	& ","


														'# 해당상품에 적용가능한 쿠폰이 아닌 경우
														ELSE
																oConn.RollbackTrans

																oRs3.Close : oRs2.Close : oRs1.Close : oRs.Close
																Set oRs3 = Nothing : Set oRs2 = Nothing : Set oRs1 = Nothing : Set oRs = Nothing
																oConn.Close : Set oConn = Nothing
	
																Response.Write "FAIL|||||[" & CouponName & "] 쿠폰은 해당 상품에 적용가능한 쿠폰이 아닙니다. [13039]"
																Response.End
														END IF
														oRs3.Close

														oRs2.MoveNext
												Loop 
										END IF
										oRs2.Close
								END IF
								'-----------------------------------------------------------------------------------------------------------'
								'사용쿠폰 유효성 체크 끝
								'-----------------------------------------------------------------------------------------------------------'

								'-----------------------------------------------------------------------------------------------------------'
								'사은품 증정 체크 시작
								'-----------------------------------------------------------------------------------------------------------'
								For j = 0 TO UBound(EventProdCD)
										EventProdCD(j)		= ""
										EventProdNM(j)		= ""
										EventProdQty(j)		= 0
								Next

								SET oCmd = Server.CreateObject("ADODB.Command")
								WITH oCmd
										.ActiveConnection	 = oConn
										.CommandType		 = adCmdStoredProc
										.CommandText		 = "USP_Front_EShop_SubProduct_Event_Select_By_ProductCode"

										.Parameters.Append .CreateParameter("@ProductCode",		 adInteger, adParaminput,		, oRs1("ProductCode"))
								End WITH
								oRs2.CursorLocation = adUseClient
								oRs2.Open oCmd, , adOpenStatic, adLockReadOnly
								SET oCmd = Nothing

								IF NOT oRs2.EOF THEN
										j = 0
										Do Until oRs2.EOF
												EventProdCD(j)		= oRs2("EventProdCD")
												EventProdNM(j)		= oRs2("EventProdNM")
												EventProdQty(j)		= oRs2("Qty")

												oRs2.MoveNext
												j = j + 1
										Loop
								END IF
								oRs2.Close
								'-----------------------------------------------------------------------------------------------------------'
								'사은품 증정 체크 끝
								'-----------------------------------------------------------------------------------------------------------'

								'-----------------------------------------------------------------------------------------------------------'
								'EShop_Order_Product 데이터 입력 START
								'-----------------------------------------------------------------------------------------------------------'
								Set oCmd = Server.CreateObject("ADODB.Command")
								WITH oCmd
										.ActiveConnection = oConn
										.CommandType = adCmdStoredProc
										.CommandText = "USP_Front_EShop_Order_Product_Insert"
										.Parameters.Append .CreateParameter("@OrderCode",				adVarChar,	adParamInput,	 20,	OrderCode)
										.Parameters.Append .CreateParameter("@OPIdx_Group",				adInteger,	adParamInput,	   ,	0)
										.Parameters.Append .CreateParameter("@Vendor",					adVarChar,	adParamInput,	 10,	oRs1("ShopCD"))
										.Parameters.Append .CreateParameter("@ProductCode",				adInteger,	adParamInput,	   ,	oRs1("ProductCode"))
										.Parameters.Append .CreateParameter("@ProductCD",				adVarChar,	adParamInput,	 10,	oRs1("ProductCD"))
										.Parameters.Append .CreateParameter("@ProductName",				adVarChar,	adParamInput,	100,	oRs1("ProductName"))
										.Parameters.Append .CreateParameter("@OrderType",				adChar,		adParamInput,	  1,	oRs1("OrderType"))
										.Parameters.Append .CreateParameter("@ProductType",				adChar,		adParamInput,	  1,	"P")
										.Parameters.Append .CreateParameter("@ProductPoint",			adCurrency,	adParamInput,	   ,	SavePoint)
										.Parameters.Append .CreateParameter("@ProdCD",					adVarChar,	adParamInput,	 50,	oRs1("ProdCD"))
										.Parameters.Append .CreateParameter("@ColorCD",					adVarChar,	adParamInput,	100,	oRs1("ColorCD"))
										.Parameters.Append .CreateParameter("@SizeCD",					adVarChar,	adParamInput,	 50,	oRs1("SizeCD"))
										.Parameters.Append .CreateParameter("@OrderCnt",				adInteger,	adParamInput,	   ,	oRs1("OrderCnt"))
										.Parameters.Append .CreateParameter("@TagPrice",				adCurrency,	adParamInput,	   ,	TagPrice)
										.Parameters.Append .CreateParameter("@SalePrice",				adCurrency,	adParamInput,	   ,	SalePrice)
										.Parameters.Append .CreateParameter("@OrderPrice",				adCurrency,	adParamInput,	   ,	OrderPrice)
										.Parameters.Append .CreateParameter("@UseCouponPrice",			adCurrency,	adParamInput,	   ,	UseCouponPrice)
										.Parameters.Append .CreateParameter("@UseScashPrice",			adCurrency,	adParamInput,	   ,	UseScashPrice)
										.Parameters.Append .CreateParameter("@UsePointPrice",			adCurrency,	adParamInput,	   ,	UsePointPrice)
										.Parameters.Append .CreateParameter("@DeliveryPrice",			adCurrency,	adParamInput,	   ,	DeliveryPrice)
										.Parameters.Append .CreateParameter("@ReceiveName",				adVarChar,	adParamInput,	 50,	ProductReceiveName)
										.Parameters.Append .CreateParameter("@ReceiveTel",				adVarChar,	adParamInput,	 20,	ProductReceiveTel)
										.Parameters.Append .CreateParameter("@ReceiveHp",				adVarChar,	adParamInput,	 20,	ProductReceiveHp)
										.Parameters.Append .CreateParameter("@ReceiveZipCode",			adVarChar,	adParamInput,	  7,	ProductReceiveZipCode)
										.Parameters.Append .CreateParameter("@ReceiveAddr1",			adVarChar,	adParamInput,	200,	ProductReceiveAddr1)
										.Parameters.Append .CreateParameter("@ReceiveAddr2",			adVarChar,	adParamInput,	200,	ProductReceiveAddr2)
										.Parameters.Append .CreateParameter("@Memo",					adVarChar,	adParamInput,	500,	Memo)
										.Parameters.Append .CreateParameter("@OrderState",				adChar,		adParamInput,	  1,	OrderState)
										.Parameters.Append .CreateParameter("@CancelState1",			adChar,		adParamInput,	  1,	CancelState1)
										.Parameters.Append .CreateParameter("@CancelState2",			adChar,		adParamInput,	  1,	CancelState2)
										.Parameters.Append .CreateParameter("@ProductEmployeeFlag",		adChar,		adParamInput,	  1,	ProductEmployeeFlag)
										.Parameters.Append .CreateParameter("@ProductEmployeeType",		adChar,		adParamInput,	  1,	ProductEmployeeType)
										.Parameters.Append .CreateParameter("@ProductEmployeeNo",		adVarChar,	adParamInput,	 10,	ProductEmployeeNo)
										.Parameters.Append .CreateParameter("@ProductEmployeeCardNo",	adVarChar,	adParamInput,	 20,	ProductEmployeeCardNo)
										.Parameters.Append .CreateParameter("@ShopCD",					adVarChar,	adParamInput,	 10,	ShopCD)
										.Parameters.Append .CreateParameter("@WareHouseType",			adChar,		adParamInput,	  1,	WareHouseType)
										.Parameters.Append .CreateParameter("@OnlineGB",				adVarChar,	adParamInput,	 10,	OnlineGB)
										.Parameters.Append .CreateParameter("@OutShopCD",				adVarChar,	adParamInput,	 10,	OutShopCD)
										.Parameters.Append .CreateParameter("@OutletGB",				adChar,		adParamInput,	  1,	oRs1("OutletFlag"))
										.Parameters.Append .CreateParameter("@DelvType",				adChar,		adParamInput,	  1,	oRs1("DelvType"))
	
										.Parameters.Append .CreateParameter("@EventProd1CD",			adVarChar,	adParamInput,	  7,	EventProdCD(0))
										.Parameters.Append .CreateParameter("@EventProd1NM",			adVarChar,	adParamInput,	 50,	EventProdNM(0))
										.Parameters.Append .CreateParameter("@EventProd1ODQty",			adInteger,	adParamInput,	   ,	EventProdQty(0))
										.Parameters.Append .CreateParameter("@EventProd2CD",			adVarChar,	adParamInput,	  7,	EventProdCD(1))
										.Parameters.Append .CreateParameter("@EventProd2NM",			adVarChar,	adParamInput,	 50,	EventProdNM(1))
										.Parameters.Append .CreateParameter("@EventProd2ODQty",			adInteger,	adParamInput,	   ,	EventProdQty(1))
										.Parameters.Append .CreateParameter("@EventProd3CD",			adVarChar,	adParamInput,	  7,	EventProdCD(2))
										.Parameters.Append .CreateParameter("@EventProd3NM",			adVarChar,	adParamInput,	 50,	EventProdNM(2))
										.Parameters.Append .CreateParameter("@EventProd3ODQty",			adInteger,	adParamInput,	   ,	EventProdQty(2))
	
										.Parameters.Append .CreateParameter("@CreateNM",				adVarChar,	adParamInput,	100,	OrderName)
										.Parameters.Append .CreateParameter("@CreateID",				adVarChar,	adParamInput,	 20,	U_NUM)
										.Parameters.Append .CreateParameter("@CreateIP",				adVarChar,	adParamInput,	 15,	U_IP)
										.Parameters.Append .CreateParameter("@OrderSheetIdx",			adInteger,	adParamInput,	   ,	OrderSheetIdx)
										.Parameters.Append .CreateParameter("@OPIdx_Org",				adInteger,	adParamOutput)
			
										.Execute, , adExecuteNoRecords

										OPIdx_Org = .Parameters("@OPIdx_Org").Value
								END WITH
								Set oCmd = Nothing

								IF Err.number <> 0 THEN
										oConn.RollbackTrans

										oRs1.Close : oRs.Close
										Set oRs3 = Nothing : Set oRs2 = Nothing : Set oRs1 = Nothing : Set oRs = Nothing
										oConn.Close : Set oConn = Nothing
	
										Response.Write "FAIL|||||주문 처리 도중 오류가 발생하였습니다. [12001]"
										Response.End
								END IF
								'-----------------------------------------------------------------------------------------------------------'
								'EShop_Order_Product 데이터 입력 END
								'-----------------------------------------------------------------------------------------------------------'


								IF U_NUM <> "" AND oRs1("DelvType") = "P" THEN
										'-----------------------------------------------------------------------------------------------------------'
										' 나의 주소록에 주소 추가 시작
										'-----------------------------------------------------------------------------------------------------------'
										Set oCmd = Server.CreateObject("ADODB.Command")
										WITH oCmd
												.ActiveConnection = oConn
												.CommandType = adCmdStoredProc
												.CommandText = "USP_Front_EShop_MyAddress_Insert"
												.Parameters.Append .CreateParameter("@MemberNum",			adVarChar,	adParamInput,	 20,	U_NUM)
												.Parameters.Append .CreateParameter("@AddressName",			adVarChar,	adParamInput,	 20,	ProductAddressName)
												.Parameters.Append .CreateParameter("@ReceiveName",			adVarChar,	adParamInput,	 50,	ProductReceiveName)
												.Parameters.Append .CreateParameter("@ReceiveTel",			adVarChar,	adParamInput,	 20,	ProductReceiveTel)
												.Parameters.Append .CreateParameter("@ReceiveHp",			adVarChar,	adParamInput,	 20,	ProductReceiveHp)
												.Parameters.Append .CreateParameter("@ReceiveEmail",		adVarChar,	adParamInput,	 50,	OrderEmail)
												.Parameters.Append .CreateParameter("@ReceiveZipCode",		adVarChar,	adParamInput,	  7,	ProductReceiveZipCode)
												.Parameters.Append .CreateParameter("@ReceiveAddr1",		adVarChar,	adParamInput,	200,	ProductReceiveAddr1)
												.Parameters.Append .CreateParameter("@ReceiveAddr2",		adVarChar,	adParamInput,	200,	ProductReceiveAddr2)
												.Parameters.Append .CreateParameter("@MainFlag",			adChar,		adParamInput,	 1,		ProductMainFlag)
												.Parameters.Append .CreateParameter("@CreateID",			adVarChar,	adParamInput,	 20,	U_NUM)
												.Parameters.Append .CreateParameter("@CreateIP",			adVarChar,	adParamInput,	 15,	U_IP)

												.Execute, , adExecuteNoRecords
										END WITH
										Set oCmd = Nothing
										IF Err.number <> 0 THEN
												oConn.RollbackTrans

												oRs1.Close : oRs.Close
												Set oRs3 = Nothing : Set oRs2 = Nothing : Set oRs1 = Nothing : Set oRs = Nothing
												oConn.Close : Set oConn = Nothing
	
												Response.Write "FAIL|||||주문 처리 도중 오류가 발생하였습니다. [12002]"
												Response.End
										END IF
										'-----------------------------------------------------------------------------------------------------------'
										' 나의 주소록에 주소 추가 시작
										'-----------------------------------------------------------------------------------------------------------'
								END IF


								'-----------------------------------------------------------------------------------------------------------'
								' 1+1상품 체크 시작
								'-----------------------------------------------------------------------------------------------------------'
								IF oRs1("GroupCnt") > 1 THEN
										SET oCmd = Server.CreateObject("ADODB.Command")
										WITH oCmd
												.ActiveConnection	 = oConn
												.CommandType		 = adCmdStoredProc
												.CommandText		 = "USP_Front_EShop_OrderSheet_Select_For_OnePlusOne"

												.Parameters.Append .CreateParameter("@CartID",		adVarChar,	adParamInput, 20,	oRs1("CartID"))
												.Parameters.Append .CreateParameter("@GroupIdx",	adInteger,	adParamInput,   ,	oRs1("GroupIdx"))
										END WITH
										oRs2.CursorLocation = adUseClient
										oRs2.Open oCmd, , adOpenStatic, adLockReadOnly
										SET oCmd = Nothing

										IF NOT oRs2.EOF THEN

												'-----------------------------------------------------------------------------------------------------------'
												'재고 체크 START
												'-----------------------------------------------------------------------------------------------------------'
												'# 매장픽업
												IF oRs2("DelvType") = "S" THEN
														'# ShopCD			= oRs2("PickupShopCD")
														'# 1+1상품은 본상품 픽업매장으로 재고 체크
														'# 재고보유 여부 체크
														Set oCmd = Server.CreateObject("ADODB.Command")
														WITH oCmd
																.ActiveConnection = oConn
																.CommandType = adCmdStoredProc
																.CommandText = "USP_Admin_EShop_Stock_Select_By_Key"

																.Parameters.Append .CreateParameter("@ProductCode",		 adInteger,	 adParamInput,   ,	 oRs2("ProductCode"))
																.Parameters.Append .CreateParameter("@SizeCD",			 adVarChar,	 adParamInput, 50,	 oRs2("SizeCD"))
																.Parameters.Append .CreateParameter("@ShopCD",			 adVarChar,	 adParamInput, 10,	 ShopCD)
														END WITH
														oRs3.CursorLocation = adUseClient
														oRs3.Open oCmd, , adOpenStatic, adLockReadOnly
														Set oCmd = Nothing

														IF NOT oRs3.EOF THEN
																IF oRs3("UseFlag") <> "Y" OR oRs3("SizeCDUseFlag") <> "Y" OR oRs3("SoldOut") = "Y" OR oRs3("RestQty") <= 0 THEN
																		oConn.RollbackTrans

																		oRs3.Close : oRs2.Close : oRs1.Close : oRs.Close
																		Set oRs3 = Nothing : Set oRs2 = Nothing : Set oRs1 = Nothing : Set oRs = Nothing
																		oConn.Close : Set oConn = Nothing
	
																		Response.Write "FAIL|||||재고부족으로 주문 처리 도중 오류가 발생하였습니다. [12111]"
																		Response.End
																END IF
														ELSE
																oConn.RollbackTrans

																oRs3.Close : oRs2.Close : oRs1.Close : oRs.Close
																Set oRs3 = Nothing : Set oRs2 = Nothing : Set oRs1 = Nothing : Set oRs = Nothing
																oConn.Close : Set oConn = Nothing
	
																Response.Write "FAIL|||||재고부족으로 주문 처리 도중 오류가 발생하였습니다. [12112]"
																Response.End
														END IF
														oRs3.Close
												'# 일반택배
												ELSE
														'# 본상품 할당매장 우선 재고보유 여부 체크
														Set oCmd = Server.CreateObject("ADODB.Command")
														WITH oCmd
																.ActiveConnection = oConn
																.CommandType = adCmdStoredProc
																.CommandText = "USP_Admin_EShop_Stock_Select_By_Key"

																.Parameters.Append .CreateParameter("@ProductCode",		 adInteger,	 adParamInput,   ,	 oRs2("ProductCode"))
																.Parameters.Append .CreateParameter("@SizeCD",			 adVarChar,	 adParamInput, 50,	 oRs2("SizeCD"))
																.Parameters.Append .CreateParameter("@ShopCD",			 adVarChar,	 adParamInput, 10,	 ShopCD)
														END WITH
														oRs3.CursorLocation = adUseClient
														oRs3.Open oCmd, , adOpenStatic, adLockReadOnly
														Set oCmd = Nothing

														IF NOT oRs3.EOF THEN
																IF oRs3("UseFlag") <> "Y" OR oRs3("SizeCDUseFlag") <> "Y" OR oRs3("SoldOut") = "Y" OR oRs3("RestQty") <= 0 THEN
																		ShopCD	= ""
																END IF
														ELSE
																ShopCD	= ""
														END IF
														oRs3.Close


														'# 본상품 할당매장에 재고가 없을 경우 재고보유 창고/매장 찾기
														IF ShopCD = "" THEN
																Set oCmd = Server.CreateObject("ADODB.Command")
																WITH oCmd
																		.ActiveConnection = oConn
																		.CommandType = adCmdStoredProc
																		.CommandText = "USP_Admin_EShop_Stock_Select_For_GetShopCD_By_ProductCode_N_SizeCD"

																		.Parameters.Append .CreateParameter("@ProductCode",		 adInteger,	 adParamInput,   ,	 oRs2("ProductCode"))
																		.Parameters.Append .CreateParameter("@SizeCD",			 adVarChar,	 adParamInput, 50,	 oRs2("SizeCD"))
																END WITH
																oRs3.CursorLocation = adUseClient
																oRs3.Open oCmd, , adOpenStatic, adLockReadOnly
																Set oCmd = Nothing

																IF NOT oRs3.EOF THEN
																		ShopCD			= oRs3("ShopCD")
																ELSE
																		oConn.RollbackTrans

																		oRs3.Close : oRs2.Close : oRs1.Close : oRs.Close
																		Set oRs3 = Nothing : Set oRs2 = Nothing : Set oRs1 = Nothing : Set oRs = Nothing
																		oConn.Close : Set oConn = Nothing
	
																		Response.Write "FAIL|||||재고부족으로 주문 처리 도중 오류가 발생하였습니다. [12113]"
																		Response.End
																END IF
																oRs3.Close

																WareHouseType	= GetWareHouseType(ShopCD)
														END IF
												END IF
												'-----------------------------------------------------------------------------------------------------------'
												'재고 체크 END
												'-----------------------------------------------------------------------------------------------------------'

												'-----------------------------------------------------------------------------------------------------------'
												'업체정보 체크 START
												'-----------------------------------------------------------------------------------------------------------'
												Set oCmd = Server.CreateObject("ADODB.Command")
												WITH oCmd
														.ActiveConnection = oConn
														.CommandType = adCmdStoredProc
														.CommandText = "USP_Admin_EShop_Store_Select_By_ShopCD"
														.Parameters.Append .CreateParameter("@ShopCD",		 adChar,	 adParamInput, 6,	 oRs2("ShopCD"))
												END WITH
												oRs3.CursorLocation = adUseClient
												oRs3.Open oCmd, , adOpenStatic, adLockReadOnly
												Set oCmd = Nothing

												IF NOT oRs3.EOF THEN
														'# 입점몰 여부 셋팅
														IF oRs3("OutShopFlag") = "Y" AND oRs3("ShopCD") <> "006740" THEN
																OnlineGB	= "I"
																OutShopCD	= "009899"
														ELSE
																OnlineGB	= "S"
																OutShopCD	= ""
														END IF
												ELSE
														oConn.RollbackTrans

														oRs3.Close : oRs2.Close : oRs1.Close : oRs.Close
														Set oRs3 = Nothing : Set oRs2 = Nothing : Set oRs1 = Nothing : Set oRs = Nothing
														oConn.Close : Set oConn = Nothing
	
														Response.Write "FAIL|||||주문 처리 도중 오류가 발생하였습니다. [12114]"
														Response.End
												END IF
												oRs3.Close
												'-----------------------------------------------------------------------------------------------------------'
												'업체정보 체크 END
												'-----------------------------------------------------------------------------------------------------------'


												'-----------------------------------------------------------------------------------------------------------'
												'EShop_Order_Product 데이터 입력 START
												'-----------------------------------------------------------------------------------------------------------'
												Set oCmd = Server.CreateObject("ADODB.Command")
												WITH oCmd
														.ActiveConnection = oConn
														.CommandType = adCmdStoredProc
														.CommandText = "USP_Front_EShop_Order_Product_Insert"
														.Parameters.Append .CreateParameter("@OrderCode",				adVarChar,	adParamInput,	 20,	OrderCode)
														.Parameters.Append .CreateParameter("@OPIdx_Group",				adInteger,	adParamInput,	   ,	OPIdx_Org)
														.Parameters.Append .CreateParameter("@Vendor",					adVarChar,	adParamInput,	 10,	oRs2("ShopCD"))
														.Parameters.Append .CreateParameter("@ProductCode",				adInteger,	adParamInput,	   ,	oRs2("ProductCode"))
														.Parameters.Append .CreateParameter("@ProductCD",				adVarChar,	adParamInput,	 10,	oRs2("ProductCD"))
														.Parameters.Append .CreateParameter("@ProductName",				adVarChar,	adParamInput,	100,	oRs2("ProductName"))
														.Parameters.Append .CreateParameter("@OrderType",				adChar,		adParamInput,	  1,	oRs2("OrderType"))
														.Parameters.Append .CreateParameter("@ProductType",				adChar,		adParamInput,	  1,	"O")
														.Parameters.Append .CreateParameter("@ProductPoint",			adCurrency,	adParamInput,	   ,	0)
														.Parameters.Append .CreateParameter("@ProdCD",					adVarChar,	adParamInput,	 50,	oRs2("ProdCD"))
														.Parameters.Append .CreateParameter("@ColorCD",					adVarChar,	adParamInput,	100,	oRs2("ColorCD"))
														.Parameters.Append .CreateParameter("@SizeCD",					adVarChar,	adParamInput,	 50,	oRs2("SizeCD"))
														.Parameters.Append .CreateParameter("@OrderCnt",				adInteger,	adParamInput,	   ,	oRs2("OrderCnt"))
														.Parameters.Append .CreateParameter("@TagPrice",				adCurrency,	adParamInput,	   ,	oRs2("TagPrice"))
														.Parameters.Append .CreateParameter("@SalePrice",				adCurrency,	adParamInput,	   ,	0)
														.Parameters.Append .CreateParameter("@OrderPrice",				adCurrency,	adParamInput,	   ,	0)
														.Parameters.Append .CreateParameter("@UseCouponPrice",			adCurrency,	adParamInput,	   ,	0)
														.Parameters.Append .CreateParameter("@UseScashPrice",			adCurrency,	adParamInput,	   ,	0)
														.Parameters.Append .CreateParameter("@UsePointPrice",			adCurrency,	adParamInput,	   ,	0)
														.Parameters.Append .CreateParameter("@DeliveryPrice",			adCurrency,	adParamInput,	   ,	0)
														.Parameters.Append .CreateParameter("@ReceiveName",				adVarChar,	adParamInput,	 50,	ProductReceiveName)
														.Parameters.Append .CreateParameter("@ReceiveTel",				adVarChar,	adParamInput,	 20,	ProductReceiveTel)
														.Parameters.Append .CreateParameter("@ReceiveHp",				adVarChar,	adParamInput,	 20,	ProductReceiveHp)
														.Parameters.Append .CreateParameter("@ReceiveZipCode",			adVarChar,	adParamInput,	  7,	ProductReceiveZipCode)
														.Parameters.Append .CreateParameter("@ReceiveAddr1",			adVarChar,	adParamInput,	200,	ProductReceiveAddr1)
														.Parameters.Append .CreateParameter("@ReceiveAddr2",			adVarChar,	adParamInput,	200,	ProductReceiveAddr2)
														.Parameters.Append .CreateParameter("@Memo",					adVarChar,	adParamInput,	500,	Memo)
														.Parameters.Append .CreateParameter("@OrderState",				adChar,		adParamInput,	  1,	OrderState)
														.Parameters.Append .CreateParameter("@CancelState1",			adChar,		adParamInput,	  1,	CancelState1)
														.Parameters.Append .CreateParameter("@CancelState2",			adChar,		adParamInput,	  1,	CancelState2)
														.Parameters.Append .CreateParameter("@ProductEmployeeFlag",		adChar,		adParamInput,	  1,	"N")
														.Parameters.Append .CreateParameter("@ProductEmployeeType",		adChar,		adParamInput,	  1,	"P")
														.Parameters.Append .CreateParameter("@ProductEmployeeNo",		adVarChar,	adParamInput,	 10,	"")
														.Parameters.Append .CreateParameter("@ProductEmployeeCardNo",	adVarChar,	adParamInput,	 20,	"")
														.Parameters.Append .CreateParameter("@ShopCD",					adVarChar,	adParamInput,	 10,	ShopCD)
														.Parameters.Append .CreateParameter("@WareHouseType",			adChar,		adParamInput,	  1,	WareHouseType)
														.Parameters.Append .CreateParameter("@OnlineGB",				adVarChar,	adParamInput,	 10,	OnlineGB)
														.Parameters.Append .CreateParameter("@OutShopCD",				adVarChar,	adParamInput,	 10,	OutShopCD)
														.Parameters.Append .CreateParameter("@OutletGB",				adChar,		adParamInput,	  1,	oRs2("OutletFlag"))
														.Parameters.Append .CreateParameter("@DelvType",				adChar,		adParamInput,	  1,	oRs2("DelvType"))
	
														.Parameters.Append .CreateParameter("@EventProd1CD",			adVarChar,	adParamInput,	  7,	"")
														.Parameters.Append .CreateParameter("@EventProd1NM",			adVarChar,	adParamInput,	 50,	"")
														.Parameters.Append .CreateParameter("@EventProd1ODQty",			adInteger,	adParamInput,	   ,	0)
														.Parameters.Append .CreateParameter("@EventProd2CD",			adVarChar,	adParamInput,	  7,	"")
														.Parameters.Append .CreateParameter("@EventProd2NM",			adVarChar,	adParamInput,	 50,	"")
														.Parameters.Append .CreateParameter("@EventProd2ODQty",			adInteger,	adParamInput,	   ,	0)
														.Parameters.Append .CreateParameter("@EventProd3CD",			adVarChar,	adParamInput,	  7,	"")
														.Parameters.Append .CreateParameter("@EventProd3NM",			adVarChar,	adParamInput,	 50,	"")
														.Parameters.Append .CreateParameter("@EventProd3ODQty",			adInteger,	adParamInput,	   ,	0)
	
														.Parameters.Append .CreateParameter("@CreateNM",				adVarChar,	adParamInput,	100,	OrderName)
														.Parameters.Append .CreateParameter("@CreateID",				adVarChar,	adParamInput,	 20,	U_NUM)
														.Parameters.Append .CreateParameter("@CreateIP",				adVarChar,	adParamInput,	 15,	U_IP)
														.Parameters.Append .CreateParameter("@OrderSheetIdx",			adInteger,	adParamInput,	   ,	oRs2("Idx"))
														.Parameters.Append .CreateParameter("@OPIdx_Org",				adInteger,	adParamOutput)
			
														.Execute, , adExecuteNoRecords

														OPIdx_Org = .Parameters("@OPIdx_Org").Value
												END WITH
												Set oCmd = Nothing

												IF Err.number <> 0 THEN
														oConn.RollbackTrans

														oRs2.Close : oRs1.Close : oRs.Close
														Set oRs3 = Nothing : Set oRs2 = Nothing : Set oRs1 = Nothing : Set oRs = Nothing
														oConn.Close : Set oConn = Nothing
	
														Response.Write "FAIL|||||주문 처리 도중 오류가 발생하였습니다. [12001]"
														Response.End
												END IF
												'-----------------------------------------------------------------------------------------------------------'
												'EShop_Order_Product 데이터 입력 END
												'-----------------------------------------------------------------------------------------------------------'

										END IF
										oRs2.Close
								END IF
								'-----------------------------------------------------------------------------------------------------------'
								' 1+1상품 체크 끝
								'-----------------------------------------------------------------------------------------------------------'



								ShopOrderCnt		= ShopOrderCnt			+ CDbl(OrderCnt)
								ShopTagPrice		= ShopTagPrice			+ CDbl(TagPrice)
								ShopSalePrice		= ShopSalePrice			+ CDbl(SalePrice)
								ShopOrderPrice		= ShopOrderPrice		+ CDbl(OrderPrice)
								ShopUseCouponPrice	= ShopUseCouponPrice	+ CDbl(UseCouponPrice)
								ShopUsePointPrice	= ShopUsePointPrice		+ CDbl(UsePointPrice)
								ShopUseScashPrice	= ShopUseScashPrice		+ CDbl(UseScashPrice)
								ShopDiscountPrice	= ShopDiscountPrice		+ CDbl(DiscountPrice)
								ShopSavePoint		= ShopSavePoint			+ CDbl(SavePoint)

								oRs1.MoveNext
						Loop 
				END IF
				oRs1.Close

				'# 일반택배배송일 경우 배송비 계산
				IF oRs("DelvType") = "P" AND ShopSalePrice < CDbl(oRs("StandardPrice")) THEN
						ShopDeliveryPrice	= CDbl(oRs("DeliveryPrice"))
				END IF


				'-----------------------------------------------------------------------------------------------------------'
				' 업체별 배송비(EShop_Order_DeliveryPrice) 데이터 입력 START
				'-----------------------------------------------------------------------------------------------------------'
				Set oCmd = Server.CreateObject("ADODB.Command")
				WITH oCmd
						.ActiveConnection = oConn
						.CommandType = adCmdStoredProc
						.CommandText = "USP_Front_EShop_Order_DeliveryPrice_Insert"
						.Parameters.Append .CreateParameter("@OrderCode",			adVarChar,	adParamInput,	 20,	OrderCode)
						.Parameters.Append .CreateParameter("@Vendor",				adVarChar,	adParamInput,	 10,	oRs("ShopCD"))
						.Parameters.Append .CreateParameter("@OPCIdx",				adInteger,	adParamInput,	   ,	0)
						.Parameters.Append .CreateParameter("@SettlePrice",			adCurrency,	adParamInput,	   ,	ShopDeliveryPrice)
						.Parameters.Append .CreateParameter("@RefundPrice",			adCurrency,	adParamInput,	   ,	0)
						.Parameters.Append .CreateParameter("@PayType",				adChar,		adParamInput,	  1,	PayType)
						.Parameters.Append .CreateParameter("@MemberCouponIdx",		adInteger,	adParamInput,	   ,	0)
						.Parameters.Append .CreateParameter("@Memo",				adVarChar,	adParamInput,	500,	"주문시 배송비")
						.Parameters.Append .CreateParameter("@CreateID",			adVarChar,	adParamInput,	 50,	U_NUM)
						.Parameters.Append .CreateParameter("@CreateIP",			adVarChar,	adParamInput,	 15,	U_IP)
			
						.Execute, , adExecuteNoRecords
				END WITH
				Set oCmd = Nothing

				IF Err.number <> 0 THEN
						oConn.RollbackTrans

						oRs.Close
						Set oRs3 = Nothing : Set oRs2 = Nothing : Set oRs1 = Nothing : Set oRs = Nothing
						oConn.Close : Set oConn = Nothing
	
						Response.Write "FAIL|||||주문 처리 도중 오류가 발생하였습니다. [10012]"
						Response.End
				END IF
				'-----------------------------------------------------------------------------------------------------------'
				'업체별 배송비(EShop_Order_DeliveryPrice) 데이터 입력 END
				'-----------------------------------------------------------------------------------------------------------'



				'# 주문금액 계산
				TotalOrderCnt			= TotalOrderCnt			+ ShopOrderCnt
				TotalTagPrice			= TotalTagPrice			+ ShopTagPrice
				TotalSalePrice			= TotalSalePrice		+ ShopSalePrice
				TotalOrderPrice			= TotalOrderPrice		+ ShopOrderPrice
				TotalUseCouponPrice		= TotalUseCouponPrice	+ ShopUseCouponPrice
				TotalUsePointPrice		= TotalUsePointPrice	+ ShopUsePointPrice
				TotalUseScashPrice		= TotalUseScashPrice	+ ShopUseScashPrice
				TotalSavePoint			= TotalSavePoint		+ ShopSavePoint
				TotalDeliveryPrice		= TotalDeliveryPrice	+ ShopDeliveryPrice


				oRs.MoveNext
		Loop
END IF
oRs.Close
'-----------------------------------------------------------------------------------------------------------'
'주문서 테이블 검색 끝
'-----------------------------------------------------------------------------------------------------------'


'-----------------------------------------------------------------------------------------------------------'
' 주문정보 금액정보 데이터 입력 START
'-----------------------------------------------------------------------------------------------------------'
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Front_EShop_Order_Update_For_PriceInfo"
		.Parameters.Append .CreateParameter("@OrderCode",			adVarChar,	adParamInput,	 20,	OrderCode)
		.Parameters.Append .CreateParameter("@TagPrice",			adCurrency,	adParamInput,	   ,	TotalTagPrice)
		.Parameters.Append .CreateParameter("@SalePrice",			adCurrency,	adParamInput,	   ,	TotalSalePrice)
		.Parameters.Append .CreateParameter("@OrderPrice",			adCurrency,	adParamInput,	   ,	TotalOrderPrice + TotalDeliveryPrice)
		.Parameters.Append .CreateParameter("@UseCouponPrice",		adCurrency,	adParamInput,	   ,	TotalUseCouponPrice)
		.Parameters.Append .CreateParameter("@UseScashPrice",		adCurrency,	adParamInput,	   ,	TotalUseScashPrice)
		.Parameters.Append .CreateParameter("@UsePointPrice",		adCurrency,	adParamInput,	   ,	TotalUsePointPrice)
		.Parameters.Append .CreateParameter("@DeliveryPrice",		adCurrency,	adParamInput,	   ,	TotalDeliveryPrice)
			
		.Execute, , adExecuteNoRecords
END WITH
Set oCmd = Nothing

IF Err.number <> 0 THEN
		oConn.RollbackTrans

		Set oRs3 = Nothing : Set oRs2 = Nothing : Set oRs1 = Nothing : Set oRs = Nothing
		oConn.Close : Set oConn = Nothing
	
		Response.Write "FAIL|||||주문 처리 도중 오류가 발생하였습니다. [20001]"
		Response.End
END IF
'-----------------------------------------------------------------------------------------------------------'
' 주문정보 금액정보 데이터 입력 END
'-----------------------------------------------------------------------------------------------------------'






oConn.CommitTrans






Set oRs3 = Nothing
Set oRs2 = Nothing
Set oRs1 = Nothing
Set oRs = Nothing
oConn.Close
Set oConn = Nothing


	

Response.Write "OK||||||||||" & OrderCode
Response.End
%>