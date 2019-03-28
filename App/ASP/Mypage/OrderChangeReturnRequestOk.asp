<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'OrderChangeReturnRequestOk.asp - 주문 교환/반품 신청 처리
'Date		: 2019.01.02
'Update		: 
'*****************************************************************************************'

'# 페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include Virtual = "/ADO/ADODBCommon.asp" -->
<!-- #include Virtual = "/Common/Common.asp" -->
<!-- #include Virtual = "/Common/SetInfo.asp" -->
<!-- #include Virtual = "/Common/ProgID1.asp" -->
<!-- #include Virtual = "/Common/md5.asp" -->



<%
IF U_NUM = "" AND N_NAME = "" THEN
		Call AlertMessage2 ("로그인 정보가 없습니다. 다시 로그인하여 주십시오.", "history.back();")
		Response.End
END IF

IF INSTR(LCASE(HOME_URL), LCASE(Request.ServerVariables("HTTP_HOST")) ) <= 0 THEN
		Call AlertMessage2 ("잘못된 경로로 접근하셨습니다" & Request.ServerVariables("HTTP_HOST"), "history.back();")
		Response.End
END IF




'ON ERROR RESUME NEXT





'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn							'# ADODB Connection 개체
DIM oRs								'# ADODB Recordset 개체
DIM oCmd							'# ADODB Command 개체

'DIM i
DIM j
DIM X

DIM wQuery
DIM sQuery

DIM OrderCode
DIM OPIdx
DIM CancelType
DIM CancelTypeNM
DIM DeliveryCouponIdx
DIM CouponName
DIM ChgSizeCD
DIM SizeChangeFlag
DIM ReasonType
DIM Memo

DIM ReturnName
DIM ReturnHp
DIM ReturnZipCode
DIM ReturnAddr1
DIM ReturnAddr2

DIM ReceiveName
DIM ReceiveHp
DIM ReceiveZipCode
DIM ReceiveAddr1
DIM ReceiveAddr2

DIM DelvFeeType
DIM DelvFee
DIM AddDeliveryPrice		: AddDeliveryPrice	= 0

DIM RefundPrice
DIM RefundBankCode
DIM RefundBankName
DIM RefundAccountNum
DIM RefundAccountName

DIM ProductCode
DIM ProductName
DIM Vendor
DIM OrderCnt
DIM OrderName
DIM OrderHp
DIM OrderEmail
DIM OrderZipCode
DIM OrderAddr1
DIM OrderAddr2
DIM OrderState
DIM CancelState1
DIM CancelState2

DIM OPIdx_Prev
DIM ProdCD
DIM ColorCD
DIM SizeCD
DIM DelvNumber
DIM ShopCD
DIM WareHouseType


DIM TempOPCIdx
DIM OPCIdx

DIM PayType

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
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

OrderCode						 = Trim(sqlFilter(Request("OrderCode")))
OPIdx							 = Trim(sqlFilter(Request("OPIdx")))
CancelType						 = Trim(sqlFilter(Request("CancelType")))
DeliveryCouponIdx				 = Trim(sqlFilter(Request("DeliveryCouponIdx")))
ChgSizeCD						 = Trim(sqlFilter(Request("ChgSizeCD")))
ReasonType						 = Trim(sqlFilter(Request("ReasonType")))
Memo							 = Trim(sqlFilter(Request("Memo")))

ReturnName						 = Trim(sqlFilter(Request("ReturnName")))
ReturnHp						 = Trim(sqlFilter(Request("ReturnHp")))
ReturnZipCode					 = Trim(sqlFilter(Request("ReturnZipCode")))
ReturnAddr1						 = Trim(sqlFilter(Request("ReturnAddr1")))
ReturnAddr2						 = Trim(sqlFilter(Request("ReturnAddr2")))

ReceiveName						 = Trim(sqlFilter(Request("ReceiveName")))
ReceiveHp						 = Trim(sqlFilter(Request("ReceiveHp")))
ReceiveZipCode					 = Trim(sqlFilter(Request("ReceiveZipCode")))
ReceiveAddr1					 = Trim(sqlFilter(Request("ReceiveAddr1")))
ReceiveAddr2					 = Trim(sqlFilter(Request("ReceiveAddr2")))

DelvFeeType						 = Trim(sqlFilter(Request("DelvFeeType")))

RefundPrice						 = Trim(sqlFilter(Request("RefundPrice")))
RefundBankCode					 = Trim(sqlFilter(Request("RefundBankCode")))
RefundAccountNum				 = Trim(sqlFilter(Request("RefundAccountNum")))
RefundAccountName				 = Trim(sqlFilter(Request("RefundAccountName")))



'# Usafe보증보험 관련
'# 배송비 결제는 보증보험을 발급하지 않는다
LGD_ESCROW_USEYN = "N"
GuaranteeInsurance					 = "N"
GuaranteeInsuranceAgreement			 = "N"
USafeJumin1							 = ""
USafeJumin2							 = ""
USafeEmailFlag						 = "N" 
USafeSmsFlag						 = "N" 


IF CancelType = "X" THEN
		CancelTypeNM		= "교환"
ELSEIF CancelType = "R" THEN
		CancelTypeNM		= "반품"
ELSE
		Call AlertMessage2 ("신청구분을 확인해 주십시오.", "history.back()")
		Response.End
END IF


IF RefundPrice	= "" THEN RefundPrice	= "0"


SET oConn	= ConnectionOpen()	'//커넥션 생성
SET oRs		= Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성





'-----------------------------------------------------------------------------------------------------------'
'주문 검색 START
'-----------------------------------------------------------------------------------------------------------'
wQuery = ""
wQuery = wQuery & "WHERE A.IsShowFlag = 'Y' AND A.ProductType IN ('P','O') "
wQuery = wQuery & "AND A.OrderCode = '" & OrderCode & "' "
wQuery = wQuery & "AND A.Idx = " & OPIdx & " "
IF U_NUM <> "" THEN
		wQuery = wQuery & "AND B.UserID = '" & U_NUM & "' "
ELSE
		wQuery = wQuery & "AND B.OrderName = '" & N_NAME & "' AND B.OrderHp = '" & N_HP & "' AND B.OrderEmail = '" & N_EMAIL & "' "
END IF

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
		IF oRs("OrderState") <> "5" OR oRs("CancelState2") <> "0" THEN
				Call AlertMessage2 (CancelTypeNM & "신청이 불가능한 상태의 주문 입니다. 고객센터에 문의해 주십시오.", "history.back();")
				Response.End
		END IF

'#		OrderDate		= oRs("OrderDate")
'#		BrandName		= oRs("BrandName")
'#		OrderPrice		= oRs("OrderPrice")
		ProductCode		= oRs("ProductCode")
		ProductName		= oRs("ProductName")
		Vendor			= oRs("Vendor")
		OrderCnt		= oRs("OrderCnt")
		OrderName		= oRs("OrderName")
		OrderHp			= oRs("OrderHp")
		OrderEmail		= oRs("OrderEmail")
		OrderZipCode	= oRs("OrderZipCode")
		OrderAddr1		= oRs("OrderAddr1")
		OrderAddr2		= oRs("OrderAddr2")
		OrderState		= oRs("OrderState")
		CancelState1	= oRs("CancelState1")
		CancelState2	= oRs("CancelState2")

		OPIdx_Prev		= oRs("OPIdx_Prev")
		ProdCD			= oRs("ProdCD")
		ColorCD			= oRs("ColorCD")
		SizeCD			= oRs("SizeCD")
		DelvNumber		= oRs("DelvNumber")
		ShopCD			= oRs("ShopCD")
		WareHouseType	= oRs("WareHouseType")

		IF CancelType = "R" THEN
				SizeChangeFlag	= "N"
				ChgSizeCD		= oRs("SizeCD")
		ELSE
				IF oRs("SizeCD") <> ChgSizeCD THEN
						SizeChangeFlag	= "Y"
				ELSE
						SizeChangeFlag	= "N"
				END IF
		END IF

		IF CancelType = "X" THEN
				DelvFee			= oRs("VendorDeliveryPrice") * 2		'# 업체 배송비 * 왕복
		ELSE
				DelvFee			= oRs("VendorDeliveryPrice")			'# 업체 배송비
		END IF
ELSE
		Call AlertMessage2 ("주문정보가 없습니다.", "history.back();")
		Response.End
END IF
oRs.Close
'-----------------------------------------------------------------------------------------------------------'
'주문 검색 END
'-----------------------------------------------------------------------------------------------------------'


'-----------------------------------------------------------------------------------------------------------'
'# 반품상품 금액계산 시작
'-----------------------------------------------------------------------------------------------------------'
IF CancelType = "R" THEN
		wQuery = ""
		wQuery = wQuery & "WHERE A.OrderCode = '" & OrderCode & "' "
		wQuery = wQuery & "AND A.ProductType = 'P' "
		wQuery = wQuery & "AND A.Idx = " & OPIdx & " "


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
				'# 배송비 결제내역이 없으면 추가배송비 산정
				IF oRs("DeliveryPrice") = 0 THEN
						DelvFee	= DelvFee * 2
				END IF
		END IF
		oRs.Close
END IF
'-----------------------------------------------------------------------------------------------------------'
'# 반품상품 금액계산 끝
'-----------------------------------------------------------------------------------------------------------'

'-----------------------------------------------------------------------------------------------------------'
'무료배송 쿠폰 유효성 검사 시작
'-----------------------------------------------------------------------------------------------------------'
IF DelvFeeType = "7" THEN
		IF DeliveryCouponIdx = "" THEN
				Call AlertMessage2 ("무료배송 쿠폰을 선택해 주십시오.", "history.back()")
				Response.End
		END IF

		Set oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection = oConn
				.CommandType = adCmdStoredProc
				.CommandText = "USP_Front_EShop_Coupon_Member_Select_By_Idx"

				.Parameters.Append .CreateParameter("@MemberNum",			adInteger,	adParamInput,   ,		U_NUM)
				.Parameters.Append .CreateParameter("@MemberCouponIdx",		adInteger,	adParamInput,   ,		DeliveryCouponIdx)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		Set oCmd = Nothing

		IF NOT oRs.EOF THEN
				CouponName		= oRs("CouponName")

				IF oRs("ReceiveFlag") <> "Y" THEN
						oRs.Close : Set oRs = Nothing
						oConn.Close : Set oConn = Nothing

						Call AlertMessage2("[" & CouponName & "] 쿠폰은 배포되지 않은 쿠폰입니다. [11]", "history.back()")
						Response.End

				ELSEIF oRs("UseFlag") = "Y" THEN
						oRs.Close : Set oRs = Nothing
						oConn.Close : Set oConn = Nothing

						Call AlertMessage2("[" & CouponName & "] 쿠폰은 이미 사용된 쿠폰입니다. [12]", "history.back()")
						Response.End

				ELSEIF oRs("CollectFlag") = "Y" THEN
						oRs.Close : Set oRs = Nothing
						oConn.Close : Set oConn = Nothing

						Call AlertMessage2("[" & CouponName & "] 쿠폰은 회수된 쿠폰입니다. [13]", "history.back()")
						Response.End

				ELSEIF oRs("StartDT") > U_DATE & LEFT(U_TIME,4) THEN
						oRs.Close : Set oRs = Nothing
						oConn.Close : Set oConn = Nothing

						Call AlertMessage2("[" & CouponName & "] 쿠폰은 아직 사용하실 수 없습니다. [14]", "history.back()")
						Response.End

				ELSEIF oRs("EndDT") < U_DATE & LEFT(U_TIME,4) THEN
						oRs.Close : Set oRs = Nothing
						oConn.Close : Set oConn = Nothing

						Call AlertMessage2("[" & CouponName & "] 쿠폰은 사용기한이 지났습니다. [15]", "history.back()")
						Response.End

				ELSEIF oRs("DeliveryCouponFlag") <> "Y" THEN
						oRs.Close : Set oRs = Nothing
						oConn.Close : Set oConn = Nothing

						Call AlertMessage2("[" & CouponName & "] 쿠폰은 무료배송 쿠폰이 아닙니다. [17]", "history.back()")
						Response.End
				END IF
		ELSE
				oRs.Close : Set oRs = Nothing
				oConn.Close : Set oConn = Nothing

				Call AlertMessage2("선택한 쿠폰은 없는 쿠폰입니다. [19]", "history.back()")
				Response.End
		END IF
		oRs.Close
ELSE
		DeliveryCouponIdx = Null
END IF
'-----------------------------------------------------------------------------------------------------------'
'무료배송 쿠폰 유효성 검사 끝
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
ELSE
		RefundBankName		= ""
END IF
'-----------------------------------------------------------------------------------------------------------'
'# 환불은행 체크 End
'-----------------------------------------------------------------------------------------------------------'


'ON ERROR RESUME NEXT


'-----------------------------------------------------------------------------------------------------------'	
'# 주문 교환/반품 요청 내역 Temp 생성 Start
'-----------------------------------------------------------------------------------------------------------'	
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Order_Product_Cancel_Temp_Insert"

		.Parameters.Append .CreateParameter("@OPIdx",			adInteger,	adParamInput,     ,	 OPIdx)
		.Parameters.Append .CreateParameter("@CancelType",		adChar,		adParamInput,    1,	 CancelType)
		.Parameters.Append .CreateParameter("@OrderCode",		adVarChar,	adParamInput,   20,	 OrderCode)
		.Parameters.Append .CreateParameter("@ProductCode",		adInteger,	adParamInput,     ,	 ProductCode)
		.Parameters.Append .CreateParameter("@SizeChangeFlag",	adChar,		adParamInput,    1,	 SizeChangeFlag)
		.Parameters.Append .CreateParameter("@SizeCD",			adVarChar,	adParamInput,   10,	 ChgSizeCD)
		.Parameters.Append .CreateParameter("@OrderCnt",		adInteger,	adParamInput,     ,	 OrderCnt)
		.Parameters.Append .CreateParameter("@ReasonType",		adVarChar,	adParamInput,   10,	 ReasonType)
		.Parameters.Append .CreateParameter("@Memo",			adVarChar,	adParamInput,  255,	 Memo)
		.Parameters.Append .CreateParameter("@DelvFee",			adCurrency,	adParamInput,     ,	 DelvFee)
		.Parameters.Append .CreateParameter("@DelvFeeType",		adChar,		adParamInput,    1,	 DelvFeeType)
		.Parameters.Append .CreateParameter("@DelvFeeMemo",		adVarChar,	adParamInput,  255,	 "")
		.Parameters.Append .CreateParameter("@MemberCouponIdx",	adInteger,	adParamInput,     ,	 DeliveryCouponIdx)
		.Parameters.Append .CreateParameter("@ContactName",		adVarChar,	adParamInput,   50,	 ReturnName)
		.Parameters.Append .CreateParameter("@ContactHp",		adVarChar,	adParamInput,   20,	 ReturnHp)
		.Parameters.Append .CreateParameter("@ReturnName",		adVarChar,	adParamInput,   50,	 ReturnName)
		.Parameters.Append .CreateParameter("@ReturnHp",		adVarChar,	adParamInput,   20,	 ReturnHp)
		.Parameters.Append .CreateParameter("@ReturnZipCode",	adVarChar,	adParamInput,    7,	 ReturnZipCode)
		.Parameters.Append .CreateParameter("@ReturnAddr1",		adVarChar,	adParamInput,  200,	 ReturnAddr1)
		.Parameters.Append .CreateParameter("@ReturnAddr2",		adVarChar,	adParamInput,  200,	 ReturnAddr2)
		.Parameters.Append .CreateParameter("@ReceiveName",		adVarChar,	adParamInput,   50,	 ReceiveName)
		.Parameters.Append .CreateParameter("@ReceiveHp",		adVarChar,	adParamInput,   20,	 ReceiveHp)
		.Parameters.Append .CreateParameter("@ReceiveZipCode",	adVarChar,	adParamInput,    7,	 ReceiveZipCode)
		.Parameters.Append .CreateParameter("@ReceiveAddr1",	adVarChar,	adParamInput,  200,	 ReceiveAddr1)
		.Parameters.Append .CreateParameter("@ReceiveAddr2",	adVarChar,	adParamInput,  200,	 ReceiveAddr2)
		.Parameters.Append .CreateParameter("@RefundPrice",		adCurrency,	adParamInput,     ,	 RefundPrice)
		.Parameters.Append .CreateParameter("@DepositBankCode",	adVarChar,	adParamInput,   10,	 RefundBankCode)
		.Parameters.Append .CreateParameter("@DepositBankName",	adVarChar,	adParamInput,   50,	 RefundBankName)
		.Parameters.Append .CreateParameter("@DepositNumber",	adVarChar,	adParamInput,   50,	 RefundAccountNum)
		.Parameters.Append .CreateParameter("@DepositName",		adVarChar,	adParamInput,   50,	 RefundAccountName)
		.Parameters.Append .CreateParameter("@CreateNM",		adVarChar,	adParamInput,  100,	 OrderName)
		.Parameters.Append .CreateParameter("@CreateID",		adVarChar,	adParamInput,   50,	 U_NUM)
		.Parameters.Append .CreateParameter("@CreateIP",		adVarChar,	adParamInput,   15,	 U_IP)
		.Parameters.Append .CreateParameter("@Idx",				adInteger,	adParamOutput)

		.Execute, , adExecuteNoRecords

		TempOPCIdx = .Parameters("@Idx").Value
END WITH
SET oCmd = Nothing

IF Err.Number <> 0 THEN
		SET oRs = Nothing
		oConn.Close
		SET oConn = Nothing

		Call AlertMessage2(CancelTypeNM & "신청 도중 오류가 발생했습니다. [20]", "history.back()")
		Response.End
END IF
'-----------------------------------------------------------------------------------------------------------'	
'# 주문 교환/반품 요청 내역 Temp 생성 End
'-----------------------------------------------------------------------------------------------------------'	





IF DelvFeeType = "6" THEN
		PayType = "C"
ELSEIF DelvFeeType = "3" THEN
		PayType = "B"
ELSE
		PayType = DelvFeeType
END IF






'# 배송비가 있고 결제수단이 PG결제(신용카드, 계좌이체)일 경우
IF DelvFee > 0 AND (PayType = "C" OR PayType = "B") THEN
		DIM HTTP_USER_AGENT
		DIM USER_AGENT

		HTTP_USER_AGENT = LCase(Request.ServerVariables("HTTP_USER_AGENT"))
		IF InStr(HTTP_USER_AGENT, "android") THEN
				USER_AGENT = "A"
		ELSEIF InStr(HTTP_USER_AGENT, "iphone") OR InStr(HTTP_USER_AGENT, "ipad") OR InStr(HTTP_USER_AGENT, "ipod") THEN
				USER_AGENT = "N"
		ELSE
				USER_AGENT = "Y"
		END IF

		'-----------------------------------------------------------------------------------------------------------'
		'# LGU+ 결제 시작
		'-----------------------------------------------------------------------------------------------------------'
		DIM LGD_CUSTOM_FIRSTPAY
		'//초기 결제 수단 선택
		IF PayType = "C" THEN										'# 카드결제
				LGD_CUSTOM_FIRSTPAY = "SC0010"
		ELSEIF PayType = "B" THEN									'# 계좌이체
				LGD_CUSTOM_FIRSTPAY = "SC0030"
				'# 아이폰에서 계좌이체는 동기방식 지원안함 비동기방식으로 설정
				IF USER_AGENT = "N" THEN
						USER_AGENT = "Y"
				END IF
		ELSEIF PayType = "V" THEN									'# 가상계좌
				LGD_CUSTOM_FIRSTPAY = "SC0040"
		ELSEIF PayType = "M" THEN									'# 모바일결제
				LGD_CUSTOM_FIRSTPAY = "SC0060"
		ELSE														'# 기타 일 경우 카드결제로 셋팅
				LGD_CUSTOM_FIRSTPAY = "SC0010"
		END IF


		DIM LGD_MID
		'IF U_ID = "distance1" THEN
		'		PAY_PLATFORM = "test"
		'END IF
		IF PAY_PLATFORM = "test" THEN								'# 상점아이디(자동생성)
				LGD_MID = CST_MID_TEST                                   
		ELSE
				LGD_MID = CST_MID                                         
		END IF

		'DIM LGD_BUYERADDRESS
		LGD_BUYERADDRESS			 = OrderZipCode & " " & OrderAddr1 & " " & OrderAddr2

		LGD_RECEIVENAME				 = OrderName
		LGD_RECEIVEZIPCODE			 = OrderZipCode
		LGD_RECEIVEADDR1			 = OrderAddr1
		LGD_RECEIVEADDR2			 = OrderAddr2
		LGD_RECEIVEHP				 = OrderHp

		DIM LGD_OID
		DIM LGD_AMOUNT
		DIM LGD_BUYER
		DIM LGD_BUYEREMAIL
		DIM LGD_TIMESTAMP
		DIM LGD_CUSTOM_SKIN
		LGD_OID						 = "OPC" & TempOPCIdx				        '주문번호(상점정의 유니크한 주문번호를 입력하세요)
		LGD_AMOUNT					 = DelvFee									'결제금액("," 를 제외한 결제금액을 입력하세요)
		'LGD_MERTKEY				 = LGD_MERTKEY								'[반드시 세팅]상점MertKey(mertkey는 상점관리자 -> 계약정보 -> 상점정보관리에서 확인하실수 있습니다')
		'LGD_PRODUCTINFO			 = LGD_PRODUCTINFO							'상품명
		LGD_PRODUCTINFO				 = ProductName & " " & CancelTypeNM & " 배송비"
		LGD_BUYER					 = TRIM(OrderName)							'구매자명
		LGD_BUYEREMAIL				 = TRIM(OrderEmail)							'구매자 이메일
		LGD_TIMESTAMP				 = Year(Now) & Right("0" & Month(Now),2) & Right("0" & Day(Now),2) & Right("0" & Hour(Now),2) & Right("0" & Minute(Now),2) & Right("0" & Second(Now),2)		'타임스탬프
		'LGD_CUSTOM_FIRSTPAY         = LGD_CUSTOM_FIRSTPAY						'상점정의 초기결제수단
		LGD_CUSTOM_SKIN				 = "SMART_XPAY2"							'상점정의 결제창 스킨 (red, blue, cyan, green, yellow)
		'# LGD_CUSTOM_SKIN				 = "red"									'상점정의 결제창 스킨 (red, blue, cyan, green, yellow)



		DIM LGD_CASFLAG
		DIM LGD_CASNOTEURL
		DIM LGD_RETURNURL
		DIM LGD_KVPMISPNOTEURL
		DIM LGD_KVPMISPWAPURL
		DIM LGD_KVPMISPCANCELURL

		' * 가상계좌(무통장) 결제 연동을 하시는 경우 아래 LGD_CASNOTEURL 을 설정하여 주시기 바랍니다.
		LGD_CASFLAG					 = "R"
		LGD_CASNOTEURL				 = HOME_DOMAIN & "/ASP/Mypage/OpenXpay/Cas_NoteUrl.asp"			'# MALL_OPENXPAY_CASNOTEURL

		' * LGD_RETURNURL 을 설정하여 주시기 바랍니다. 반드시 현재 페이지와 동일한 프로트콜 및  호스트이어야 합니다. 아래 부분을 반드시 수정하십시요.
		LGD_RETURNURL				 = HOME_DOMAIN & "/ASP/Mypage/OpenXpay/ReturnUrl.asp"			'# MALL_OPENXPAY_RETURNURL
	
		' * ISP 카드결제 연동중 모바일ISP방식(고객세션을 유지하지않는 비동기방식)의 경우, LGD_KVPMISPNOTEURL/LGD_KVPMISPWAPURL/LGD_KVPMISPCANCELURL를 설정하여 주시기 바랍니다. 
		LGD_KVPMISPNOTEURL			 = HOME_DOMAIN & "/ASP/Mypage/OpenXpay/Note_url.asp"
		LGD_KVPMISPWAPURL			 = HOME_DOMAIN & "/ASP/Mypage/OpenXpay/MispwapUrl.asp" & "?LGD_OID=" + LGD_OID    'ISP 카드 결제시, URL 대신 앱명 입력시, 앱호출함 
		LGD_KVPMISPCANCELURL		 = HOME_DOMAIN & "/ASP/Mypage/OpenXpay/Cancel_Url.asp"

		DIM LGD_HASHDATA
		DIM LGD_CUSTOM_PROCESSTYPE
		'/*
		' *************************************************
		' * 2. MD5 해쉬암호화 (수정하지 마세요) - BEGIN
		' *
		' * MD5 해쉬암호화는 거래 위변조를 막기위한 방법입니다.
		' *************************************************
		' *
		' * 해쉬 암호화 적용( LGD_MID + LGD_OID + LGD_AMOUNT + LGD_TIMESTAMP + LGD_MERTKEY )
		' * LGD_MID				: 상점아이디
		' * LGD_OID				: 주문번호
		' * LGD_AMOUNT		: 금액
		' * LGD_TIMESTAMP	: 타임스탬프
		' * LGD_MERTKEY		: 상점MertKey (mertkey는 상점관리자 -> 계약정보 -> 상점정보관리에서 확인하실수 있습니다)
		' *
		' * MD5 해쉬데이터 암호화 검증을 위해
		' * LG유플러스에서 발급한 상점키(MertKey)를 환경설정 파일(lgdacom/conf/mall.conf)에 반드시 입력하여 주시기 바랍니다.
		' */
		LGD_HASHDATA = md5( LGD_MID & LGD_OID & LGD_AMOUNT & LGD_TIMESTAMP & LGD_MERTKEY )
		LGD_CUSTOM_PROCESSTYPE = "TWOTR"
		'/*
		' *************************************************
		' * 2. MD5 해쉬암호화 (수정하지 마세요) - END
		' *************************************************
		' */

		DIM CST_WINDOW_TYPE
		DIM payReqMap

		CST_WINDOW_TYPE = "submit"
		Set payReqMap = Server.CreateObject("Scripting.Dictionary")
		payReqMap.Add "CST_PLATFORM",						 PAY_PLATFORM					'테스트, 서비스 구분
		payReqMap.Add "CST_MID",							 LGD_MID						'상점아이디
		payReqMap.Add "LGD_MID",							 LGD_MID						'상점아이디
		payReqMap.Add "LGD_OID",							 LGD_OID						'주문번호
		payReqMap.Add "LGD_BUYER",							 LGD_BUYER						'구매자
		payReqMap.Add "LGD_PRODUCTINFO",					 LGD_PRODUCTINFO				'상품정보
		payReqMap.Add "LGD_AMOUNT",							 LGD_AMOUNT						'결제금액
		payReqMap.Add "LGD_BUYERID",						 U_NUM							'구매자 아이디
		payReqMap.Add "LGD_BUYERIP",						 U_IP							'구매자 아이디
		payReqMap.Add "LGD_BUYEREMAIL",						 LGD_BUYEREMAIL					'구매자 이메일
		payReqMap.Add "LGD_BUYERPHONE",						 OrderHp						'구매자 휴대번호
		payReqMap.Add "LGD_BUYERADDRESS",					 LGD_BUYERADDRESS				'구매자 주소
		payReqMap.Add "LGD_CUSTOM_SKIN",					 LGD_CUSTOM_SKIN				'결제창 SKIN
		payReqMap.Add "LGD_CUSTOM_PROCESSTYPE",				 LGD_CUSTOM_PROCESSTYPE			'트랜잭션 처리방식
		payReqMap.Add "LGD_TIMESTAMP",						 LGD_TIMESTAMP					'타임스탬프
		payReqMap.Add "LGD_HASHDATA",						 LGD_HASHDATA					'MD5 해쉬암호값
		payReqMap.Add "LGD_VERSION",						 "ASP_SmartXPay_1.0"			'버전정보 (삭제하지 마세요)
		payReqMap.Add "LGD_CUSTOM_FIRSTPAY",				 LGD_CUSTOM_FIRSTPAY			'디폴트 결제수단
		payReqMap.Add "LGD_CUSTOM_USABLEPAY",				 LGD_CUSTOM_FIRSTPAY			'사용가능한 결제 수단
		payReqMap.Add "LGD_CUSTOM_SWITCHINGTYPE",			 "SUBMIT"						'신용카드 카드사 인증 페이지 연동 방식
		payReqMap.Add "LGD_WINDOW_TYPE",					 CST_WINDOW_TYPE				'결제창 호출 방식
		'# payReqMap.Add "LGD_WINDOW_VER",						 "2.5"							'결제창 버젼정보
		'# payReqMap.Add "LGD_OSTYPE_CHECK",					 "P"							'값 P: XPay 실행(PC 결제 모듈): PC용과 모바일용 모듈은 파라미터 및 프로세스가 다르므로 PC용은 PC 웹브라우저에서 실행 필요. "P", "M" 외의 문자(Null, "" 포함)는 모바일 또는 PC 여부를 체크하지 않음

		payReqMap.Add "LGD_RETURNURL",						 LGD_RETURNURL					'응답수신페이지
		'가상계좌(무통장) 결제연동을 하시는 경우  할당/입금 결과를 통보받기 위해 반드시 LGD_CASNOTEURL 정보를 LG 유플러스에 전송해야 합니다 . -->
		payReqMap.Add "LGD_CASFLAG",						 LGD_CASFLAG					'가상계좌 발급/입금/입금취소 상태
		payReqMap.Add "LGD_CASNOTEURL",						 LGD_CASNOTEURL					'가상계좌 NOTEURL


		'****************************************************
		'* 안드로이드폰 신용카드 ISP(국민/BC)결제에만 적용 (시작)*
		'****************************************************
		'(주의)LGD_CUSTOM_ROLLBACK 의 값을  "Y"로 넘길 경우, LG U+ 전자결제에서 보낸 ISP(국민/비씨) 승인정보를 고객서버의 note_url에서 수신시  "OK" 리턴이 안되면  해당 트랜잭션은  무조건 롤백(자동취소)처리되고,
		'LGD_CUSTOM_ROLLBACK 의 값 을 "C"로 넘길 경우, 고객서버의 note_url에서 "ROLLBACK" 리턴이 될 때만 해당 트랜잭션은  롤백처리되며  그외의 값이 리턴되면 정상 승인완료 처리됩니다.
		'만일, LGD_CUSTOM_ROLLBACK 의 값이 "N" 이거나 null 인 경우, 고객서버의 note_url에서  "OK" 리턴이  안될시, "OK" 리턴이 될 때까지 3분간격으로 2시간동안  승인결과를 재전송합니다.
		payReqMap.Add "LGD_CUSTOM_ROLLBACK",			     "C"						 	'비동기 ISP에서 트랜잭션 처리여부

		'아이폰 신용카드 적용  ISP(국민/BC)결제에만 적용 (선택)
		'# payReqMap.Add "LGD_KVPMISPAUTOAPPYN",					 "Y"
		payReqMap.Add "LGD_KVPMISPAUTOAPPYN",					 USER_AGENT
		'Y: 아이폰에서 ISP신용카드 결제시, 고객사에서 'App To App' 방식으로 국민, BC카드사에서 받은 결제 승인을 받고 고객사의 앱을 실행하고자 할때 사용

		IF USER_AGENT = "Y" THEN	'# 비동기 방식일 경우 NOTE_URL, 승인완료 및 취소 URL 설정
				payReqMap.Add "LGD_KVPMISPNOTEURL",						 LGD_KVPMISPNOTEURL			'비동기 ISP(ex. 안드로이드) 승인결과를 받는 URL
				payReqMap.Add "LGD_KVPMISPWAPURL",						 LGD_KVPMISPWAPURL				'비동기 ISP(ex. 안드로이드) 승인완료후 사용자에게 보여지는 승인완료 URL
				payReqMap.Add "LGD_KVPMISPCANCELURL",					 LGD_KVPMISPCANCELURL			'ISP 앱에서 취소시 사용자에게 보여지는 취소 URL
		ELSE								'# 동기 방식일 경우 NOTE_URL, 승인완료 및 취소 URL 설정 안함
				payReqMap.Add "LGD_KVPMISPNOTEURL",						 ""										'비동기 ISP(ex. 안드로이드) 승인결과를 받는 URL
				payReqMap.Add "LGD_KVPMISPWAPURL",						 ""										'비동기 ISP(ex. 안드로이드) 승인완료후 사용자에게 보여지는 승인완료 URL
				payReqMap.Add "LGD_KVPMISPCANCELURL",					 ""										'ISP 앱에서 취소시 사용자에게 보여지는 취소 URL
		END IF
		'****************************************************
		'* 안드로이드폰 신용카드 ISP(국민/BC)결제에만 적용 (끝) *
		'****************************************************

		'# LGD_MTRANSFERWAPURL (계좌이체 승인 완료 후 사용자에게 보여 지는 승인 완료 URL)
		'# LGD_MTRANSFERCANCELURL (계좌이체시 앱에서 취소 시 사용자에게 보여 지는 취소 URL)
		'# LGD_MTRANSFERAUTOAPPYN (계좌이체 앱에서 인증/인증취소 진행 시, 동작 방식을 설정 합니다.)
		'# LGD_MTRANSFERNOTEURL (계좌이체 승인결과를 받는 URL)
		payReqMap.Add "LGD_MTRANSFERAUTOAPPYN",					 USER_AGENT
		IF USER_AGENT = "Y" THEN	'# 비동기 방식일 경우 승인완료 및 취소 URL 설정
				payReqMap.Add "LGD_MTRANSFERWAPURL",					 LGD_KVPMISPWAPURL
				payReqMap.Add "LGD_MTRANSFERCANCELURL",					 LGD_KVPMISPCANCELURL
		ELSE								'# 동기 방식일 경우 승인완료 및 취소 URL 설정 안함
				payReqMap.Add "LGD_MTRANSFERWAPURL",					 ""
				payReqMap.Add "LGD_MTRANSFERCANCELURL",				 ""
		END IF
		payReqMap.Add "LGD_MTRANSFERNOTEURL",					 LGD_KVPMISPNOTEURL


		'수정 불가 ( 인증 후 자동 셋팅 )
		payReqMap.Add "LGD_RESPCODE",						 ""
		payReqMap.Add "LGD_RESPMSG",						 ""
		payReqMap.Add "LGD_PAYKEY",							 ""



		payReqMap.Add "LGD_ENCODING",						 "UTF-8"						'결제창 호출 문자 인코딩방식	EUC-KR	Form submit 방식으로 결제창 호출시 EUC-KR이외의 인코딩을 하는 경우만 사용
		payReqMap.Add "LGD_ENCODING_NOTEURL",				 "UTF-8"						'결과수신페이지 호출 문자 인코딩방식	EUC-KR	UTF-8로 넘기면 UTF-8로 인코딩된 값을 LGD_NOTEURL, LGD_CASNOTEURL 에 전달
		payReqMap.Add "LGD_ENCODING_RETURNURL",				 "UTF-8"						'결과수신페이지 호출 문자 인코딩방식	EUC-KR	UTF-8로 넘기면 UTF-8로 인코딩된 값을 LGD_RETURNURL 에 전달


		'# 현금영수증 발행정보
		payReqMap.Add "LGD_CASHRECEIPTYN",					 "Y"							'현금영수증 발행여부
		payReqMap.Add "LGD_AUTOCOPYYN_CASHCARDNUM",			 "Y"							'현금영수증발급시 발급번호자동채움여부
		payReqMap.Add "LGD_DEFAULTCASHRECEIPTUSE",			 "1"							'현금영수증 발급용도 디폴트 선택

		payReqMap.Add "LGD_CUSTOM_MERTNAME",				 ""								'현금영수증 상점명
		payReqMap.Add "LGD_CUSTOM_MERTPHONE",				 ""								'현금영수증 상점전화번호
		payReqMap.Add "LGD_CUSTOM_BUSINESSNUM",				 ""								'현금영수증 사업자번호
		payReqMap.Add "LGD_CUSTOM_CEONAME",					 ""								'현금영수증 대표자명

		'# 배송지 정보
		payReqMap.Add "LGD_RECEIVER",						 LGD_RECEIVENAME				'수취인
		payReqMap.Add "LGD_RECEIVERPHONE",					 LGD_RECEIVEHP					'수취인 휴대번호
		payReqMap.Add "LGD_DELIVERYINFO",					 LGD_BUYERADDRESS				'배송지 주소

		'# 무통장(가상계좌) 입금 정보
		payReqMap.Add "LGD_CLOSEDATE",						 LGD_CLOSEDATE					'결제가능일시(가상계좌 입금마감일시) yyyyMMddHHmmss 형식

		'# 에스크로 정보
		payReqMap.Add "LGD_ESCROW_USEYN",					 LGD_ESCROW_USEYN				'에스크로 사용여부(매매보호)
		payReqMap.Add "LGD_ESCROW_GOODID",					 "1"							'에스크로 사용여부(매매보호)
		payReqMap.Add "LGD_ESCROW_GOODNAME",				 LGD_PRODUCTINFO				'에스크로 사용여부(매매보호)
		payReqMap.Add "LGD_ESCROW_GOODCODE",				 LGD_OID						'에스크로 사용여부(매매보호)
		payReqMap.Add "LGD_ESCROW_UNITPRICE",				 DelvFee						'에스크로 사용여부(매매보호)
		payReqMap.Add "LGD_ESCROW_QUANTITY",				 "1"							'에스크로 사용여부(매매보호)
		payReqMap.Add "LGD_ESCROW_ZIPCODE",					 LGD_RECEIVEZIPCODE				'에스크로 사용여부(매매보호)
		payReqMap.Add "LGD_ESCROW_ADDRESS1",				 LGD_RECEIVEADDR1				'에스크로 사용여부(매매보호)
		payReqMap.Add "LGD_ESCROW_ADDRESS2",				 LGD_RECEIVEADDR2				'에스크로 사용여부(매매보호)
		payReqMap.Add "LGD_ESCROW_BUYERPHONE",				 LGD_RECEIVEHP					'에스크로 사용여부(매매보호)

		'# 보증보험 정보
		payReqMap.Add "USAFE_GuaranteeInsurance",			 GuaranteeInsurance				'보증보험 발급 여부
		payReqMap.Add "USAFE_GuaranteeInsuranceAgreement",	 GuaranteeInsuranceAgreement	'개인정보 동의 여부
		payReqMap.Add "USAFE_JuminNumber",					 USafeJumin1 & USafeJumin2		'개인정보 동의 여부
		payReqMap.Add "USAFE_EmailFlag",					 USAFEEmailFlag					'Email 동의 여부
		payReqMap.Add "USAFE_SmsFlag",						 USAFESmsFlag					'Sms 동의 여부


		Set Session("PAYREQ_MAP") = payReqMap
		'payReqMap.RemoveAll
		%>

		<html>
		<head>
		<meta http-equiv="Content-Type" content="text/html; charset=EUC-KR">
		<title>LG유플러스 전자결서비스 결제</title>
		<%
		DIM protocol	: protocol = "http"
		IF request.serverVariables("SERVER_PORT") = "443" THEN protocol = "https"

		IF PAY_PLATFORM = "test" THEN
				DIM port : port = "7080"
				IF request.serverVariables("SERVER_PORT") = "443" THEN port = "7443"
				Response.Write "<script language='javascript' src='"& protocol &"://xpay.lgdacom.net:" & port & "/xpay/js/xpay_crossplatform.js' type='text/javascript' ></script>"
		ELSE
				Response.Write "<script language='javascript' src='"& protocol &"://xpay.lgdacom.net/xpay/js/xpay_crossplatform.js' type='text/javascript'></script>"
		END IF
		%>
		<script type="text/javascript">
		<!--
			/*
			* iframe으로 결제창을 호출하시기를 원하시면 iframe으로 설정 (변수명 수정 불가)
			*/
			var LGD_window_type = '<%= CST_WINDOW_TYPE %>';

			/*
			* 수정불가
			*/
			function launchCrossPlatform() {
				lgdwin = open_paymentwindow(document.getElementById('LGD_PAYINFO'), '<%= PAY_PLATFORM %>', LGD_window_type);
				//lgdwin = openXpay(document.getElementById('LGD_PAYINFO'), '<%= PAY_PLATFORM %>', LGD_window_type, null, "", "");
			}

			/*
			* FORM 명만  수정 가능
			*/
			function getFormObject() {
				return document.getElementById("LGD_PAYINFO");
			}

			/*
			 * 인증결과 처리
			 */
			/*
			function payment_return() {
				var fDoc;
				fDoc = lgdwin.contentWindow || lgdwin.contentDocument;
	
				if (fDoc.document.getElementById('LGD_RESPCODE').value == "0000") {
					document.getElementById("LGD_PAYKEY").value = fDoc.document.getElementById('LGD_PAYKEY').value;
					document.getElementById("LGD_PAYINFO").target = "_self";
					document.getElementById("LGD_PAYINFO").action = "/ASP/Mypage/OpenXpay/PayRes.asp";
					document.getElementById("LGD_PAYINFO").submit();
		
				} else {
					alert("LGD_RESPCODE (결과코드) : " + fDoc.document.getElementById('LGD_RESPCODE').value + "\n" + "LGD_RESPMSG (결과메시지): " + fDoc.document.getElementById('LGD_RESPMSG').value);
					closeIframe();
				}
			}
			*/
			window.onload = function () {
				launchCrossPlatform();
			}
		//-->
		</script>
		</head>

		<%IF U_ID = "distance1" THEN%>
		<body>
		<%ELSE%>
		<body oncontextmenu="return false" onselectstart="return false" ondragstart="return false">
		<%END IF%>
		<form method="post" name="LGD_PAYINFO" id="LGD_PAYINFO" action="">
		<%
		DIM EachItem
		FOR EACH EachItem IN payReqMap
				Response.Cookies("PAYREQ_MAP")(EachItem)			 = payReqMap.item(EachItem)
				Response.Write "<input type=""hidden"" name="""& EachItem &""" id="""& EachItem &""" value=""" & payReqMap.item(EachItem) & """><br>"&vbLf
		NEXT
		%>
		</form>
		</body>
		</html>
		<%
		'-----------------------------------------------------------------------------------------------------------'
		'# LGU+ 결제 끝
		'-----------------------------------------------------------------------------------------------------------'


'# 배송비 PG결제가 아닐 경우
ELSE


		oConn.BeginTrans


		'-----------------------------------------------------------------------------------------------------------'	
		'# 주문 교환/반품 신청 등록 Start
		'-----------------------------------------------------------------------------------------------------------'	
		' 1. 주문상품 상태변경
		' 2. 주문상품 변경이력 생성
		' 3. 주문상품 교환/반품 신청 이력 생성
		' 4. 교환/반품 신청 Temp에 OPCIdx 셋팅
		' 5. 업체별 교환/반품 배송비 생성
		' 6. 무료배송쿠폰 사용 처리
		'-----------------------------------------------------------------------------------------------------------'	
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Order_Product_Cancel_Insert_From_Temp"

				.Parameters.Append .CreateParameter("@TempOPCIdx",			adInteger,	adParamInput,   ,	 TempOPCIdx)

				.Execute, , adExecuteNoRecords
		END WITH
		SET oCmd = Nothing

		IF Err.Number <> 0 THEN
				oConn.RollbackTrans

				SET oRs = Nothing
				oConn.Close
				SET oConn = Nothing

				Call AlertMessage2(CancelTypeNM & "신청 도중 오류가 발생했습니다. [21]", "history.back()")
				Response.End
		END IF
		'-----------------------------------------------------------------------------------------------------------'	
		'# 주문 교환/반품 신청 등록 End
		'-----------------------------------------------------------------------------------------------------------'	


		'-----------------------------------------------------------------------------------------------------------'	
		'# 반품회수 신청 등록 Start
		'-----------------------------------------------------------------------------------------------------------'	
		DIM REQUEST_DT
		DIM REQUEST_SN
		DIM WAYBILLNO
		DIM DELPRE_KEY1
		DIM RECEIVE_NM
		DIM RECEIVE_TEL_NO
		DIM RECEIVE_MOBILE_NO
		DIM ZIPCD1
		DIM ZIPCD2
		DIM RECEIVE_ADDR
		DIM PARCELCODE
		DIM TYPECD
		DIM CLIENTCD
		DIM WHCD
		DIM CUSTOMER_RQ
		DIM MANAGER
		DIM MANAGER_RQ
		DIM RETURN_CD
		DIM RETURN_NM

		DIM DELPRE_KEY2
		DIM DELPRE_KEY3
		DIM INTERNALCODE
		'# DIM PRODCD
		'# DIM COLORCD
		'# DIM SIZECD
		DIM RETURN_QTY
		DIM DEFL_FG

		REQUEST_DT			= U_DATE
		WAYBILLNO			= DelvNumber
		DELPRE_KEY1			= OrderCode
		RECEIVE_NM			= ReturnName
		RECEIVE_TEL_NO		= ReturnHp
		RECEIVE_MOBILE_NO	= ReturnHp
		ZIPCD1				= LEFT(ReturnZipCode, LEN(ReturnZipCode) - 3)
		ZIPCD2				= RIGHT(ReturnZipCode, 3)
		RECEIVE_ADDR		= ReturnAddr1 & " " & ReturnAddr2
		PARCELCODE			= "00305"				'# 물류 택배사코드 (00305:CJ대한통운)
		IF WareHouseType = "S" THEN
				TYPECD		= "3"					'# 3: 매장출고
		ELSE
				TYPECD		= "1"					'# 1: 물류출고
		END IF
		CLIENTCD			= ShopCD
		WHCD				= ""
		CUSTOMER_RQ			= ""
		MANAGER				= "슈마커고객센터"
		MANAGER_RQ			= ""
		RETURN_CD			= "00"
		RETURN_NM			= "미등록"


		DELPRE_KEY2			= OPIdx
		IF CStr(OPIdx_Prev) = "0" THEN
				DELPRE_KEY3		= "NORM"
		ELSE
				DELPRE_KEY3		= "CHNORD"
		END IF
		INTERNALCODE		= ""					'# ERP 전송시 처리한다
		'# PRODCD				= ProdCD
		'# COLORCD				= ColorCD
		'# SIZECD				= SizeCD
		RETURN_QTY			= OrderCnt
		DEFL_FG				= "X"					'# X:확인전, N:정상, Y:오배송

		'# 회수 마스터 등록
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Admin_IF_WMS_RETURNREQUEST_H_Insert"

				.Parameters.Append .CreateParameter("@REQUEST_DT",			 adVarChar, adParamInput,   14,	 REQUEST_DT)
				.Parameters.Append .CreateParameter("@WAYBILLNO",			 adVarChar, adParamInput,   50,	 WAYBILLNO)
				.Parameters.Append .CreateParameter("@DELPRE_KEY1",			 adVarChar, adParamInput,   40,	 DELPRE_KEY1)
				.Parameters.Append .CreateParameter("@RECEIVE_NM",			 adVarChar, adParamInput,   40,	 RECEIVE_NM)
				.Parameters.Append .CreateParameter("@RECEIVE_TEL_NO",		 adVarChar, adParamInput,   40,	 RECEIVE_TEL_NO)
				.Parameters.Append .CreateParameter("@RECEIVE_MOBILE_NO",	 adVarChar, adParamInput,   40,	 RECEIVE_MOBILE_NO)
				.Parameters.Append .CreateParameter("@ZIPCD1",				 adVarChar, adParamInput,    3,	 ZIPCD1)
				.Parameters.Append .CreateParameter("@ZIPCD2",				 adVarChar, adParamInput,    3,	 ZIPCD2)
				.Parameters.Append .CreateParameter("@RECEIVE_ADDR",		 adVarChar, adParamInput,  800,	 RECEIVE_ADDR)
				.Parameters.Append .CreateParameter("@PARCELCODE",			 adVarChar, adParamInput,   20,	 PARCELCODE)
				.Parameters.Append .CreateParameter("@TYPECD",				 adVarChar, adParamInput,   10,	 TYPECD)
				.Parameters.Append .CreateParameter("@CLIENTCD",			 adVarChar, adParamInput,   20,	 CLIENTCD)
				.Parameters.Append .CreateParameter("@WHCD",				 adVarChar, adParamInput,   10,	 WHCD)
				.Parameters.Append .CreateParameter("@CUSTOMER_RQ",			 adVarChar, adParamInput,  255,	 CUSTOMER_RQ)
				.Parameters.Append .CreateParameter("@MANAGER",				 adVarChar, adParamInput,   20,	 MANAGER)
				.Parameters.Append .CreateParameter("@MANAGER_RQ",			 adVarChar, adParamInput,  255,	 MANAGER_RQ)
				.Parameters.Append .CreateParameter("@RETURN_CD",			 adVarChar, adParamInput,    5,	 RETURN_CD)
				.Parameters.Append .CreateParameter("@RETURN_NM",			 adVarChar, adParamInput,   50,	 RETURN_NM)
				.Parameters.Append .CreateParameter("@INSERT_DT",			 adVarChar, adParamInput,   14,	 U_DATE & U_TIME)
				.Parameters.Append .CreateParameter("@REQUEST_STATE",		 adVarChar, adParamInput,    1,	 "0")				'# 상태 (0:요청, 1:수신)
				.Parameters.Append .CreateParameter("@CreateID",			 adVarChar, adParamInput,   20,	 U_NUM)
				.Parameters.Append .CreateParameter("@CreateIP",			 adVarChar, adParamInput,   15,	 U_IP)
				.Parameters.Append .CreateParameter("@REQUEST_SN",			 adInteger, adParamOutput)

				.Execute, , adExecuteNoRecords

				REQUEST_SN = .Parameters("@REQUEST_SN").Value
		END WITH
		SET oCmd = Nothing

		IF Err.Number <> 0 THEN
				oConn.RollbackTrans
				SET oRs = Nothing : oConn.Close : SET oConn = Nothing

				Call AlertMessage2(CancelTypeNM & "신청 도중 오류가 발생했습니다. [31]", "history.back()")
				Response.End
		END IF


		'# 회수요청 상세 정보 등록
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Admin_IF_WMS_RETURNREQUEST_D_Insert"

				.Parameters.Append .CreateParameter("@REQUEST_DT",			 adVarChar, adParamInput,   14,	 REQUEST_DT)
				.Parameters.Append .CreateParameter("@REQUEST_SN",			 adInteger, adParamInput,     ,	 REQUEST_SN)
				.Parameters.Append .CreateParameter("@WAYBILLNO",			 adVarChar, adParamInput,   50,	 WAYBILLNO)
				.Parameters.Append .CreateParameter("@DELPRE_KEY1",			 adVarChar, adParamInput,   40,	 DELPRE_KEY1)
				.Parameters.Append .CreateParameter("@DELPRE_KEY2",			 adVarChar, adParamInput,   10,	 DELPRE_KEY2)
				.Parameters.Append .CreateParameter("@DELPRE_KEY3",			 adVarChar, adParamInput,   40,	 DELPRE_KEY3)
				.Parameters.Append .CreateParameter("@INTERNALCODE",		 adVarChar, adParamInput,   50,	 INTERNALCODE)
				.Parameters.Append .CreateParameter("@PRODCD",				 adVarChar, adParamInput,   20,	 PRODCD)
				.Parameters.Append .CreateParameter("@COLORCD",				 adVarChar, adParamInput,  100,	 COLORCD)
				.Parameters.Append .CreateParameter("@SIZECD",				 adVarChar, adParamInput,   20,	 SIZECD)
				.Parameters.Append .CreateParameter("@RETURN_QTY",			 adInteger, adParamInput,     ,	 RETURN_QTY)
				.Parameters.Append .CreateParameter("@DEFL_FG_IG",			 adVarChar, adParamInput,    5,	 DEFL_FG)
				.Parameters.Append .CreateParameter("@INSERT_DT",			 adVarChar, adParamInput,   14,	 U_DATE & U_TIME)
				.Parameters.Append .CreateParameter("@REQUEST_STATE",		 adVarChar, adParamInput,    1,	 "0")				'# 상태 (0:요청, 1:수신)
				.Parameters.Append .CreateParameter("@OPIdx",				 adInteger, adParamInput,     ,	 OPIdx)
				.Parameters.Append .CreateParameter("@CreateID",			 adVarChar, adParamInput,   20,	 U_NUM)
				.Parameters.Append .CreateParameter("@CreateIP",			 adVarChar, adParamInput,   15,	 U_IP)

				.Execute, , adExecuteNoRecords
		END WITH
		SET oCmd = Nothing

		IF Err.Number <> 0 THEN
				oConn.RollbackTrans
				SET oRs = Nothing : oConn.Close : SET oConn = Nothing

				Call AlertMessage2(CancelTypeNM & "신청 도중 오류가 발생했습니다. [32]", "history.back()")
				Response.End
		END IF


		'# 회수요청으로 인한 주문 변경이력 생성
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Admin_EShop_Order_Product_Change_History_Insert"

				.Parameters.Append .CreateParameter("@OPIdx",		 adInteger,	 adParamInput,     ,	 OPIdx)
				.Parameters.Append .CreateParameter("@Contents",	 adVarChar,	 adParamInput, 8000,	 "물류 회수 요청")
				.Parameters.Append .CreateParameter("@CreateNM",	 adVarChar,	 adParamInput,  100,	 U_NAME)
				.Parameters.Append .CreateParameter("@CreateID",	 adVarChar,	 adParamInput,   50,	 U_NUM)
				.Parameters.Append .CreateParameter("@CreateIP",	 adVarChar,	 adParamInput,   20,	 U_IP)

				.Execute, , adExecuteNoRecords
		END WITH
		SET oCmd = Nothing

		IF Err.Number <> 0 THEN
				oConn.RollbackTrans
				SET oRs = Nothing : oConn.Close : SET oConn = Nothing

				Call AlertMessage2(CancelTypeNM & "신청 도중 오류가 발생했습니다. [33]", "history.back()")
				Response.End
		END IF
		'-----------------------------------------------------------------------------------------------------------'	
		'# 반품회수 신청 등록 End
		'-----------------------------------------------------------------------------------------------------------'	


		oConn.CommitTrans


		'-----------------------------------------------------------------------------------------------------------'	
		'문자발송 시작
		'-----------------------------------------------------------------------------------------------------------'	
		DIM SmsCode
		IF CancelType = "X" THEN
				SmsCode		= "ORD_S591"		'# 교환신청
		ELSEIF CancelType = "R" THEN
				SmsCode		= "ORD_S581"		'# 반품신청
		END IF
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Admin_EShop_Order_Sms_Send"

				.Parameters.Append .CreateParameter("@OrderCode",	 adVarChar,	 adParamInput,   20,	 OrderCode)
				.Parameters.Append .CreateParameter("@OPIdx",		 adInteger,	 adParamInput,     ,	 OPIdx)
				.Parameters.Append .CreateParameter("@SmsCode",		 adVarChar,	 adParamInput,   20,	 SmsCode)

				.Execute, , adExecuteNoRecords
		END WITH
		SET oCmd = Nothing
		'-----------------------------------------------------------------------------------------------------------'	
		'문자발송 끝
		'-----------------------------------------------------------------------------------------------------------'	


		Call AlertMessage2(CancelTypeNM & "신청 되었습니다.", "location.replace('/ASP/Mypage/OrderList.asp');")

END IF
%>

<%
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>