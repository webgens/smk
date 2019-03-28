<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'Cas_Note_Url.asp - 가상계좌 결제 입금/입금취소시 처리 페이지
'Date		: 2018.12.30
'Update		: 
'*****************************************************************************************'

'# 페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'-----------------------------------------------------------------------------------------'
%>



<!-- #include virtual="/ADO/ADODBCommon_NOHttps.asp" -->
<!-- #include Virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/OpenXpay/lgdacom/md5.asp" -->

<%
'ON ERROR RESUME NEXT
'/*
' * [상점 결제결과처리(DB) 페이지]
' *
' * 1) 위변조 방지를 위한 hashdata값 검증은 반드시 적용하셔야 합니다.
' *
' */
'/*
'* 공통결제결과 정보 
'*/
DIM LGD_RESPCODE				'# 응답코드: 0000(성공) 그외 실패
DIM LGD_RESPMSG					'# 응답메세지
DIM LGD_MID						'# 상점아이디 
DIM LGD_OID						'# 주문번호
DIM LGD_AMOUNT					'# 거래금액
DIM LGD_TID						'# LG텔레콤이 부여한 거래번호
DIM LGD_PAYTYPE					'# 결제수단코드
DIM LGD_PAYDATE					'# 거래일시(승인일시/이체일시)
DIM LGD_HASHDATA				'# 해쉬값
DIM LGD_HASHDATA2				'# 해쉬값
DIM LGD_FINANCECODE				'# 결제기관코드(카드종류/은행코드/이통사코드)
DIM LGD_FINANCENAME				'# 결제기관이름(카드이름/은행이름/이통사이름)
DIM LGD_ESCROWYN				'# 에스크로 적용여부
DIM LGD_TIMESTAMP				'# 타임스탬프
DIM LGD_FINANCEAUTHNUM			'# 결제기관 승인번호(신용카드, 계좌이체, 상품권)

DIM LGD_CARDNUM
DIM LGD_CARDINSTALLMONTH
DIM LGD_CARDNOINTYN				'# 무이자할부여부(신용카드) - '1'이면 무이자할부 '0'이면 일반할부
DIM LGD_PCANCELFLAG				'# 0: 부분취소불가능,  1: 부분취소가능
DIM LGD_PCANCELSTR				'# 부분취소가능시는 "0" 으로 리턴
'/*
'* 가상계좌 결제결과 정보
'*/
DIM LGD_ACCOUNTNUM				'# 계좌번호(무통장입금)
DIM LGD_ACCOUNTOWNER			'# 계좌주명
DIM LGD_CASTAMOUNT				'# 입금총액(무통장입금)
DIM LGD_CASCAMOUNT				'# 현입금액(무통장입금)
DIM LGD_CASFLAG					'# 무통장입금 플래그(무통장입금) - 'R':계좌할당, 'I':입금, 'C':입금취소
DIM LGD_CASSEQNO				'# 입금순서(무통장입금)
DIM LGD_CASHRECEIPTNUM			'# 현금영수증 승인번호
DIM LGD_CASHRECEIPTSELFYN		'# 현금영수증자진발급제유무 Y: 자진발급제 적용, 그외 : 미적용
DIM LGD_CASHRECEIPTKIND			'# 현금영수증 종류 0: 소득공제용 , 1: 지출증빙용
DIM LGD_PAYER					'# 입금자명
DIM LGD_SAOWNER					'# 가상계좌 입금계좌주명.상점명이 디폴트로 리턴

DIM LGD_TELNO					'# 모바일 결제 휴대폰번호

'/*
'* 구매정보
'*/
DIM LGD_BUYER					'# 구매자
DIM LGD_PRODUCTINFO				'# 상품명
DIM LGD_BUYERID					'# 구매자 ID
DIM LGD_BUYERADDRESS			'# 구매자 주소
DIM LGD_BUYERPHONE				'# 구매자 전화번호
DIM LGD_BUYEREMAIL				'# 구매자 이메일
DIM LGD_BUYERSSN				'# 구매자 주민번호
DIM LGD_PRODUCTCODE				'# 상품코드
DIM LGD_RECEIVER				'# 수취인
DIM LGD_RECEIVERPHONE			'# 수취인 전화번호
DIM LGD_DELIVERYINFO			'# 배송지

DIM resultMSG					'# 결과처리 메시지
DIM LGD_CUSTOM_SMSMSG

DIM ReceiptFlag		: ReceiptFlag	= "N"
DIM CouponCode		: CouponCode	= ""

DIM OrderState
DIM SettleFlag
DIM SettleDate
DIM SettleTime
DIM CasFlag
DIM PayType

'USafe 보증보험 관련
DIM USAFE_GuaranteeInsurance
DIM USAFE_GuaranteeInsuranceAgreement
DIM USAFE_JuminNumber
DIM USAFE_EmailFlag
DIM USAFE_SmsFlag
	

LGD_RESPCODE							= Trim(Request("LGD_RESPCODE"))					'# 응답코드: 0000(성공) 그외 실패
LGD_RESPMSG								= Trim(Request("LGD_RESPMSG"))					'# 응답메세지
LGD_MID									= Trim(Request("LGD_MID"))						'# 상점아이디
LGD_OID									= Trim(Request("LGD_OID"))						'# 주문번호
LGD_AMOUNT								= Trim(Request("LGD_AMOUNT"))					'# 거래금액
LGD_TID									= Trim(Request("LGD_TID"))						'# LG텔레콤이 부여한 거래번호
LGD_PAYTYPE								= Trim(Request("LGD_PAYTYPE"))					'# 결제수단코드
LGD_PAYDATE								= Trim(Request("LGD_PAYDATE"))					'# 거래일시(승인일시/이체일시)
LGD_HASHDATA							= Trim(Request("LGD_HASHDATA"))					'# 해쉬값
LGD_FINANCECODE							= Trim(Request("LGD_FINANCECODE"))				'# 결제기관코드(은행코드)
LGD_FINANCENAME							= Trim(Request("LGD_FINANCENAME"))				'# 결제기관이름(은행이름)
LGD_ESCROWYN							= Trim(Request("LGD_ESCROWYN"))					'# 에스크로 적용여부
LGD_TIMESTAMP							= Trim(Request("LGD_TIMESTAMP"))				'# 타임스탬프

LGD_ACCOUNTNUM							= Trim(Request("LGD_ACCOUNTNUM"))				'# 계좌번호(무통장입금)
LGD_ACCOUNTOWNER						= Trim(Request("LGD_ACCOUNTOWNER"))				'# 계좌주명	
LGD_CASTAMOUNT							= Trim(Request("LGD_CASTAMOUNT"))				'# 입금총액(무통장입금)
LGD_CASCAMOUNT							= Trim(Request("LGD_CASCAMOUNT"))				'# 현입금액(무통장입금)
LGD_CASFLAG								= Trim(Request("LGD_CASFLAG"))					'# 무통장입금 플래그(무통장입금) - 'R':계좌할당, 'I':입금, 'C':입금취소
LGD_CASSEQNO							= Trim(Request("LGD_CASSEQNO"))					'# 입금순서(무통장입금)
LGD_CASHRECEIPTNUM						= Trim(Request("LGD_CASHRECEIPTNUM"))			'# 현금영수증 승인번호
LGD_CASHRECEIPTSELFYN					= Trim(Request("LGD_CASHRECEIPTSELFYN"))		'# 현금영수증자진발급제유무 Y: 자진발급제 적용, 그외 : 미적용
LGD_CASHRECEIPTKIND						= Trim(Request("LGD_CASHRECEIPTKIND"))			'# 현금영수증 종류 0: 소득공제용 , 1: 지출증빙용
LGD_PAYER								= Trim(Request("LGD_PAYER"))					'# 입금자명
LGD_SAOWNER								= Trim(Request("LGD_SAOWNER"))					'# 입금자명

'/*
' * 구매정보
' */
LGD_BUYER								= Trim(Request("LGD_BUYER"))					'# 구매자
LGD_PRODUCTINFO							= Trim(Request("LGD_PRODUCTINFO"))				'# 상품명
LGD_BUYERID								= Trim(Request("LGD_BUYERID"))					'# 구매자 ID
LGD_BUYERADDRESS						= Trim(Request("LGD_BUYERADDRESS"))				'# 구매자 주소
LGD_BUYERPHONE							= Trim(Request("LGD_BUYERPHONE"))				'# 구매자 전화번호
LGD_BUYEREMAIL							= Trim(Request("LGD_BUYEREMAIL"))				'# 구매자 이메일
LGD_BUYERSSN							= Trim(Request("LGD_BUYERSSN"))					'# 구매자 주민번호
LGD_PRODUCTCODE							= Trim(Request("LGD_PRODUCTCODE"))				'# 상품코드
LGD_RECEIVER							= Trim(Request("LGD_RECEIVER"))					'# 수취인
LGD_RECEIVERPHONE						= Trim(Request("LGD_RECEIVERPHONE"))			'# 수취인 전화번호
LGD_DELIVERYINFO						= Trim(Request("LGD_DELIVERYINFO"))				'# 배송지
	
CouponCode								= Trim(Request("CouponCode"))

'USafe 보증보험 관련
USAFE_GuaranteeInsurance				= Trim(Request("USAFE_GuaranteeInsurance"))
USAFE_GuaranteeInsuranceAgreement		= Trim(Request("USAFE_GuaranteeInsuranceAgreement"))
USAFE_JuminNumber						= Trim(Request("USAFE_JuminNumber"))
USAFE_EmailFlag							= Trim(Request("USAFE_EmailFlag"))
USAFE_SmsFlag							= Trim(Request("USAFE_SmsFlag"))


IF USAFE_GuaranteeInsurance				= "" THEN USAFE_GuaranteeInsurance				= "N"
IF USAFE_GuaranteeInsuranceAgreement	= "" THEN USAFE_GuaranteeInsuranceAgreement		= "N"
IF USAFE_EmailFlag						= "" THEN USAFE_EmailFlag						= "N"
IF USAFE_SmsFlag						= "" THEN USAFE_SmsFlag							= "N"

'/*
' * hashdata 검증을 위한 mertkey는 상점관리자 -> 계약정보 -> 상점정보관리에서 확인하실수 있습니다.
' * LG텔레콤에서 발급한 상점키로 반드시변경해 주시기 바랍니다.
' */
'LGD_MERTKEY = ""  '//mertkey	--shindongjoo 2011.07.19 변경됨	SetInfo.asp 파일에 정의되어 있음
	
LGD_HASHDATA2				 = md5( LGD_MID & LGD_OID & LGD_AMOUNT & LGD_RESPCODE & LGD_TIMESTAMP & LGD_MERTKEY )

IF LGD_CASHRECEIPTNUM <> "" THEN
		ReceiptFlag = "Y"
END IF

IF LGD_ESCROWYN = "" OR IsNull(LGD_ESCROWYN) THEN LGD_ESCROWYN = "N"

'# LGD_OID			= "C0001000016"
'# LGD_RESPCODE	= "0000"
'# LGD_CASFLAG		= "I"
'# LGD_AMOUNT		= 59000 
'# LGD_CASTAMOUNT	= 59000
'# LGD_PAYTYPE		= "SC0040"
'# LGD_HASHDATA	= LGD_HASHDATA2

SELECT CASE LGD_PAYTYPE
		CASE "SC0010" : PayType = "C"
		CASE "SC0030" : PayType = "B"
		CASE "SC0040" : PayType = "V"
		CASE "SC0060" : PayType = "M"
END SELECT


'/*
' * 상점 처리결과 리턴메세지
' *
' * OK  : 상점 처리결과 성공
' * 그외 : 상점 처리결과 실패
' *
' * ※ 주의사항 : 성공시 'OK' 문자이외의 다른문자열이 포함되면 실패처리 되오니 주의하시기 바랍니다.
' */


DIM oConn						'# ADODB Connection 개체
DIM oRs							'# ADODB Recordset 개체
DIM oRs1						'# ADODB Recordset 개체
DIM oCmd						'# ADODB Command 개체

SET oConn	= ConnectionOpen()							'//커넥션 생성
SET oRs		= Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성
SET oRs1	= Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성



'-----------------------------------------------------------------------------------------'
'//주문아이디(LGD_OID)에 해당하는 아이디를 검색
'-----------------------------------------------------------------------------------------'
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Admin_EShop_Order_Select_By_OrderCode"

		.Parameters.Append .CreateParameter("@OrderCode",		adVarChar,	adParamInput,	20,		LGD_OID)
END WITH
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing
	
IF NOT oRs.EOF THEN
		U_NUM			= oRs("UserID")
		U_NAME			= oRs("OrderName")
END IF
oRs.Close


'-----------------------------------------------------------------------------------------'
'주문 상태 확인
'-----------------------------------------------------------------------------------------'
DIM DB_OrderState

Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Admin_EShop_Order_Product_Select_By_OrderCode"

		.Parameters.Append .CreateParameter("@OrderCode",		adVarChar,	adParamInput,	20,		LGD_OID)
END WITH
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing

IF NOT oRs.EOF THEN
		DB_OrderState	= oRs("OrderState")
End IF
oRs.Close


resultMSG = "결제결과 상점 DB처리(LGD_CASNOTEURL) 결과값을 입력해 주시기 바랍니다."
	


'# resultMSG = LGD_CASFLAG & "|" & LGD_OID & "|가상계좌 입금오류."
'# Response.Write resultMSG
'# Response.End


IF LGD_HASHDATA2 = LGD_HASHDATA THEN

		'//해쉬값 검증이 성공이면
		IF LGD_RESPCODE = "0000" THEN


				'// 결제가 성공이면
				IF LGD_CASFLAG = "R" THEN				'# 계좌 할당 : 결제완료시 PayRes.asp에서 처리

						'-----------------------------------------------------------------------------------------'
						'주문 상태 확인(미결제상태) OrderState = "1"
						'-----------------------------------------------------------------------------------------'
						IF Trim(DB_OrderState) = "1" THEN
								resultMSG = "OK"   
								Response.Write resultMSG
								Response.End 

						ELSE
								resultMSG = "주문상태가 오류 입니다. 다시 주문 하여 주십시오."
								Response.Write resultMSG
								Response.End
						END IF

								
				ELSEIF LGD_CASFLAG = "I" THEN			'# 입금
						'/*
						' * 무통장 입금 성공 결과 상점 처리(DB) 부분
						' * 상점 결과 처리가 정상이면 "OK"
						' */

						'-----------------------------------------------------------------------------------------'
						'주문 상태 확인(미결제상태) OrderState = "1"
						'-----------------------------------------------------------------------------------------'
						IF Trim(DB_OrderState) <> "1" THEN
								resultMSG = "주문상태가 오류 입니다."
								Response.Write resultMSG
								Response.End
						END IF


						oConn.BeginTrans	

				
						'# 입금 누적액이 결제금액보다 크거나 같을 때 결제완료 처리
						'# Response.Write "LGD_AMOUNT=" & LGD_AMOUNT & " / LGD_CASTAMOUNT=" & LGD_CASTAMOUNT
						IF LGD_AMOUNT <= LGD_CASTAMOUNT THEN
								OrderState		= "3"
								SettleFlag		= "Y"
								SettleDate		= U_DATE
								SettleTime		= U_TIME
								CasFlag			= LGD_CASFLAG
						ELSE
								OrderState		= "1"
								SettleFlag		= "N"
								SettleDate		= U_DATE
								SettleTime		= U_TIME
								CasFlag			= LGD_CASFLAG
						END IF

						'-----------------------------------------------------------------------------------------------------------'	
						'EShop_Order  업데이트 START
						'-----------------------------------------------------------------------------------------------------------'	
						Set oCmd = Server.CreateObject("ADODB.Command")
						WITH oCmd
								.ActiveConnection = oConn
								.CommandType = adCmdStoredProc
								.CommandText = "USP_Front_EShop_Order_Update_For_SettleState"

								.Parameters.Append .CreateParameter("@OrderCode",			adVarChar,	adParamInput,	 20,		LGD_OID)
								.Parameters.Append .CreateParameter("@OrderState",			adChar,		adParamInput,	  1,		OrderState)
								.Parameters.Append .CreateParameter("@SettleFlag",			adChar,		adParamInput,	  1,		SettleFlag)
								.Parameters.Append .CreateParameter("@SettleDate",			adChar,		adParamInput,	  8,		SettleDate)
								.Parameters.Append .CreateParameter("@SettleTime",			adChar,		adParamInput,	  6,		SettleTime)
								.Parameters.Append .CreateParameter("@ReceiptFlag",			adChar,		adParamInput,	  1,		ReceiptFlag)
								.Parameters.Append .CreateParameter("@ReceiptKind",			adChar,		adParamInput,	  1,		LGD_CASHRECEIPTKIND)
								.Parameters.Append .CreateParameter("@EscrowFlag",			adChar,		adParamInput,	  1,		LGD_ESCROWYN)
								.Parameters.Append .CreateParameter("@CasFlag",				adChar,		adParamInput,	  1,		CasFlag)
								.Parameters.Append .CreateParameter("@PayType",				adChar,		adParamInput,	  1,		PayType)
								.Parameters.Append .CreateParameter("@UpdateNM",			adVarChar,	adParamInput,	100,		U_NAME)
								.Parameters.Append .CreateParameter("@UpdateID",			adVarChar,	adParamInput,	 20,		U_NUM)
								.Parameters.Append .CreateParameter("@UpdateIP",			adVarChar,	adParamInput,	 15,		U_IP)
			
								.Execute, , adExecuteNoRecords
						END WITH
						Set oCmd = Nothing

						IF Err.number <> 0 THEN
								oConn.RollbackTrans

								Set oRs1 = Nothing
								Set oRs = Nothing
								oConn.Close
								Set oConn = Nothing

								resultMSG = "FAIL_02-1"
								Response.Write resultMSG
								Response.End
						END IF
						'-----------------------------------------------------------------------------------------------------------'	
						'EShop_Order  업데이트 End
						'-----------------------------------------------------------------------------------------------------------'



						'-----------------------------------------------------------------------------------------------------------'	
						'결제 정보 저장 START
						'-----------------------------------------------------------------------------------------------------------'	
						Set oCmd = Server.CreateObject("ADODB.Command")
						WITH oCmd
								.ActiveConnection = oConn
								.CommandType = adCmdStoredProc
								.CommandText = "USP_Front_EShop_Order_Settle_Update_For_NBank"
								.Parameters.Append .CreateParameter("@OrderCode",					adVarChar,	adParamInput,	 20,	LGD_OID)
								.Parameters.Append .CreateParameter("@LGD_RESPCODE",				adVarChar,	adParamInput,	  4,	LGD_RESPCODE)
								.Parameters.Append .CreateParameter("@LGD_RESPMSG",					adVarChar,	adParamInput,	512,	LGD_RESPMSG)
								.Parameters.Append .CreateParameter("@LGD_AMOUNT",					adVarChar,	adParamInput,	 12,	LGD_AMOUNT)
								.Parameters.Append .CreateParameter("@LGD_MID",						adVarChar,	adParamInput,	 15,	LGD_MID)
								.Parameters.Append .CreateParameter("@LGD_TID",						adVarChar,	adParamInput,	 24,	LGD_TID)
								.Parameters.Append .CreateParameter("@LGD_OID",						adVarChar,	adParamInput,	 64,	LGD_OID)
								.Parameters.Append .CreateParameter("@LGD_TIMESTAMP",				adVarChar,	adParamInput,	 14,	LGD_TIMESTAMP)
								.Parameters.Append .CreateParameter("@LGD_PAYTYPE",					adVarChar,	adParamInput,	  6,	LGD_PAYTYPE)
								.Parameters.Append .CreateParameter("@LGD_PAYDATE",					adVarChar,	adParamInput,	 14,	LGD_PAYDATE)
								.Parameters.Append .CreateParameter("@LGD_HASHDATA",				adVarChar,	adParamInput,	512,	LGD_HASHDATA)
								.Parameters.Append .CreateParameter("@LGD_FINANCECODE",				adVarChar,	adParamInput,	 50,	LGD_FINANCECODE)
								.Parameters.Append .CreateParameter("@LGD_FINANCENAME",				adVarChar,	adParamInput,	 20,	LGD_FINANCENAME)
								.Parameters.Append .CreateParameter("@LGD_ESCROWYN",				adVarChar,	adParamInput,	  1,	LGD_ESCROWYN)
								.Parameters.Append .CreateParameter("@LGD_CASHRECEIPTNUM",			adVarChar,	adParamInput,	 10,	LGD_CASHRECEIPTNUM)
								.Parameters.Append .CreateParameter("@LGD_CASHRECEIPTSELFYN",		adVarChar,	adParamInput,	  1,	LGD_CASHRECEIPTSELFYN)
								.Parameters.Append .CreateParameter("@LGD_CASHRECEIPTKIND",			adVarChar,	adParamInput,	  1,	LGD_CASHRECEIPTKIND)
								.Parameters.Append .CreateParameter("@LGD_ACCOUNTNUM",				adVarChar,	adParamInput,	 20,	LGD_ACCOUNTNUM)
								.Parameters.Append .CreateParameter("@LGD_ACCOUNTOWNER",			adVarChar,	adParamInput,	 40,	LGD_ACCOUNTOWNER)
								.Parameters.Append .CreateParameter("@LGD_PAYER",					adVarChar,	adParamInput,	 40,	LGD_PAYER)
								.Parameters.Append .CreateParameter("@LGD_CASTAMOUNT",				adVarChar,	adParamInput,	 12,	LGD_CASTAMOUNT)
								.Parameters.Append .CreateParameter("@LGD_CASCAMOUNT",				adVarChar,	adParamInput,	 12,	LGD_CASCAMOUNT)
								.Parameters.Append .CreateParameter("@LGD_CASFLAG",					adVarChar,	adParamInput,	 10,	LGD_CASFLAG)
								.Parameters.Append .CreateParameter("@LGD_CASSEQNO",				adVarChar,	adParamInput,	  3,	LGD_CASSEQNO)
								.Parameters.Append .CreateParameter("@LGD_SAOWNER",					adVarChar,	adParamInput,	 40,	LGD_SAOWNER)
								.Parameters.Append .CreateParameter("@UpdateID",					adVarChar,	adParamInput,	 50,	U_NUM)
								.Parameters.Append .CreateParameter("@UpdateIP",					adVarChar,	adParamInput,	 15,	U_IP)

								.Execute, , adExecuteNoRecords
						END WITH
						Set oCmd = Nothing
				
						IF Err.number <> 0 THEN
								oConn.RollbackTrans

								Set oRs1 = Nothing
								Set oRs = Nothing
								oConn.Close
								Set oConn = Nothing

								resultMSG =  "FAIL_02-2"
								Response.Write resultMSG
								Response.End	
						END IF
						'-----------------------------------------------------------------------------------------------------------'	
						'결제 정보 저장 End
						'-----------------------------------------------------------------------------------------------------------'	


						'-----------------------------------------------------------------------------------------------------------'	
						'ERP 전송용 I/F 주문 생성 START
						'-----------------------------------------------------------------------------------------------------------'	
						'//입금완료(입금 누적액이 결제금액보다 크거나 같을 때) 실행
						IF LGD_AMOUNT <= LGD_CASTAMOUNT THEN

								Set oCmd = Server.CreateObject("ADODB.Command")
								WITH oCmd
										.ActiveConnection = oConn
										.CommandType = adCmdStoredProc
										.CommandText = "USP_Admin_EShop_Order_Product_Select_By_OrderCode"

										.Parameters.Append .CreateParameter("@OrderCode",		adVarChar,	adParamInput,	20,		LGD_OID)
								END WITH
								oRs.Open oCmd, , adOpenStatic, adLockReadOnly
								Set oCmd = Nothing

								IF NOT oRs.EOF THEN
										Do Until oRs.EOF
												'# 예약상품이 아닌 경우만 ERP 전송
												IF oRs("OrderType") <> "R" THEN

														'# 일반택배 배송일 경우만 재고체크
														IF oRs("DelvType") = "P" THEN

																'-----------------------------------------------------------------------------------------------------------'
																'재고 체크 START
																'-----------------------------------------------------------------------------------------------------------'
																DIM ShopCD
																DIM WareHouseType

																ShopCD				= oRs("ShopCD")
																WareHouseType		= oRs("WareHouseType")

																'# 주문시 할당매장/창고에 재고보유 여부 체크
																Set oCmd = Server.CreateObject("ADODB.Command")
																WITH oCmd
																		.ActiveConnection = oConn
																		.CommandType = adCmdStoredProc
																		.CommandText = "USP_Admin_EShop_Stock_Select_By_Key"

																		.Parameters.Append .CreateParameter("@ProductCode",		 adInteger,	 adParamInput,   ,	 oRs("ProductCode"))
																		.Parameters.Append .CreateParameter("@SizeCD",			 adVarChar,	 adParamInput, 50,	 oRs("SizeCD"))
																		.Parameters.Append .CreateParameter("@ShopCD",			 adVarChar,	 adParamInput, 10,	 ShopCD)
																END WITH
																oRs1.CursorLocation = adUseClient
																oRs1.Open oCmd, , adOpenStatic, adLockReadOnly
																Set oCmd = Nothing

																IF NOT oRs1.EOF THEN
																		IF oRs1("UseFlag") <> "Y" OR oRs1("SizeCDUseFlag") <> "Y" OR oRs1("SoldOut") = "Y" OR oRs1("RestQty") <= 0 THEN
																				ShopCD				= ""
																				WareHouseType		= ""
																		END IF
																ELSE
																		ShopCD				= ""
																		WareHouseType		= ""
																END IF
																oRs1.Close


																'# 주문시 할당된 매장/창고에 재고가 없을 경우 재할당
																IF ShopCD = "" THEN
																		'# 재고보유 창고/매장 찾기
																		Set oCmd = Server.CreateObject("ADODB.Command")
																		WITH oCmd
																				.ActiveConnection = oConn
																				.CommandType = adCmdStoredProc
																				.CommandText = "USP_Admin_EShop_Stock_Select_For_GetShopCD_By_ProductCode_N_SizeCD"

																				.Parameters.Append .CreateParameter("@ProductCode",		 adInteger,	 adParamInput,   ,	 oRs("ProductCode"))
																				.Parameters.Append .CreateParameter("@SizeCD",			 adVarChar,	 adParamInput, 50,	 oRs("SizeCD"))
																		END WITH
																		oRs1.CursorLocation = adUseClient
																		oRs1.Open oCmd, , adOpenStatic, adLockReadOnly
																		Set oCmd = Nothing

																		IF NOT oRs1.EOF THEN
																				ShopCD				= oRs1("ShopCD")
																				WareHouseType		= GetWareHouseType(ShopCD)
																		ELSE
																				'# 재고가 없을 경우 기존 할당된 정보로 처리
																				ShopCD				= oRs("ShopCD")
																				WareHouseType		= oRs("WareHouseType")
																		END IF
																		oRs1.Close
																END IF

																'# 재할당된 경우 주문정보에 UPDATE
																IF ShopCD <> oRs("ShopCD") THEN
																		SET oCmd = Server.CreateObject("ADODB.Command")
																		WITH oCmd
																				.ActiveConnection	 = oConn
																				.CommandType		 = adCmdStoredProc
																				.CommandText		 = "USP_Admin_EShop_Order_Product_Update_For_ShopCD"

																				.Parameters.Append .CreateParameter("@Idx",				 adInteger,	 adParamInput,     ,	 oRs("Idx"))
																				.Parameters.Append .CreateParameter("@ShopCD",			 adVarChar,	 adParamInput,   10,	 ShopCD)
																				.Parameters.Append .CreateParameter("@WareHouseType",	 adChar,	 adParamInput,    1,	 WareHouseType)
																				.Parameters.Append .CreateParameter("@UpdateID",		 adVarChar,	 adParamInput,   20,	 U_NUM)
																				.Parameters.Append .CreateParameter("@UpdateIP",		 adVarChar,	 adParamInput,   15,	 U_IP)

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

																				resultMSG =  "FAIL_02-3"
																				Response.Write resultMSG
																				Response.End
																		END IF
																END IF
																'-----------------------------------------------------------------------------------------------------------'
																'재고 체크 END
																'-----------------------------------------------------------------------------------------------------------'
														END IF



														'# 주문/결제 생성전송
														SET oCmd = Server.CreateObject("ADODB.Command")
														WITH oCmd
																.ActiveConnection	 = oConn
																.CommandType		 = adCmdStoredProc
																.CommandText		 = "USP_Admin_IF_ONLINE_ORDER_Insert_With_IF_ONLINE_ORDER_APP"

																.Parameters.Append .CreateParameter("@Idx",			 adInteger,	 adParamInput,     ,	 oRs("Idx"))
																.Parameters.Append .CreateParameter("@DOCTYPECD",	 adVarChar,	 adParamInput,   40,	 "NORM")
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

																resultMSG =  "FAIL_02-4"
																Response.Write resultMSG
																Response.End
														END IF
												END IF

												oRs.MoveNext
										Loop 
								End IF
								oRs.Close
						END IF
						'-----------------------------------------------------------------------------------------------------------'	
						'ERP 전송용 I/F 주문 생성 End
						'-----------------------------------------------------------------------------------------------------------'	


						oConn.CommitTrans



						'//입금완료(입금 누적액이 결제금액보다 크거나 같을 때) 실행
						IF LGD_AMOUNT <= LGD_CASTAMOUNT THEN
								'-----------------------------------------------------------------------------------------------------------'	
								'문자발송
								'-----------------------------------------------------------------------------------------------------------'	
								'# Server.Execute("/Common/SMS/OrderSmsSend.asp")
								SET oCmd = Server.CreateObject("ADODB.Command")
								WITH oCmd
										.ActiveConnection	 = oConn
										.CommandType		 = adCmdStoredProc
										.CommandText		 = "USP_Admin_EShop_Order_Sms_Send"

										.Parameters.Append .CreateParameter("@OrderCode",	 adVarChar,	 adParamInput,   20,	 LGD_OID)
										.Parameters.Append .CreateParameter("@OPIdx",		 adInteger,	 adParamInput,     ,	 0)
										.Parameters.Append .CreateParameter("@SmsCode",		 adVarChar,	 adParamInput,   20,	 "ORD_S300")

										.Execute, , adExecuteNoRecords
								END WITH
								SET oCmd = Nothing
								'-----------------------------------------------------------------------------------------------------------'	
								'문자발송
								'-----------------------------------------------------------------------------------------------------------'	
						END IF
					

						Set oRs1 = Nothing
						Set oRs = Nothing
						oConn.Close
						Set oConn = Nothing

						resultMSG = "OK"	
						Response.Write resultMSG

						Response.End
				

				ELSEIF LGD_CASFLAG = "C" THEN			'# 입금취소
				
						'/*
						' * 무통장 입금취소 성공 결과 상점 처리(DB) 부분
						' * 상점 결과 처리가 정상이면 "OK"
						' */
				
						oConn.BeginTrans	

						'-----------------------------------------------------------------------------------------------------------'	
						'EShop_Order  업데이트 START
						'-----------------------------------------------------------------------------------------------------------'	
						'# OrderState		= "2"
						'# SettleFlag		= "N"
						'# SettleDate		= ""
						'# SettleTime		= ""
						'# CasFlag			= LGD_CASFLAG
						'# 
						'# Set oCmd = Server.CreateObject("ADODB.Command")
						'# WITH oCmd
						'# 		.ActiveConnection = oConn
						'# 		.CommandType = adCmdStoredProc
						'# 		.CommandText = "USP_Front_EShop_Order_Update_For_SettleState"
						'# 
						'# 		.Parameters.Append .CreateParameter("@OrderCode",			adVarChar,	adParamInput,	20,		LGD_OID)
						'# 		.Parameters.Append .CreateParameter("@OrderState",			adChar,		adParamInput,	 1,		OrderState)
						'# 		.Parameters.Append .CreateParameter("@SettleFlag",			adChar,		adParamInput,	 1,		SettleFlag)
						'# 		.Parameters.Append .CreateParameter("@SettleDate",			adChar,		adParamInput,	 8,		SettleDate)
						'# 		.Parameters.Append .CreateParameter("@SettleTime",			adChar,		adParamInput,	 6,		SettleTime)
						'# 		.Parameters.Append .CreateParameter("@ReceiptFlag",			adChar,		adParamInput,	 1,		ReceiptFlag)
						'# 		.Parameters.Append .CreateParameter("@ReceiptKind",			adChar,		adParamInput,	 1,		LGD_CASHRECEIPTKIND)
						'# 		.Parameters.Append .CreateParameter("@EscrowFlag",			adChar,		adParamInput,	 1,		LGD_ESCROWYN)
						'# 		.Parameters.Append .CreateParameter("@CasFlag",				adChar,		adParamInput,	 1,		CasFlag)
						'# 		.Parameters.Append .CreateParameter("@PayType",				adChar,		adParamInput,	 1,		PayType)
						'# 		.Parameters.Append .CreateParameter("@UpdateID",			adVarChar,	adParamInput,	20,		U_NUM)
						'# 		.Parameters.Append .CreateParameter("@UpdateIP",			adVarChar,	adParamInput,	15,		U_IP)
						'# 
						'# 		.Execute, , adExecuteNoRecords
						'# END WITH
						'# Set oCmd = Nothing
						'# 
						'# IF Err.number <> 0 THEN
						'# 		oConn.RollbackTrans
						'# 
						'# 		Set oRs1 = Nothing
						'# 		Set oRs = Nothing
						'# 		oConn.Close
						'# 		Set oConn = Nothing
						'# 
						'# 		resultMSG = "FAIL_03-1"
						'# 		Response.Write resultMSG
						'# 		Response.End
						'# END IF
						'-----------------------------------------------------------------------------------------------------------'	
						'EShop_Order  업데이트 End
						'-----------------------------------------------------------------------------------------------------------'
				
						'-----------------------------------------------------------------------------------------------------------'
						'결제 정보 저장 START
						'-----------------------------------------------------------------------------------------------------------'	
						Set oCmd = Server.CreateObject("ADODB.Command")
						WITH oCmd
								.ActiveConnection = oConn
								.CommandType = adCmdStoredProc
								.CommandText = "USP_Front_EShop_Order_Settle_Insert"
								.Parameters.Append .CreateParameter("@OrderCode",					adVarChar,	adParamInput,	 20,	LGD_OID)
								.Parameters.Append .CreateParameter("@LGD_RESPCODE",				adVarChar,	adParamInput,	  4,	LGD_RESPCODE)
								.Parameters.Append .CreateParameter("@LGD_RESPMSG",					adVarChar,	adParamInput,	512,	LGD_RESPMSG)
								.Parameters.Append .CreateParameter("@LGD_AMOUNT",					adVarChar,	adParamInput,	 12,	LGD_AMOUNT)
								.Parameters.Append .CreateParameter("@LGD_MID",						adVarChar,	adParamInput,	 15,	LGD_MID)
								.Parameters.Append .CreateParameter("@LGD_TID",						adVarChar,	adParamInput,	 24,	LGD_TID)
								.Parameters.Append .CreateParameter("@LGD_OID",						adVarChar,	adParamInput,	 64,	LGD_OID)
								.Parameters.Append .CreateParameter("@LGD_TIMESTAMP",				adVarChar,	adParamInput,	 14,	LGD_TIMESTAMP)
								.Parameters.Append .CreateParameter("@LGD_PAYTYPE",					adVarChar,	adParamInput,	  6,	LGD_PAYTYPE)
								.Parameters.Append .CreateParameter("@LGD_PAYDATE",					adVarChar,	adParamInput,	 14,	LGD_PAYDATE)
								.Parameters.Append .CreateParameter("@LGD_HASHDATA",				adVarChar,	adParamInput,	512,	LGD_HASHDATA)
								.Parameters.Append .CreateParameter("@LGD_FINANCECODE",				adVarChar,	adParamInput,	 50,	LGD_FINANCECODE)
								.Parameters.Append .CreateParameter("@LGD_FINANCENAME",				adVarChar,	adParamInput,	 20,	LGD_FINANCENAME)
								.Parameters.Append .CreateParameter("@LGD_FINANCEAUTHNUM",			adVarChar,	adParamInput,	 20,	"")
								.Parameters.Append .CreateParameter("@LGD_CARDNUM",					adVarChar,	adParamInput,	 30,	"")
								.Parameters.Append .CreateParameter("@LGD_CARDINSTALLMONTH",		adVarChar,	adParamInput,	  2,	"")
								.Parameters.Append .CreateParameter("@LGD_CARDNOINTYN",				adVarChar,	adParamInput,	  1,	"")
								.Parameters.Append .CreateParameter("@LGD_PCANCELFLAG",				adVarChar,	adParamInput,	  1,	"")
								.Parameters.Append .CreateParameter("@LGD_PCANCELSTR",				adVarChar,	adParamInput,	128,	"")
								.Parameters.Append .CreateParameter("@LGD_ESCROWYN",				adVarChar,	adParamInput,	  1,	LGD_ESCROWYN)
								.Parameters.Append .CreateParameter("@LGD_CASHRECEIPTNUM",			adVarChar,	adParamInput,	 10,	LGD_CASHRECEIPTNUM)
								.Parameters.Append .CreateParameter("@LGD_CASHRECEIPTSELFYN",		adVarChar,	adParamInput,	  1,	LGD_CASHRECEIPTSELFYN)
								.Parameters.Append .CreateParameter("@LGD_CASHRECEIPTKIND",			adVarChar,	adParamInput,	  1,	LGD_CASHRECEIPTKIND)
								.Parameters.Append .CreateParameter("@LGD_ACCOUNTNUM",				adVarChar,	adParamInput,	 20,	LGD_ACCOUNTNUM)
								.Parameters.Append .CreateParameter("@LGD_ACCOUNTOWNER",			adVarChar,	adParamInput,	 40,	LGD_ACCOUNTOWNER)
								.Parameters.Append .CreateParameter("@LGD_PAYER",					adVarChar,	adParamInput,	 40,	LGD_PAYER)
								.Parameters.Append .CreateParameter("@LGD_CASTAMOUNT",				adVarChar,	adParamInput,	 12,	LGD_CASTAMOUNT)
								.Parameters.Append .CreateParameter("@LGD_CASCAMOUNT",				adVarChar,	adParamInput,	 12,	LGD_CASCAMOUNT)
								.Parameters.Append .CreateParameter("@LGD_CASFLAG",					adVarChar,	adParamInput,	 10,	LGD_CASFLAG)
								.Parameters.Append .CreateParameter("@LGD_CASSEQNO",				adVarChar,	adParamInput,	  3,	LGD_CASSEQNO)
								.Parameters.Append .CreateParameter("@LGD_SAOWNER",					adVarChar,	adParamInput,	 40,	LGD_SAOWNER)
								.Parameters.Append .CreateParameter("@LGD_TELNO",					adVarChar,	adParamInput,	 40,	"")
								.Parameters.Append .CreateParameter("@CreateID",					adVarChar,	adParamInput,	 50,	U_NUM)
								.Parameters.Append .CreateParameter("@CreateIP",					adVarChar,	adParamInput,	 15,	U_IP)

								.Execute, , adExecuteNoRecords
						END WITH
						Set oCmd = Nothing

						IF Err.number <> 0 THEN
								oConn.RollbackTrans

								Set oRs1 = Nothing
								Set oRs = Nothing
								oConn.Close
								Set oConn = Nothing

								resultMSG = "FAIL_03-2"
								Response.Write resultMSG
								Response.End
						END IF
						'-----------------------------------------------------------------------------------------------------------'
						'결제 정보 저장 End
						'-----------------------------------------------------------------------------------------------------------'


						oConn.CommitTrans

						Set oRs1 = Nothing
						Set oRs = Nothing
						oConn.Close
						Set oConn = Nothing

						resultMSG = "OK"
						Response.Write resultMSG
						Response.End
				
				END IF

		ELSE
				'-----------------------------------------------------------------------------------------------------------'
				'//결제가 실패이면
				'/*
				' * 거래실패 결과 상점 처리(DB) 부분
				' * 상점결과 처리가 정상이면 "OK"
				' */
				'-----------------------------------------------------------------------------------------------------------'


				oConn.BeginTrans


				'-----------------------------------------------------------------------------------------------------------'
				'결제 정보 저장 START
				'-----------------------------------------------------------------------------------------------------------'	
				Set oCmd = Server.CreateObject("ADODB.Command")
				WITH oCmd
						.ActiveConnection = oConn
						.CommandType = adCmdStoredProc
						.CommandText = "USP_Front_EShop_Order_Settle_Insert"
						.Parameters.Append .CreateParameter("@OrderCode",					adVarChar,	adParamInput,	 20,	LGD_OID)
						.Parameters.Append .CreateParameter("@LGD_RESPCODE",				adVarChar,	adParamInput,	  4,	LGD_RESPCODE)
						.Parameters.Append .CreateParameter("@LGD_RESPMSG",					adVarChar,	adParamInput,	512,	LGD_RESPMSG)
						.Parameters.Append .CreateParameter("@LGD_AMOUNT",					adVarChar,	adParamInput,	 12,	LGD_AMOUNT)
						.Parameters.Append .CreateParameter("@LGD_MID",						adVarChar,	adParamInput,	 15,	LGD_MID)
						.Parameters.Append .CreateParameter("@LGD_TID",						adVarChar,	adParamInput,	 24,	LGD_TID)
						.Parameters.Append .CreateParameter("@LGD_OID",						adVarChar,	adParamInput,	 64,	LGD_OID)
						.Parameters.Append .CreateParameter("@LGD_TIMESTAMP",				adVarChar,	adParamInput,	 14,	LGD_TIMESTAMP)
						.Parameters.Append .CreateParameter("@LGD_PAYTYPE",					adVarChar,	adParamInput,	  6,	LGD_PAYTYPE)
						.Parameters.Append .CreateParameter("@LGD_PAYDATE",					adVarChar,	adParamInput,	 14,	LGD_PAYDATE)
						.Parameters.Append .CreateParameter("@LGD_HASHDATA",				adVarChar,	adParamInput,	512,	LGD_HASHDATA)
						.Parameters.Append .CreateParameter("@LGD_FINANCECODE",				adVarChar,	adParamInput,	 50,	LGD_FINANCECODE)
						.Parameters.Append .CreateParameter("@LGD_FINANCENAME",				adVarChar,	adParamInput,	 20,	LGD_FINANCENAME)
						.Parameters.Append .CreateParameter("@LGD_FINANCEAUTHNUM",			adVarChar,	adParamInput,	 20,	"")
						.Parameters.Append .CreateParameter("@LGD_CARDNUM",					adVarChar,	adParamInput,	 30,	"")
						.Parameters.Append .CreateParameter("@LGD_CARDINSTALLMONTH",		adVarChar,	adParamInput,	  2,	"")
						.Parameters.Append .CreateParameter("@LGD_CARDNOINTYN",				adVarChar,	adParamInput,	  1,	"")
						.Parameters.Append .CreateParameter("@LGD_PCANCELFLAG",				adVarChar,	adParamInput,	  1,	"")
						.Parameters.Append .CreateParameter("@LGD_PCANCELSTR",				adVarChar,	adParamInput,	128,	"")
						.Parameters.Append .CreateParameter("@LGD_ESCROWYN",				adVarChar,	adParamInput,	  1,	LGD_ESCROWYN)
						.Parameters.Append .CreateParameter("@LGD_CASHRECEIPTNUM",			adVarChar,	adParamInput,	 10,	LGD_CASHRECEIPTNUM)
						.Parameters.Append .CreateParameter("@LGD_CASHRECEIPTSELFYN",		adVarChar,	adParamInput,	  1,	LGD_CASHRECEIPTSELFYN)
						.Parameters.Append .CreateParameter("@LGD_CASHRECEIPTKIND",			adVarChar,	adParamInput,	  1,	LGD_CASHRECEIPTKIND)
						.Parameters.Append .CreateParameter("@LGD_ACCOUNTNUM",				adVarChar,	adParamInput,	 20,	LGD_ACCOUNTNUM)
						.Parameters.Append .CreateParameter("@LGD_ACCOUNTOWNER",			adVarChar,	adParamInput,	 40,	LGD_ACCOUNTOWNER)
						.Parameters.Append .CreateParameter("@LGD_PAYER",					adVarChar,	adParamInput,	 40,	LGD_PAYER)
						.Parameters.Append .CreateParameter("@LGD_CASTAMOUNT",				adVarChar,	adParamInput,	 12,	LGD_CASTAMOUNT)
						.Parameters.Append .CreateParameter("@LGD_CASCAMOUNT",				adVarChar,	adParamInput,	 12,	LGD_CASCAMOUNT)
						.Parameters.Append .CreateParameter("@LGD_CASFLAG",					adVarChar,	adParamInput,	 10,	LGD_CASFLAG)
						.Parameters.Append .CreateParameter("@LGD_CASSEQNO",				adVarChar,	adParamInput,	  3,	LGD_CASSEQNO)
						.Parameters.Append .CreateParameter("@LGD_SAOWNER",					adVarChar,	adParamInput,	 40,	LGD_SAOWNER)
						.Parameters.Append .CreateParameter("@LGD_TELNO",					adVarChar,	adParamInput,	 40,	"")
						.Parameters.Append .CreateParameter("@CreateID",					adVarChar,	adParamInput,	 50,	U_NUM)
						.Parameters.Append .CreateParameter("@CreateIP",					adVarChar,	adParamInput,	 15,	U_IP)

						.Execute, , adExecuteNoRecords
				END WITH
				Set oCmd = Nothing
				
				IF Err.number <> 0 THEN
						oConn.RollbackTrans

						Set oRs1 = Nothing
						Set oRs = Nothing
						oConn.Close
						Set oConn = Nothing

						resultMSG = "FAIL"	
						Response.Write resultMSG
						Response.End
				END IF
				'-----------------------------------------------------------------------------------------------------------'
				'결제 정보 저장 End
				'-----------------------------------------------------------------------------------------------------------'

				oConn.CommitTrans

				Set oRs1 = Nothing
				Set oRs = Nothing
				oConn.Close
				Set oConn = Nothing

				resultMSG = "OK"
				Response.Write resultMSG
				Response.End
		END IF

ELSE
		'//해쉬값이 검증이 실패이면
		'/*
		' * hashdata검증 실패 로그를 처리하시기 바랍니다.
		' */
		
		Set oRs1 = Nothing
		Set oRs = Nothing
		oConn.Close
		Set oConn = Nothing

		resultMSG = "결제결과 상점 DB처리(LGD_CASNOTEURL) 해쉬값 검증이 실패하였습니다."
		Response.Write resultMSG
		Response.End
END IF
%>

