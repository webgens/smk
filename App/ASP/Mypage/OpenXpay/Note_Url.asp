<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'Note_Url.asp - 카드결제 ISP / 계좌이체 결과 처리 페이지
'Date		: 2019.01.03
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
<!-- #include virtual="/Common/ProgID1.asp" -->
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

'/*
'* 신용카드 결제결과 정보
'*/
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

DIM DelvFee
DIM PayType

DIM CancelType
DIM OrderCode
DIM OPIdx
DIM OPIdx_Prev
DIM ProdCD
DIM ColorCD
DIM SizeCD
DIM OrderCnt
DIM DelvNumber
DIM ShopCD
DIM WareHouseType
DIM ReturnName
DIM ReturnHp
DIM ReturnZipCode
DIM ReturnAddr1
DIM ReturnAddr2

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
DIM oCmd						'# ADODB Command 개체

SET oConn	= ConnectionOpen()							'//커넥션 생성
SET oRs		= Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성



'-----------------------------------------------------------------------------------------'
'# 교환/반품 요청정보 검색
'-----------------------------------------------------------------------------------------'
DIM DB_PayType

Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Front_EShop_Order_Product_Cancel_Temp_Select_By_Idx"

		.Parameters.Append .CreateParameter("@Idx",		adInteger,	adParamInput,	,		Replace(LGD_OID, "OPC", ""))
END WITH
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing
	
IF NOT oRs.EOF THEN
		U_NUM				= oRs("CreateID")
		U_NAME				= oRs("CreateNM")
		SELECT CASE oRs("DelvFeeType")
				CASE "6"	: DB_PayType	= "C"
				CASE "3"	: DB_PayType	= "B"
				CASE ELSE	: DB_PayType	= oRs("DelvFeeType")
		END SELECT
		DelvFee				= oRs("DelvFee")

		CancelType			= oRs("CancelType")
		OrderCode			= oRs("OrderCode")
		OPIdx				= oRs("OPIdx")
		OPIdx_Prev			= oRs("OPIdx_Prev")
		ProdCD				= oRs("ProdCD")
		ColorCD				= oRs("ColorCD")
		SizeCD				= oRs("SizeCD")
		OrderCnt			= oRs("OrderCnt")
		DelvNumber			= oRs("DelvNumber")
		ShopCD				= oRs("ShopCD")
		WareHouseType		= oRs("WareHouseType")

		ReturnName			= oRs("ReturnName")
		ReturnHp			= oRs("ReturnHp")
		ReturnZipCode		= oRs("ReturnZipCode")
		ReturnAddr1			= oRs("ReturnAddr1")
		ReturnAddr2			= oRs("ReturnAddr2")
ELSE
		oRs.Close
		Set oRs = Nothing
		oConn.Close
		Set oConn = Nothing

		resultMSG = "ROLLBACK"
		Response.Write resultMSG
		Response.End
END IF
oRs.Close


IF DB_PayType <> "C" AND DB_PayType <> "B" THEN
		Response.End
END IF




IF LGD_HASHDATA2 = LGD_HASHDATA THEN

		'//해쉬값 검증이 성공이면
		IF LGD_RESPCODE = "0000" THEN


				'# DB에 있는 금액과 PG사에서 넘어온 결재금액이 다르면 취소
				IF CDbl(LGD_AMOUNT) <> CDbl(DelvFee) THEN
						Set oRs = Nothing
						oConn.Close
						Set oConn = Nothing

						resultMSG = "ROLLBACK"
						Response.Write resultMSG
						Response.End
				END IF




				oConn.BeginTrans	

				
				'-----------------------------------------------------------------------------------------------------------'	
				'# 주문 교환/반품 신청 등록 Start
				'-----------------------------------------------------------------------------------------------------------'	
				' 1. 주문상품 상태변경
				' 2. 주문상품 변경이력 생성
				' 3. 주문상품 교환/반품 신청 이력 생성
				' 4. 교환/반품 신청 Temp에 OPCIdx 셋팅
				' 5. 업체별 교환/반품 배송비 생성
				'-----------------------------------------------------------------------------------------------------------'	
				SET oCmd = Server.CreateObject("ADODB.Command")
				WITH oCmd
						.ActiveConnection	 = oConn
						.CommandType		 = adCmdStoredProc
						.CommandText		 = "USP_Front_EShop_Order_Product_Cancel_Insert_From_Temp"

						.Parameters.Append .CreateParameter("@TempOPCIdx",			adInteger,	adParamInput,   ,	 Replace(LGD_OID, "OPC", ""))

						.Execute, , adExecuteNoRecords
				END WITH
				SET oCmd = Nothing

				IF Err.number <> 0 THEN
						oConn.RollbackTrans

						Set oRs = Nothing
						oConn.Close
						Set oConn = Nothing

						resultMSG = "ROLLBACK"
						Response.Write resultMSG
						Response.End
				END IF
				'-----------------------------------------------------------------------------------------------------------'	
				'# 주문 교환/반품 신청 등록 End
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
						.Parameters.Append .CreateParameter("@LGD_FINANCEAUTHNUM",			adVarChar,	adParamInput,	 20,	LGD_FINANCEAUTHNUM)
						.Parameters.Append .CreateParameter("@LGD_CARDNUM",					adVarChar,	adParamInput,	 30,	LGD_CARDNUM)
						.Parameters.Append .CreateParameter("@LGD_CARDINSTALLMONTH",		adVarChar,	adParamInput,	  2,	LGD_CARDINSTALLMONTH)
						.Parameters.Append .CreateParameter("@LGD_CARDNOINTYN",				adVarChar,	adParamInput,	  1,	LGD_CARDNOINTYN)
						.Parameters.Append .CreateParameter("@LGD_PCANCELFLAG",				adVarChar,	adParamInput,	  1,	LGD_PCANCELFLAG)
						.Parameters.Append .CreateParameter("@LGD_PCANCELSTR",				adVarChar,	adParamInput,	128,	LGD_PCANCELSTR)
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
						.Parameters.Append .CreateParameter("@LGD_TELNO",					adVarChar,	adParamInput,	 40,	LGD_TELNO)
						.Parameters.Append .CreateParameter("@CreateID",					adVarChar,	adParamInput,	 50,	U_NUM)
						.Parameters.Append .CreateParameter("@CreateIP",					adVarChar,	adParamInput,	 15,	U_IP)

						.Execute, , adExecuteNoRecords
				END WITH
				Set oCmd = Nothing

				IF Err.number <> 0 THEN
						oConn.RollbackTrans

						Set oRs = Nothing
						oConn.Close
						Set oConn = Nothing

						resultMSG = "ROLLBACK"
						Response.Write resultMSG
						Response.End
				END IF
				'-----------------------------------------------------------------------------------------------------------'	
				'결제 정보 저장 End
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

						Set oRs = Nothing
						oConn.Close
						Set oConn = Nothing

						resultMSG = "ROLLBACK"
						Response.Write resultMSG
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

						Set oRs = Nothing
						oConn.Close
						Set oConn = Nothing

						resultMSG = "ROLLBACK"
						Response.Write resultMSG
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

						Set oRs = Nothing
						oConn.Close
						Set oConn = Nothing

						resultMSG = "ROLLBACK"
						Response.Write resultMSG
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

				'-----------------------------------------------------------------------------------------------------------'	
				'메일발송 시작
				'-----------------------------------------------------------------------------------------------------------'	
				'Server.Execute("/Common/Mail/OrderMailSend.asp")
				'-----------------------------------------------------------------------------------------------------------'	
				'메일발송 끝
				'-----------------------------------------------------------------------------------------------------------'	
					

				Set oRs = Nothing
				oConn.Close
				Set oConn = Nothing

				resultMSG = "OK"	
				Response.Write resultMSG

				Response.End

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
						.Parameters.Append .CreateParameter("@LGD_FINANCEAUTHNUM",			adVarChar,	adParamInput,	 20,	LGD_FINANCEAUTHNUM)
						.Parameters.Append .CreateParameter("@LGD_CARDNUM",					adVarChar,	adParamInput,	 30,	LGD_CARDNUM)
						.Parameters.Append .CreateParameter("@LGD_CARDINSTALLMONTH",		adVarChar,	adParamInput,	  2,	LGD_CARDINSTALLMONTH)
						.Parameters.Append .CreateParameter("@LGD_CARDNOINTYN",				adVarChar,	adParamInput,	  1,	LGD_CARDNOINTYN)
						.Parameters.Append .CreateParameter("@LGD_PCANCELFLAG",				adVarChar,	adParamInput,	  1,	LGD_PCANCELFLAG)
						.Parameters.Append .CreateParameter("@LGD_PCANCELSTR",				adVarChar,	adParamInput,	128,	LGD_PCANCELSTR)
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
						.Parameters.Append .CreateParameter("@LGD_TELNO",					adVarChar,	adParamInput,	 40,	LGD_TELNO)
						.Parameters.Append .CreateParameter("@CreateID",					adVarChar,	adParamInput,	 50,	U_NUM)
						.Parameters.Append .CreateParameter("@CreateIP",					adVarChar,	adParamInput,	 15,	U_IP)

						.Execute, , adExecuteNoRecords
				END WITH
				Set oCmd = Nothing
				
				IF Err.number <> 0 THEN
						oConn.RollbackTrans

						Set oRs = Nothing
						oConn.Close
						Set oConn = Nothing

						resultMSG = "ROLLBACK"	
						Response.Write resultMSG
						Response.End
				END IF
				'-----------------------------------------------------------------------------------------------------------'
				'결제 정보 저장 End
				'-----------------------------------------------------------------------------------------------------------'

				oConn.CommitTrans

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
		
		Set oRs = Nothing
		oConn.Close
		Set oConn = Nothing

		resultMSG = "결제결과 상점 DB처리(NOTE_URL) 해쉬값 검증이 실패하였습니다."
		Response.Write resultMSG
		Response.End
END IF
%>

