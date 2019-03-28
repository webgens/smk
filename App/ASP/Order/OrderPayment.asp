<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'OrderPayment.asp - PG사 결제창 호출
'Date		: 2019.01.18
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

DIM OrderCode

DIM HTTP_USER_AGENT
DIM USER_AGENT
DIM wasUrl
DIM canUrl

DIM OrderName
DIM OrderTel
DIM OrderHp
DIM OrderEmail

DIM ReceiveName
DIM ReceiveTel
DIM ReceiveHp
DIM ReceiveZipCode
DIM ReceiveAddr1
DIM ReceiveAddr2

DIM GuaranteeInsurance
DIM GuaranteeInsuranceGubun
DIM USafeJumin1
DIM USafeJumin2

DIM PayType
DIM OrderPrice
DIM DeliveryPrice
DIM SettlePrice

DIM EscrowFlag

DIM LGD_PRODUCTINFO
DIM OrderCnt
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'



OrderCode			 = sqlFilter(Request("OrderCode"))

IF OrderCode = "" THEN
		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=주문 정보가 없습니다.<br />다시 시도하여 주십시오.&Script=APP_PopupHistoryBack();"
		Response.End
END IF




HTTP_USER_AGENT = LCase(Request.ServerVariables("HTTP_USER_AGENT"))
IF InStr(HTTP_USER_AGENT, "android") THEN
		USER_AGENT	 = "A"
		wasUrl		 = ""
		canUrl		 = ""
ELSEIF InStr(HTTP_USER_AGENT, "iphone") OR InStr(HTTP_USER_AGENT, "ipad") OR InStr(HTTP_USER_AGENT, "ipod") THEN
		USER_AGENT	 = "A"
		wasUrl		 = "krcoshoemarkerapp://applink?cont="
		canUrl		 = "krcoshoemarkerapp://applink?cont="
END IF






SET oConn			 = ConnectionOpen()							'# 커넥션 생성
SET oRs				 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성



SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Order_Product_Select_For_PG"

		.Parameters.Append .CreateParameter("@OrderCode",	adVarChar, adParaminput, 20, OrderCode)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing




IF NOT oRs.EOF THEN
		OrderName					 = oRs("OrderName")
		OrderTel					 = oRs("OrderTel")
		OrderHp						 = oRs("OrderHp")
		OrderEmail					 = oRs("OrderEmail")
	
		ReceiveName					 = oRs("ReceiveName")
		ReceiveTel					 = oRs("ReceiveTel")
		ReceiveHp					 = oRs("ReceiveHp")
		ReceiveZipCode				 = oRs("ReceiveZipCode")
		ReceiveAddr1				 = oRs("ReceiveAddr1")
		ReceiveAddr2				 = oRs("ReceiveAddr2")

		GuaranteeInsurance			 = oRs("GuaranteeInsurance")
		GuaranteeInsuranceGubun		 = oRs("GuaranteeInsuranceGubun")
		USafeJumin1					 = oRs("USafeJumin1")
		USafeJumin2					 = oRs("USafeJumin2")

		PayType						 = oRs("PayType")
	
		OrderPrice					 = oRs("OrderPrice")
		DeliveryPrice				 = oRs("DeliveryPrice")
		SettlePrice					 = CDbl(OrderPrice) + CDbl(SettlePrice)

		EscrowFlag					 = oRs("EscrowFlag")

		LGD_PRODUCTINFO				 = oRs("ProductName")
		OrderCnt					 = oRs("OrderCnt")
		IF OrderCnt > 1 THEN
		LGD_PRODUCTINFO				 = LGD_PRODUCTINFO & " 외 " & OrderCnt - 1 & "건"
		END IF
		
ELSE
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing

		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=주문 정보가 없습니다.<br />다시 시도하여 주십시오.&Script=APP_PopupHistoryBack();"
		Response.End
END IF
oRs.Close






'# 네이버페이
IF PayType = "N" THEN

		DIM NPay_ProductName
		DIM NPay_ProductItems


		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Order_Product_Select_For_PG_NPay"

				.Parameters.Append .CreateParameter("@OrderCode", adVarChar, adParamInput, 20, OrderCode)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		SET oCmd = Nothing

		IF NOT oRs.EOF THEN
				Do Until oRs.EOF

						'# 네이버페이 결제요청시 전송할 상품리스트
						IF NPay_ProductItems = "" THEN
								NPay_ProductName	= oRs("ProductName")
								NPay_ProductItems	= "{""categoryType"": ""PRODUCT"", ""categoryId"": ""GENERAL"", ""uid"": """ & oRs("ProductCD") & """, ""name"": """ & oRs("ProductName") & """, ""payReferrer"": ""PARTNER_DIRECT"", ""count"": " & oRs("OrderCnt") & "}"
						ELSE
								NPay_ProductItems	= NPay_ProductItems & ", " & "{""categoryType"": ""PRODUCT"", ""categoryId"": ""GENERAL"", ""uid"": """ & oRs("ProductCD") & """, ""name"": """ & oRs("ProductName") & """, ""payReferrer"": ""PARTNER_DIRECT"", ""count"": " & oRs("OrderCnt") & "}"
						END IF

						oRs.MoveNext
				Loop
		END IF
		oRs.Close

		'-----------------------------------------------------------------------------------------------------------'
		'# 네이버페이 결제 Start
		'-----------------------------------------------------------------------------------------------------------'
		NPay_ProductItems = "[" & NPay_ProductItems & "]"
%>
		<html>
		<head>
		<meta http-equiv="Content-Type" content="text/html; charset=EUC-KR">
		<title>네이버페이 결제</title>
		<script src="//nsp.pay.naver.com/sdk/js/naverpay.min.js"></script>
		<script type="text/javascript">
			function openNaverPay(orderCode, productName, productCount, totalPayAmount, productItems) {
				var oPay = Naver.Pay.create({
					"mode": "<%=NAVER_PAY_PLATFORM%>", // development or production
					"clientId": "<%=NAVER_PAY_CLIENTID%>", // clientId
					"openType": "page"
				});

				oPay.open({
					//"merchantUserKey": "123456",				// 가맹점 사용자 식별키
					"merchantPayKey": orderCode,			// 가맹점 주문 번호
					"productName": productName,			// 상품명을 입력하세요
					"productCount": productCount,			// 상품 수량
					"totalPayAmount": totalPayAmount,		// 결제금액
					"taxScopeAmount": totalPayAmount,		// 과세대상금액
					"taxExScopeAmount": "0",					// 면세대상금액
					"returnUrl": "<%=NAVER_PAY_RETURNURL%>?OrderCode=" + orderCode,		// 사용자 결제 완료 후 결제 결과를 받을 URL
					"productItems": productItems
				});
			}

			window.onload = function () {
				openNaverPay("<%=OrderCode%>", "<%=NPay_ProductName%>", "<%=OrderCnt%>", "<%=SettlePrice%>", <%=NPay_ProductItems%>);
			}
		</script>
		</head>

		<body oncontextmenu="return false" onselectstart="return false" ondragstart="return false">
			<table cellpadding="0" cellspacing="0" style="width:100%; height:100%;">
				<tr>
					<td align="center" valign="middle"><img src="<%=HOME_URL%>/Images/loading.gif" width="100" alt="LOADING" /></td>
				</tr>
			</table>
		</body>
		</html>
<%
		'-----------------------------------------------------------------------------------------------------------'
		'# 네이버페이 결제 End
		'-----------------------------------------------------------------------------------------------------------'
ELSE
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
		IF PAY_PLATFORM = "test" THEN								'# 상점아이디(자동생성)
				LGD_MID = CST_MID_TEST                                   
		ELSE
				LGD_MID = CST_MID                                         
		END IF

		DIM LGD_BUYERADDRESS
		LGD_BUYERADDRESS = ReceiveZipCode & " " & ReceiveAddr1 & " " & ReceiveAddr2


		DIM LGD_OID
		DIM LGD_AMOUNT
		DIM LGD_BUYER
		DIM LGD_BUYEREMAIL
		DIM LGD_TIMESTAMP
		DIM LGD_CUSTOM_SKIN
		LGD_OID						 = TRIM(OrderCode)					        '주문번호(상점정의 유니크한 주문번호를 입력하세요)
		LGD_AMOUNT					 = SettlePrice								'결제금액("," 를 제외한 결제금액을 입력하세요)
		'LGD_MERTKEY				 = LGD_MERTKEY								'[반드시 세팅]상점MertKey(mertkey는 상점관리자 -> 계약정보 -> 상점정보관리에서 확인하실수 있습니다')
		'LGD_PRODUCTINFO			 = LGD_PRODUCTINFO							'상품명
		LGD_BUYER					 = TRIM(OrderName)							'구매자명
		LGD_BUYEREMAIL				 = TRIM(OrderEmail)							'구매자 이메일
		LGD_TIMESTAMP				 = Year(Now) & Right("0" & Month(Now),2) & Right("0" & Day(Now),2) & Right("0" & Hour(Now),2) & Right("0" & Minute(Now),2) & Right("0" & Second(Now),2)		'타임스탬프
		'LGD_CUSTOM_FIRSTPAY	     = LGD_CUSTOM_FIRSTPAY						'상점정의 초기결제수단
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
		LGD_CASNOTEURL				 = MALL_OPENXPAY_CASNOTEURL

		' * LGD_RETURNURL 을 설정하여 주시기 바랍니다. 반드시 현재 페이지와 동일한 프로트콜 및  호스트이어야 합니다. 아래 부분을 반드시 수정하십시요.
		LGD_RETURNURL				 = MALL_OPENXPAY_RETURNURL
	

		' * ISP 카드결제 연동중 모바일ISP방식(고객세션을 유지하지않는 비동기방식)의 경우, LGD_KVPMISPNOTEURL/LGD_KVPMISPWAPURL/LGD_KVPMISPCANCELURL를 설정하여 주시기 바랍니다. 
		LGD_KVPMISPNOTEURL			 = MALL_KVPMISPNOREURL
		LGD_KVPMISPWAPURL			 = MALL_KVPMISPWAPURL & "?LGD_OID=" + LGD_OID    'ISP 카드 결제시, URL 대신 앱명 입력시, 앱호출함 
		LGD_KVPMISPCANCELURL		 = MALL_KVPMISPCANCELURL


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
				payReqMap.Add "LGD_KVPMISPWAPURL",						 wasUrl									'비동기 ISP(ex. 안드로이드) 승인완료후 사용자에게 보여지는 승인완료 URL
				payReqMap.Add "LGD_KVPMISPCANCELURL",					 canUrl									'ISP 앱에서 취소시 사용자에게 보여지는 취소 URL
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
				payReqMap.Add "LGD_MTRANSFERWAPURL",					 wasUrl
				payReqMap.Add "LGD_MTRANSFERCANCELURL",					 canUrl
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
		payReqMap.Add "LGD_RECEIVER",						 ReceiveName					'수취인
		payReqMap.Add "LGD_RECEIVERPHONE",					 ReceiveHp						'수취인 휴대번호
		payReqMap.Add "LGD_DELIVERYINFO",					 LGD_BUYERADDRESS				'배송지 주소

		'# 무통장(가상계좌) 입금 정보
		DIM CurDate 
		DIM LGD_CLOSEDATE
		CurDate												 = DATEADD("d", MALL_CLOSEDATE, Now)  
		LGD_CLOSEDATE										 = YEAR(CurDate) & RIGHT("0" & MONTH(CurDate), 2) & RIGHT("0" & DAY(CurDate), 2) & RIGHT("0" & HOUR(CurDate), 2) & RIGHT("0" & MINUTE(CurDate), 2) & RIGHT("0" & SECOND(CurDate), 2)
		payReqMap.Add "LGD_CLOSEDATE",						 LGD_CLOSEDATE					'결제가능일시(가상계좌 입금마감일시) yyyyMMddHHmmss 형식

		'# 에스크로 정보
		payReqMap.Add "LGD_ESCROW_USEYN",					 EscrowFlag						'에스크로 사용여부(매매보호)
		payReqMap.Add "LGD_ESCROW_GOODID",					 "1"							'에스크로 사용여부(매매보호)
		payReqMap.Add "LGD_ESCROW_GOODNAME",				 LGD_PRODUCTINFO				'에스크로 사용여부(매매보호)
		payReqMap.Add "LGD_ESCROW_GOODCODE",				 OrderCode						'에스크로 사용여부(매매보호)
		payReqMap.Add "LGD_ESCROW_UNITPRICE",				 SettlePrice					'에스크로 사용여부(매매보호)
		payReqMap.Add "LGD_ESCROW_QUANTITY",				 "1"							'에스크로 사용여부(매매보호)
		payReqMap.Add "LGD_ESCROW_ZIPCODE",					 ReceiveZipCode					'에스크로 사용여부(매매보호)
		payReqMap.Add "LGD_ESCROW_ADDRESS1",				 ReceiveAddr1					'에스크로 사용여부(매매보호)
		payReqMap.Add "LGD_ESCROW_ADDRESS2",				 ReceiveAddr2					'에스크로 사용여부(매매보호)
		payReqMap.Add "LGD_ESCROW_BUYERPHONE",				 ReceiveHp						'에스크로 사용여부(매매보호)

		'# 보증보험 정보
		IF GuaranteeInsurance = "Y" THEN
		payReqMap.Add "USAFE_GuaranteeInsurance",			 "Y"							'보증보험 발급 여부
		payReqMap.Add "USAFE_GuaranteeInsuranceAgreement",	 "Y"							'개인정보 동의 여부
		payReqMap.Add "USAFE_JuminNumber",					 USafeJumin1 & USafeJumin2		'개인정보 동의 여부
		payReqMap.Add "USAFE_EmailFlag",					 "Y"							'Email 동의 여부
		payReqMap.Add "USAFE_SmsFlag",						 "Y"							'Sms 동의 여부
		ELSE
		payReqMap.Add "USAFE_GuaranteeInsurance",			 "N"							'보증보험 발급 여부
		payReqMap.Add "USAFE_GuaranteeInsuranceAgreement",	 "N"							'개인정보 동의 여부
		payReqMap.Add "USAFE_JuminNumber",					 ""								'개인정보 동의 여부
		payReqMap.Add "USAFE_EmailFlag",					 "N"							'Email 동의 여부
		payReqMap.Add "USAFE_SmsFlag",						 "N"							'Sms 동의 여부
		END IF

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
					document.getElementById("LGD_PAYINFO").action = "PayRes.asp";
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
END IF


SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>