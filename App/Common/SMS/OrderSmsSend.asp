<%@ Language=VBScript codepage="65001" %>
<%option Explicit%>
<%
'*****************************************************************************************'
'OrderSmsSend.asp - 주문관련 문자 보내기
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
<%
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 START
'-----------------------------------------------------------------------------------------------------------'
DIM oSConn
DIM oConn
DIM oRs
DIM oCmd

DIM sqlQuery

DIM N

DIM LGD_OID
DIM OrderCode
DIM CasFlag

DIM Idx			'//EShop_Order_Product의  Idx값
DIM CancelType
DIM PayType
DIM PayTypeName
DIM PayInfo
DIM OrderPrice
DIM UserID
DIM OrderName
DIM OrderHp
DIM OrderEmail
DIM ProductName
DIM SizeCD
DIM OrderCnt
DIM OrderDate
DIM ReceiveAddr
DIM LGD_PAYER
DIM LGD_FINANCENAME
DIM LGD_ACCOUNTNUM

DIM SendNum
DIM ReceiveNum
DIM SmsCode
DIM Msg
DIM SmsSubject
DIM SmsMessage
DIM OrderState
DIM CancelRequestFlag
DIM SMSYN
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


LGD_OID			= Trim(Request("LGD_OID"))
CasFlag			= Trim(Request("LGD_CASFLAG"))
OrderCode		= Trim(Request("OrderCode"))

IF OrderCode = "" THEN OrderCode = LGD_OID



SET oConn		= ConnectionOpen()	'//커넥션 생성
SET oRs			= Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성

SET oSConn		= SConnectionOpen()	'//커넥션 생성




'-----------------------------------------------------------------------------------------------------------'
'주문 정보 검색 START
'-----------------------------------------------------------------------------------------------------------'
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Admin_EShop_Order_Select_By_OrderCode"

		.Parameters.Append .CreateParameter("@OrderCode",			adVarChar,	adParamInput,	20,		OrderCode)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing

IF NOT oRs.EOF Then
		UserID			= oRs("UserID")
		OrderName		= oRs("OrderName")
		OrderHp			= oRs("OrderHp")
		OrderEmail		= oRs("OrderEmail")

		ReceiveNum		= OrderHp
END IF
oRs.Close


Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Front_EShop_Order_Select_For_OrderInfo"

		.Parameters.Append .CreateParameter("@OrderCode",	adVarChar, adParaminput,	20,		OrderCode)
		.Parameters.Append .CreateParameter("@UserID",		adVarChar, adParamInput,	20,		UserID)
		.Parameters.Append .CreateParameter("@OrderName",	adVarChar, adParamInput,	50,		OrderName)
		.Parameters.Append .CreateParameter("@OrderHp",		adVarChar, adParamInput,	20,		OrderHp)
		.Parameters.Append .CreateParameter("@OrderEmail",	adVarChar, adParamInput,	50,		OrderEmail)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing

IF NOT oRs.EOF THEN

		ProductName					 = oRs("ProductName")
		SizeCD						 = oRs("SizeCD")
		OrderCnt					 = oRs("OrderCnt")
		OrderPrice					 = oRs("OrderPrice")
		'SalePrice					 = oRs("SalePrice")
		'UseCouponPrice				 = oRs("UseCouponPrice")	
		'UsePointPrice				 = oRs("UsePointPrice")	
		'UseScashPrice				 = oRs("UseScashPrice")	
		'DeliveryPrice				 = oRs("DeliveryPrice")	
		PayType						 = oRs("PayType")
		PayTypeName					 = GetPayType(PayType)
		'DelvType					 = oRs("DelvType")
		'ShopNM						 = oRs("ShopNM")
		'ReceiveName					 = oRs("ReceiveName")
		'ReceiveTel					 = oRs("ReceiveTel")
		'ReceiveHp					 = oRs("ReceiveHp")
		'ReceiveZipCode				 = oRs("ReceiveZipCode")
		'ReceiveAddr1				 = oRs("ReceiveAddr1")
		'ReceiveAddr2				 = oRs("ReceiveAddr2")
		IF oRs("DelvType") = "S" THEN
				ReceiveAddr		= oRs("ShopNM")
		ELSE
				ReceiveAddr		= "(" & oRs("ReceiveZipCode") & ")" & oRs("ReceiveAddr1") & " " & oRs("ReceiveAddr2")
		END IF
		'ReceiptFlag					 = oRs("ReceiptFlag")
		'Memo						 = oRs("Memo")
		OrderDate					 = oRs("OrderDate")
		'OrderTime					 = oRs("OrderTime")
		LGD_PAYER					 = oRs("LGD_PAYER")
		LGD_FINANCENAME				 = oRs("LGD_FINANCENAME")
		'LGD_CARDINSTALLMONTH		 = oRs("LGD_CARDINSTALLMONTH")
		LGD_ACCOUNTNUM				 = oRs("LGD_ACCOUNTNUM")
		'LGD_TELNO					 = oRs("LGD_TELNO")
ELSE
		oRs.Close
		SET oRs = Nothing
		oConn.Close
		SET oConn = Nothing

		Call AlertMessage2("잘못된 주문 정보입니다", "location.href='/';")
		Response.End
END IF
oRs.Close
'-----------------------------------------------------------------------------------------------------------'
'주문정보 검색 END
'-----------------------------------------------------------------------------------------------------------'

IF CInt(OrderCnt) > 1 THEN
		ProductName		= ProductName & " 외 " & FormatNumber(CInt(OrderCnt) - 1, 0) & "건"
END IF



'-----------------------------------------------------------------------------------------------------------'
'관리자 정보 검색 START
'-----------------------------------------------------------------------------------------------------------'
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.Commandtype = adCmdStoredProc
		.CommandText = "USP_Admin_EShop_BizInfo_Select"
End WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing

IF NOT oRs.EOF THEN
		SendNum		= oRs("Tel")
End IF
oRs.Close
'-----------------------------------------------------------------------------------------------------------'
'관리자 정보 검색 END
'-----------------------------------------------------------------------------------------------------------'


'ON ERROR RESUME NEXT



'-----------------------------------------------------------------------------------------------------------'
'SMS 발송 START
'-----------------------------------------------------------------------------------------------------------'
IF PayType = "C" OR PayType = "B" OR PayType = "M" OR PayType = "N"  THEN		'//카드, 실시간 계좌체, 모바일결제, 네이버페이
		SmsCode		= "ORD_S300"

ELSEIF PayType = "V" THEN				'# 가상계좌
		IF CasFlag = "R" THEN				'# 계좌발급
				SmsCode		= "ORD_S100"

		ELSEIF CasFlag = "I" THEN			'# 입금완료
				SmsCode		= "ORD_S300"

		END IF
END IF


'# SMS 내용 검색
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Front_EShop_SmsMessage_Template_Select_By_Code"

		.Parameters.Append .CreateParameter("@Code",	adVarChar,		adParamInput,	20,	SmsCode)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing

IF NOT oRs.EOF THEN
		SmsSubject	= oRs("Subject")
		Msg			= oRs("SmsMsg")
END IF
oRs.Close
			

IF Msg <> "" THEN
		SmsMessage = Msg
		SmsMessage = Replace(SmsMessage, "#{이름}",			OrderName)
		SmsMessage = Replace(SmsMessage, "#{주문번호}",		OrderCode)
		SmsMessage = Replace(SmsMessage, "#{주문일자}",		GetDateYMD(OrderDate))
		SmsMessage = Replace(SmsMessage, "#{상품명}",		ProductName)
		SmsMessage = Replace(SmsMessage, "#{사이즈}",		SizeCD)
		SmsMessage = Replace(SmsMessage, "#{결제수단}",		PayTypeName)
		SmsMessage = Replace(SmsMessage, "#{주소}",			ReceiveAddr)
		SmsMessage = Replace(SmsMessage, "#{금액}",			FormatNumber(OrderPrice, 0))
		SmsMessage = Replace(SmsMessage, "#{입금계좌}",		"[" & LGD_FINANCENAME & "]" & LGD_ACCOUNTNUM)


		'# SMS 전송
		SmsSend oSConn, SmsSubject, SendNum, ReceiveNum, SmsMessage
END IF
'-----------------------------------------------------------------------------------------------------------'
'SMS 발송 END
'-----------------------------------------------------------------------------------------------------------'



Set oRs = Nothing
oConn.Close
Set oConn = Nothing
oSConn.Close
Set oSConn = Nothing
%>