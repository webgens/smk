<%@ Language=VBScript codepage="65001" %>
<%
'*****************************************************************************************'
'ReturnUrl.asp - 결제 리턴 페이지
'Date		: 2018.11.30
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
<!-- #include virtual = "/API/json_for_asp/aspJSON1.17.asp" -->

<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<meta http-equiv="X-UA-Compatible" content="IE=edge">
	<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0, user-scalable=no">
	<title></title>
</head>
<body> <!-- oncontextmenu="return false" onselectstart="return false" ondragstart="return false">-->
	<table cellpadding="0" cellspacing="0" style="width:100%; height:100%;">
		<tr>
			<td align="center" valign="middle"><img src="<%=HOME_URL%>/Images/loading.gif" width="100" alt="LOADING" /></td>
		</tr>
	</table>
</body>
<%




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

DIM OrderCode
DIM PaymentID
DIM ResultCode
DIM ResultMsg

DIM OrderPrice
DIM OrderDate
DIM OrderTime
DIM OrderState
DIM SettleFlag
DIM SettleDate
DIM SettleTime
DIM PayType			: PayType	= "N"
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

OrderCode						 = Trim(sqlFilter(Request("OrderCode")))
PaymentID						 = Trim(sqlFilter(Request("PaymentID")))
ResultCode						 = Trim(sqlFilter(Request("ResultCode")))
ResultMsg						 = Trim(sqlFilter(Request("ResultMsg")))

'Response.Write "OrderCode = " & OrderCode & "<br>"
'Response.Write "PaymentID = " & PaymentID & "<br>"


IF ResultCode <> "Success" THEN
		IF ResultMsg = "userCancel" THEN
				ResultMsg	= "결제진행을 취소하셨습니다."
		ELSEIF ResultMsg = "paymentTimeExpire" THEN
				ResultMsg	= "결제시간을 초과 하셨습니다."
		ELSEIF ResultMsg = "OwnerAuthFail" THEN
				ResultMsg	= "본인 카드 인증 오류입니다."
		END IF

		Call AlertMessage2("주문 처리 도중 오류가 발생하였습니다. [" & ResultMsg & "] 다시 시도하여 주십시오.", "location.replace('/');")
		Response.End
END IF


IF OrderCode = "" OR PaymentID = "" THEN
		Call AlertMessage2("주문 처리 도중 오류가 발생하였습니다. 주문정보 / 결제정보가 없습니다. 다시 시도하여 주십시오.", "location.replace('/');")
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
				.Parameters.Append .CreateParameter("@Location",	 adChar,	 adParamInput,    1,	 "A")
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



DIM DBPayType
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Admin_EShop_Order_Select_By_OrderCode"
		.Parameters.Append .CreateParameter("@OrderCode",		adVarChar,	adParamInput,	20,		OrderCode)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing
	
IF NOT oRs.EOF THEN
		U_NUM			= oRs("UserID")
		U_NAME			= oRs("OrderName")
		DBPayType		= oRs("PayType")
		OrderPrice		= oRs("OrderPrice") + oRs("DeliveryPrice")
		OrderDate		= oRs("OrderDate")
		OrderTime		= oRs("OrderTime")
ELSE
		oRs.Close
		Set oRs = Nothing
		oConn.Close
		Set oConn = Nothing

		Call SettleErrorLogWrite(OrderCode, "Y", "NP01", "NPay_ReturnUrl", "EShop_Order Select 오류 / " & GetPayType(PayType) & " 취소 완료", Err.Description)
		Call AlertMessage2("주문 처리 도중 오류가 발생하였습니다. 결제 자동취소가 정상적으로 완료 되었습니다. 다시 시도하여 주십시오.", "location.replace('/');")
		Response.End
END IF
oRs.Close


'-----------------------------------------------------------------------------------------------------------'
'DB에 있는 결제수단이 네이버페이 인지 체크 START
'-----------------------------------------------------------------------------------------------------------'	
IF Trim(DBPayType) <> PayType THEN
		Set oRs = Nothing
		oConn.Close
		Set oConn = Nothing

		Call SettleErrorLogWrite(OrderCode, "Y", "NP02", "NPay_ReturnUrl", "결제수단상이 결:" & GetPayType(PayType) & " / 주:" & GetPayType(DBPayType) & " 취소 완료", Err.Description)
		Call AlertMessage2("주문 처리 도중 오류가 발생하였습니다. 결제 자동취소가 정상적으로 완료 되었습니다. 다시 시도하여 주십시오.", "location.replace('/');")
END IF
'-----------------------------------------------------------------------------------------------------------'
'DB에 있는 결제수단이 네이버페이 인지 체크 END
'-----------------------------------------------------------------------------------------------------------'	



OrderState		= "3"
SettleFlag		= "Y"
SettleDate		= U_DATE
SettleTime		= U_TIME



oConn.BeginTrans	



'-----------------------------------------------------------------------------------------------------------'	
'주문결제완료  업데이트 START
'1. 주문 정보 테이블에 결제 정보 Update
'2. 주문 상품 정보 테이블에 주문 상태 정보 Update
'3. 가상계좌 발급 또는 결제완료일 경우 처리
'	3-1. 쿠폰 사용 처리 Upudate
'	3-2. 포인트, 슈즈상품권 사용 처리
'		3-2-1. 포인트 사용 처리
'			3-2-1-1. 포인트 사용 등록 처리
'			3-2-1-2. 포인트 사용이력 등록 처리 시작
'				3-2-1-2-1. 회원포인트 사용이력 등록
'				3-2-1-2-2. 회원포인트 사용처리
'			3-2-1-3. 회원정보 포인트 누적처리
'		3-2-2. 슈즈상품권 사용 처리 시작
'			3-2-2-1. 슈즈상품권 사용 등록 처리
'			3-2-2-2. 슈즈상품권 사용이력 등록 처리
'				3-2-2-2-1. 회원슈즈상품권 사용이력 등록
'				3-2-2-2-2. 회원슈즈상품권 사용처리
'			3-2-2-3. 회원정보 슈즈상품권 누적처리
'	3-3. 임직원쿠폰 사용 처리 Upudate
'	3-4. 주문 상품 재고 Upudate
'	3-5. 장바구니 비우기
'	3-6. 주문서테이블 비우기
'-----------------------------------------------------------------------------------------------------------'	
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Front_EShop_Order_Update_For_SettleState"
		.Parameters.Append .CreateParameter("@OrderCode",			adVarChar,	adParamInput,	 20,		OrderCode)
		.Parameters.Append .CreateParameter("@OrderState",			adChar,		adParamInput,	  1,		OrderState)
		.Parameters.Append .CreateParameter("@SettleFlag",			adChar,		adParamInput,	  1,		SettleFlag)
		.Parameters.Append .CreateParameter("@SettleDate",			adChar,		adParamInput,	  8,		SettleDate)
		.Parameters.Append .CreateParameter("@SettleTime",			adChar,		adParamInput,	  6,		SettleTime)
		.Parameters.Append .CreateParameter("@ReceiptFlag",			adChar,		adParamInput,	  1,		"N")
		.Parameters.Append .CreateParameter("@ReceiptKind",			adChar,		adParamInput,	  1,		"")
		.Parameters.Append .CreateParameter("@EscrowFlag",			adChar,		adParamInput,	  1,		"N")
		.Parameters.Append .CreateParameter("@CasFlag",				adChar,		adParamInput,	  1,		"")
		.Parameters.Append .CreateParameter("@PayType",				adChar,		adParamInput,	  1,		PayType)
		.Parameters.Append .CreateParameter("@UpdateNM",			adVarChar,	adParamInput,	100,		U_NAME)
		.Parameters.Append .CreateParameter("@UpdateID",			adVarChar,	adParamInput,	 20,		U_NUM)
		.Parameters.Append .CreateParameter("@UpdateIP",			adVarChar,	adParamInput,	 15,		U_IP)
			
		.Execute, , adExecuteNoRecords
END WITH
Set oCmd = Nothing

IF Err.number <> 0 THEN
		oConn.RollbackTrans

		Set oRs = Nothing
		oConn.Close
		Set oConn = Nothing

		Call SettleErrorLogWrite(OrderCode, "Y", "NP11", "NPay_ReturnUrl", "EShop_Order 결제정보 업데이트 오류 / " & GetPayType(PayType) & " 취소 완료", Err.Description)
		Call AlertMessage2("주문 처리 도중 오류가 발생하였습니다. 결제 자동취소가 정상적으로 완료 되었습니다. 다시 시도하여 주십시오.", "location.replace('/');")
		Response.End
END IF
'-----------------------------------------------------------------------------------------------------------'	
'EShop_Order  업데이트 End
'-----------------------------------------------------------------------------------------------------------'	


'-----------------------------------------------------------------------------------------------------------'	
'네이버페이 결제승인 요청 Start
'-----------------------------------------------------------------------------------------------------------'	
DIM ResponseText
DIM HTTP_Object
	
Set HTTP_Object = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
With HTTP_Object
		'API 통신 Timeout 을 30초로 지정
		.SetTimeouts		30000, 30000, 30000, 30000
		.Open				"POST",						NAVER_PAY_PAYMENTURL, False
		.SetRequestHeader	"Content-Type",				"application/x-www-form-urlencoded"
		.SetRequestHeader	"X-Naver-Client-Id",		NAVER_PAY_CLIENTID
		.SetRequestHeader	"X-Naver-Client-Secret",	NAVER_PAY_CLIENTSECRET
		.Send				"paymentId=" & PaymentID
		.WaitForResponse

		IF .Status = 200 THEN
				ResponseText = .ResponseText
		ELSE
				ResponseText = ""
		END IF
End With


'# Response.Write "ResponseText = " & ResponseText & "<br><br>"

DIM Read_Data

DIM PayHistID
DIM MerchantID
DIM MerchantName
DIM MerchantPayKey
DIM MerchantUserKey
DIM AdmissionTypeCode
DIM AdmissionYmDT
DIM TradeConfirmYmDT
DIM AdmissionState
DIM TotalPayAmount
DIM PrimaryPayAmount
DIM NPointPayAmount
DIM PrimaryPayMeans
DIM CardCorpCode
DIM CardNo
DIM CardAuthNo
DIM CardInstCount
DIM BankCorpCode
DIM BankAccountNo
DIM ProductName
DIM SettleExpected
DIM SettleExpectAmount
DIM PayCommissionAmount
DIM ExtraDeduction
DIM UseCfmYmDT

IF ResponseText <> "" THEN
		Set Read_Data = New aspJSON
		Read_Data.loadJSON(ResponseText)
		With Read_Data
				ResultCode		= .data("code")
				ResultMsg		= .data("message")
				IF ResultCode = "Success" THEN
						PayHistID				= .data("body").item("detail").item("payHistId")
						MerchantID				= .data("body").item("detail").item("merchantId")
						MerchantName			= .data("body").item("detail").item("merchantName")
						MerchantPayKey			= .data("body").item("detail").item("merchantPayKey")
						MerchantUserKey			= .data("body").item("detail").item("merchantUserKey")
						AdmissionTypeCode		= .data("body").item("detail").item("admissionTypeCode")
						AdmissionYmDT			= .data("body").item("detail").item("admissionYmdt")
						TradeConfirmYmDT		= .data("body").item("detail").item("tradeConfirmYmdt")
						AdmissionState			= .data("body").item("detail").item("admissionState")
						TotalPayAmount			= .data("body").item("detail").item("totalPayAmount")
						PrimaryPayAmount		= .data("body").item("detail").item("primaryPayAmount")
						NPointPayAmount			= .data("body").item("detail").item("npointPayAmount")
						PrimaryPayMeans			= .data("body").item("detail").item("primaryPayMeans")
						CardCorpCode			= .data("body").item("detail").item("cardCorpCode")
						CardNo					= .data("body").item("detail").item("cardNo")
						CardAuthNo				= .data("body").item("detail").item("cardAuthNo")
						CardInstCount			= .data("body").item("detail").item("cardInstCount")
						BankCorpCode			= .data("body").item("detail").item("bankCorpCode")
						BankAccountNo			= .data("body").item("detail").item("bankAccountNo")
						ProductName				= .data("body").item("detail").item("productName")
						SettleExpected			= .data("body").item("detail").item("settleExpected")
						SettleExpectAmount		= .data("body").item("detail").item("settleExpectAmount")
						PayCommissionAmount		= .data("body").item("detail").item("payCommissionAmount")
						ExtraDeduction			= .data("body").item("detail").item("extraDeduction")
						UseCfmYmDT				= .data("body").item("detail").item("useCfmYmdt")

						IF SettleExpected				THEN SettleExpected			= "Y" ELSE SettleExpected = "N"
						IF ExtraDeduction				THEN ExtraDeduction			= "Y" ELSE ExtraDeduction = "N"
						IF TotalPayAmount		= ""	THEN TotalPayAmount			= "0"
						IF PrimaryPayAmount		= ""	THEN PrimaryPayAmount		= "0"
						IF NPointPayAmount		= ""	THEN NPointPayAmount		= "0"
						IF CardInstCount		= ""	THEN CardInstCount			= "0"
						IF SettleExpectAmount	= ""	THEN SettleExpectAmount		= "0"
						IF PayCommissionAmount	= ""	THEN PayCommissionAmount	= "0"

				ELSE
						oConn.RollbackTrans

						Set oRs = Nothing
						oConn.Close
						Set oConn = Nothing

						Call SettleErrorLogWrite(OrderCode, "Y", "NP12", "NPay_ReturnUrl", "네이버페이 결제승인요청 오류 / " & GetPayType(PayType) & " 취소 완료", ResultMsg)
						Call AlertMessage2("주문 처리 도중 오류가 발생하였습니다. 결제 자동취소가 정상적으로 완료 되었습니다. 다시 시도하여 주십시오.", "location.replace('/');")
						Response.End
				END IF
		End With
End If
'-----------------------------------------------------------------------------------------------------------'	
'네이버페이 결제승인 요청 Start
'-----------------------------------------------------------------------------------------------------------'	

'# Response.Write "PayHistID				= " & PayHistID				& "<br>"
'# Response.Write "MerchantID				= " & MerchantID			& "<br>"
'# Response.Write "MerchantName			= " & MerchantName			& "<br>"
'# Response.Write "MerchantPayKey			= " & MerchantPayKey		& "<br>"
'# Response.Write "MerchantUserKey			= " & MerchantUserKey		& "<br>"
'# Response.Write "AdmissionTypeCode		= " & AdmissionTypeCode		& "<br>"
'# Response.Write "AdmissionYmDT			= " & AdmissionYmDT			& "<br>"
'# Response.Write "TradeConfirmYmDT		= " & TradeConfirmYmDT		& "<br>"
'# Response.Write "AdmissionState			= " & AdmissionState		& "<br>"
'# Response.Write "TotalPayAmount			= " & TotalPayAmount		& "<br>"
'# Response.Write "PrimaryPayAmount		= " & PrimaryPayAmount		& "<br>"
'# Response.Write "NPointPayAmount			= " & NPointPayAmount		& "<br>"
'# Response.Write "PrimaryPayMeans			= " & PrimaryPayMeans		& "<br>"
'# Response.Write "CardCorpCode			= " & CardCorpCode			& "<br>"
'# Response.Write "CardNo					= " & CardNo				& "<br>"
'# Response.Write "CardAuthNo				= " & CardAuthNo			& "<br>"
'# Response.Write "CardInstCount			= " & CardInstCount			& "<br>"
'# Response.Write "BankCorpCode			= " & BankCorpCode			& "<br>"
'# Response.Write "BankAccountNo			= " & BankAccountNo			& "<br>"
'# Response.Write "ProductName				= " & ProductName			& "<br>"
'# Response.Write "SettleExpected			= " & SettleExpected		& "<br>"
'# Response.Write "SettleExpectAmount		= " & SettleExpectAmount	& "<br>"
'# Response.Write "PayCommissionAmount		= " & PayCommissionAmount	& "<br>"
'# Response.Write "ExtraDeduction			= " & ExtraDeduction		& "<br>"
'# Response.Write "UseCfmYmDT				= " & UseCfmYmDT			& "<br>"

'# PayHistID				= 20190118NP1000534570<br>
'# MerchantID				= shoemarker01<br>
'# MerchantName			= 슈마커 공식쇼핑몰<br>
'# MerchantPayKey			= C0001000152<br>
'# MerchantUserKey			= <br>
'# AdmissionTypeCode		= 01<br>
'# AdmissionYmDT			= 20190118124250<br>
'# TradeConfirmYmDT		= <br>
'# AdmissionState			= SUCCESS<br>
'# TotalPayAmount			= 79000<br>
'# PrimaryPayAmount		= 79000<br>
'# NPointPayAmount			= 0<br>
'# PrimaryPayMeans			= CARD<br>
'# CardCorpCode			= C0<br>
'# CardNo					= 0000-***********<br>
'# CardAuthNo				= 00000000<br>
'# CardInstCount			= 0<br>
'# BankCorpCode			= <br>
'# BankAccountNo			= <br>
'# ProductName				= 락카디아 트레일<br>
'# SettleExpected			= N<br>
'# SettleExpectAmount		= 0<br>
'# PayCommissionAmount		= 0<br>
'# ExtraDeduction			= N<br>
'# UseCfmYmDT				= <br> <font face="Arial" size=2>

'-----------------------------------------------------------------------------------------------------------'	
'결제 정보 저장 START
'-----------------------------------------------------------------------------------------------------------'
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Front_EShop_Order_Settle_Npay_Insert"

		.Parameters.Append .CreateParameter("@OrderCode",					adVarChar,	adParamInput,	  20,	OrderCode)
		.Parameters.Append .CreateParameter("@ResultCode",					adVarChar,	adParamInput,	  20,	ResultCode)
		.Parameters.Append .CreateParameter("@ResultMsg",					adVarChar,	adParamInput,	1000,	ResultMsg)
		.Parameters.Append .CreateParameter("@PaymentID",					adVarChar,	adParamInput,	  50,	PaymentID)
		.Parameters.Append .CreateParameter("@PayHistID",					adVarChar,	adParamInput,	  50,	PayHistID)
		.Parameters.Append .CreateParameter("@MerchantID",					adVarChar,	adParamInput,	  50,	MerchantID)
		.Parameters.Append .CreateParameter("@MerchantName",				adVarChar,	adParamInput,	  50,	MerchantName)
		.Parameters.Append .CreateParameter("@MerchantPayKey",				adVarChar,	adParamInput,	  64,	MerchantPayKey)
		.Parameters.Append .CreateParameter("@MerchantUserKey",				adVarChar,	adParamInput,	  50,	MerchantUserKey)
		.Parameters.Append .CreateParameter("@AdmissionTypeCode",			adVarChar,	adParamInput,	   2,	AdmissionTypeCode)
		.Parameters.Append .CreateParameter("@AdmissionYmDT",				adVarChar,	adParamInput,	  14,	AdmissionYmDT)
		.Parameters.Append .CreateParameter("@TradeConfirmYmDT",			adVarChar,	adParamInput,	  50,	TradeConfirmYmDT)
		.Parameters.Append .CreateParameter("@AdmissionState",				adVarChar,	adParamInput,	  10,	AdmissionState)
		.Parameters.Append .CreateParameter("@TotalPayAmount",				adCurrency,	adParamInput,	    ,	TotalPayAmount)
		.Parameters.Append .CreateParameter("@PrimaryPayAmount",			adCurrency,	adParamInput,	    ,	PrimaryPayAmount)
		.Parameters.Append .CreateParameter("@NPointPayAmount",				adCurrency,	adParamInput,	    ,	NPointPayAmount)
		.Parameters.Append .CreateParameter("@PrimaryPayMeans",				adVarChar,	adParamInput,	  10,	PrimaryPayMeans)
		.Parameters.Append .CreateParameter("@CardCorpCode",				adVarChar,	adParamInput,	  10,	CardCorpCode)
		.Parameters.Append .CreateParameter("@CardNo",						adVarChar,	adParamInput,	  50,	CardNo)
		.Parameters.Append .CreateParameter("@CardAuthNo",					adVarChar,	adParamInput,	  30,	CardAuthNo)
		.Parameters.Append .CreateParameter("@CardInstCount",				adInteger,	adParamInput,	    ,	CardInstCount)
		.Parameters.Append .CreateParameter("@BankCorpCode",				adVarChar,	adParamInput,	  10,	BankCorpCode)
		.Parameters.Append .CreateParameter("@BankAccountNo",				adVarChar,	adParamInput,	  50,	BankAccountNo)
		.Parameters.Append .CreateParameter("@ProductName",					adVarChar,	adParamInput,	 128,	ProductName)
		.Parameters.Append .CreateParameter("@SettleExpected",				adChar,		adParamInput,	   1,	SettleExpected)
		.Parameters.Append .CreateParameter("@SettleExpectAmount",			adCurrency,	adParamInput,	    ,	SettleExpectAmount)
		.Parameters.Append .CreateParameter("@PayCommissionAmount",			adCurrency,	adParamInput,	    ,	PayCommissionAmount)
		.Parameters.Append .CreateParameter("@ExtraDeduction",				adChar,		adParamInput,	   1,	ExtraDeduction)
		.Parameters.Append .CreateParameter("@UseCfmYmDT",					adVarChar,	adParamInput,	   8,	UseCfmYmDT)
		.Parameters.Append .CreateParameter("@CreateID",					adVarChar,	adParamInput,	  50,	U_NUM)
		.Parameters.Append .CreateParameter("@CreateIP",					adVarChar,	adParamInput,	  15,	U_IP)

		.Execute, , adExecuteNoRecords
END WITH
Set oCmd = Nothing

IF Err.number <> 0 THEN
		oConn.RollbackTrans

		Set oRs = Nothing
		oConn.Close
		Set oConn = Nothing

		IF cancelNaverpay(PaymentID, TotalPayAmount, "DBSaveError", "2") THEN
				Call SettleErrorLogWrite(OrderCode, "Y", "NP13", "NPay_ReturnUrl", "EShop_Order_Settle_NPay 저장 오류 / " & GetPayType(PayType) & " 취소 완료", Err.Description)
				Call AlertMessage2("주문 처리 도중 오류가 발생하였습니다. 결제 자동취소가 정상적으로 완료 되었습니다. 다시 시도하여 주십시오.", "location.replace('/');")
				Response.End
		ELSE
				Call SettleErrorLogWrite(OrderCode, "N", "NP13", "NPay_ReturnUrl", "EShop_Order_Settle_NPay 저장 오류 / " & GetPayType(PayType) & " 취소 오류", Err.Description)
				Call AlertMessage2("주문 처리 도중 오류가 발생하였습니다. 결제 자동취소가 정상적으로 처리되지 않았습니다. 관리자에게 문의 바랍니다.", "location.replace('/');")
				Response.End
		END IF
END IF
'-----------------------------------------------------------------------------------------------------------'	
'결제 정보 저장 End
'-----------------------------------------------------------------------------------------------------------'	



'-----------------------------------------------------------------------------------------------------------'	
'ERP 전송용 I/F 주문 생성 START
'-----------------------------------------------------------------------------------------------------------'	
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Admin_EShop_Order_Product_Select_By_OrderCode"

		.Parameters.Append .CreateParameter("@OrderCode",		adVarChar,	adParamInput,	20,		OrderCode)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing

IF NOT oRs.EOF THEN
		Do Until oRs.EOF
				'# 예약상품이 아닌 경우만 ERP 전송
				IF oRs("OrderType") <> "R" THEN
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
								SET oRs = Nothing
								oConn.Close
								SET oConn = Nothing


								IF cancelNaverpay(PaymentID, TotalPayAmount, "DBSaveError", "2") THEN
										Call SettleErrorLogWrite(OrderCode, "Y", "NP14", "NPay_ReturnUrl", "IF_ONLINE_ORDER 입력 오류 / " & GetPayType(PayType) & " 취소 완료", Err.Description)
										Call AlertMessage2("주문 처리 도중 오류가 발생하였습니다. 결제 자동취소가 정상적으로 완료 되었습니다. 다시 시도하여 주십시오.", "location.replace('/');")
										Response.End
								ELSE
										Call SettleErrorLogWrite(OrderCode, "N", "NP14", "NPay_ReturnUrl", "IF_ONLINE_ORDER 입력 오류 / " & GetPayType(PayType) & " 취소 오류", Err.Description)
										Call AlertMessage2("주문 처리 도중 오류가 발생하였습니다. 결제 자동취소가 정상적으로 처리되지 않았습니다. 관리자에게 문의 바랍니다.", "location.replace('/');")
										Response.End
								END IF
						END IF
				END IF

				oRs.MoveNext
		Loop 
End IF
oRs.Close
'-----------------------------------------------------------------------------------------------------------'	
'ERP 전송용 I/F 주문 생성 End
'-----------------------------------------------------------------------------------------------------------'	


oConn.CommitTrans



'-----------------------------------------------------------------------------------------------------------'	
'문자발송 시작
'-----------------------------------------------------------------------------------------------------------'	
'# Server.Execute("/Common/SMS/OrderSmsSend.asp")
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_Order_Sms_Send"

		.Parameters.Append .CreateParameter("@OrderCode",	 adVarChar,	 adParamInput,   20,	 OrderCode)
		.Parameters.Append .CreateParameter("@OPIdx",		 adInteger,	 adParamInput,     ,	 0)
		.Parameters.Append .CreateParameter("@SmsCode",		 adVarChar,	 adParamInput,   20,	 "ORD_S300")

		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing
'-----------------------------------------------------------------------------------------------------------'	
'문자발송 끝
'-----------------------------------------------------------------------------------------------------------'	

'-----------------------------------------------------------------------------------------------------------'	
'메일발송 시작
'-----------------------------------------------------------------------------------------------------------'	
Server.Execute("/Common/Mail/OrderMailSend.asp")
'-----------------------------------------------------------------------------------------------------------'	
'메일발송 끝
'-----------------------------------------------------------------------------------------------------------'	


'-----------------------------------------------------------------------------------------------------------'	
'네이버페이 결제취소 함수 Start
'-----------------------------------------------------------------------------------------------------------'	
Function cancelNaverPay(ByVal fPaymentID, ByVal fCancelAmount, ByVal fCancelReason, ByVal fCancelRequester)
		DIM retVal			: retVal	= false
		DIM fHTTP_Object
		DIM fResponseText	: fResponseText	= ""
		DIM fRead_Data
		DIM fResultCode
		DIM fParam

		fParam	= ""
		fParam	= fParam & "paymentId="			& fPaymentID
		fParam	= fParam & "&cancelAmount="		& fCancelAmount
		fParam	= fParam & "&taxScopeAmount="	& fCancelAmount
		fParam	= fParam & "&taxExScopeAmount=" & "0"
		fParam	= fParam & "&cancelReason="		& fCancelReason
		fParam	= fParam & "&cancelRequester="	& fCancelRequester				'# 1:사용자, 2:관리자

		Set fHTTP_Object = Server.CreateObject("WinHttp.WinHttpRequest.5.1")
		With fHTTP_Object
				'결제취소는 API 통신 Timeout 을 60초로 지정
				.SetTimeouts 60000, 60000, 60000, 60000
				.Open				"POST",						NAVER_PAY_CANCELURL, False
				.SetRequestHeader	"Content-Type",				"application/x-www-form-urlencoded"
				.SetRequestHeader	"X-Naver-Client-Id",		NAVER_PAY_CLIENTID
				.SetRequestHeader	"X-Naver-Client-Secret",	NAVER_PAY_CLIENTSECRET
				.Send				fParam
				.WaitForResponse

				IF .Status = 200 THEN
						fResponseText = .ResponseText
				ELSE
						fResponseText = ""
				END IF
		End With

		IF fResponseText <> "" THEN
				Set fRead_Data = New aspJSON
				fRead_Data.loadJSON(fResponseText)
				With fRead_Data
						fResultCode		= .data("code")
						IF fResultCode = "Success" THEN
								retVal	= true
						END IF
				End With
		End If

		cancelNaverPay = retVal
End Function
'-----------------------------------------------------------------------------------------------------------'	
'네이버페이 결제취소 함수 End
'-----------------------------------------------------------------------------------------------------------'	

Set oRs = Nothing
oConn.Close
Set oConn = Nothing


Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=&Msg=&Script=APP_PopupHistoryBack_Move('/ASP/Order/OrderComplete.asp?OrderCode=" & OrderCode & "');"
Response.End
%>
<script type="text/javascript">
	//alert("결제가 정상적으로 처리되었습니다.");
	APP_PopupHistoryBack_Move("/ASP/Order/OrderComplete.asp?OrderCode=<%=OrderCode%>");
</script>
