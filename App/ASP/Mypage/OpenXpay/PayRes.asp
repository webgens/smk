<%@ Language=VBScript codepage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************/
'PayRes.asp - 카드결제(안심결제) 결과 처리 및 리턴 / 가상계좌 리턴 페이지
'Date		: 2019.01.02
'Update	: 
'/****************************************************************************************/

'# 페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'Response.Buffer = True
'-----------------------------------------------------------------------------------------'
%>

<!-- #include Virtual = "/ADO/ADODBCommon.asp" -->
<!-- #include Virtual = "/Common/Common.asp" -->
<!-- #include Virtual = "/Common/SetInfo.asp" -->

<html>
<head>
	<meta http-equiv="Content-Type" content="text/html; charset=utf-8">
	<meta http-equiv="X-UA-Compatible" content="IE=edge">
	<meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, minimum-scale=1.0, user-scalable=no">
	<title></title>
</head>
<body> <!-- oncontextmenu="return false" onselectstart="return false" ondragstart="return false">-->
	<table cellpadding="0" cellspacing="0" width="100%" height="100%">
		<tr>
			<td align="center" valign="middle"><img src="<%=HOME_URL%>/Images/loading.gif" width="100" alt="LOADING" /></td>
		</tr>
	</table>
</body>
</html>

<%
'Response.Flush


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




ON ERROR RESUME NEXT

'/*
' * [최종결제요청 페이지(STEP2-2)]
' *
' * LG유플러스으로 부터 내려받은 LGD_PAYKEY(인증Key)를 가지고 최종 결제요청.(파라미터 전달시 POST를 사용하세요)
' */
DIM configPath
configPath = "C:/LGDacom"  'LG유플러스에서 제공한 환경파일("/conf/lgdacom.conf, /conf/mall.conf") 위치 지정.  
'configPath = "C:/lgdacom"  'LG유플러스에서 제공한 환경파일("/conf/lgdacom.conf, /conf/mall.conf") 위치 지정.  


IF Err.number <> 0 THEN
		Call SettleErrorLogWrite(Trim(Request("LGD_OID")), "N", "PR01", "PayRes", "configPath 설정 오류", Err.Description)
		Call AlertMessage2("주문 처리 도중 오류가 발생하였습니다.[PR01]\n\n결제가 정상 처리 되지 않았습니다.\n\n관리자에게 문의 바랍니다.", "location.replace('/ASP/Mypage/OrderList.asp');")
		Response.End
END IF


'/*
' *************************************************
' * 1.최종결제 요청 - BEGIN
' *  (단, 최종 금액체크를 원하시는 경우 금액체크 부분 주석을 제거 하시면 됩니다.)
' *************************************************
' */
DIM CST_PLATFORM
DIM LGD_OID
DIM LGD_MID
DIM LGD_PAYKEY

CST_PLATFORM			= Trim(Request("CST_PLATFORM"))
'#LGD_OID				= Trim(Request("LGD_OID"))
LGD_MID					= Trim(Request("CST_MID"))
IF CST_PLATFORM = "test" THEN
		LGD_MID			= "t" & CST_MID
ELSE
		LGD_MID			= CST_MID
END IF
LGD_PAYKEY				= Trim(request("LGD_PAYKEY"))

DIM xpay				'결제요청 API 객체
DIM amount_check		'금액비교 결과
DIM i, j
DIM itemName

'해당 API를 사용하기 위해 setup.exe 를 설치해야 합니다.
Set xpay = server.CreateObject("XPayClientCOM.XPayClient")

IF Err.number <> 0 THEN
		Call SettleErrorLogWrite(LGD_OID, "N", "PR02", "PayRes", "Server.CreateObject(""XPayClientCOM.XPayClient"") 개체 생성 오류", Err.Description)
		Call AlertMessage2("주문 처리 도중 오류가 발생하였습니다.[PR02]\n\n결제가 정상 처리 되지 않았습니다.\n\n관리자에게 문의 바랍니다.", "location.replace('/ASP/Mypage/OrderList.asp');")
		Response.End
END IF

xpay.Init configPath, CST_PLATFORM

xpay.Init_TX(LGD_MID)
xpay.Set "LGD_TXNAME", "PaymentByKey"
xpay.Set "LGD_PAYKEY", LGD_PAYKEY


    
DIM oConn						'# ADODB Connection 개체
DIM oRs							'# ADODB Recordset 개체
DIM oCmd						'# ADODB Command 개체






'금액을 체크하시기 원하는 경우 아래 주석을 풀어서 이용하십시요.
'DB_AMOUNT = "DB나 세션에서 가져온 금액" 	'반드시 위변조가 불가능한 곳(DB나 세션)에서 금액을 가져오십시요.
'xpay.Set "LGD_AMOUNTCHECKYN", "Y"
'xpay.Set "LGD_AMOUNT", DB_AMOUNT




	    
'/*
' *************************************************
' * 1.최종결제 요청(수정하지 마세요) - END
' *************************************************
' */

'/*
' * 2. 최종결제 요청 결과처리
' *
' * 최종 결제요청 결과 리턴 파라미터는 연동메뉴얼을 참고하시기 바랍니다.
' */



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


DIM ReceiptFlag : ReceiptFlag = "N"

DIM isDBOK

DIM LGD_RESPCODE
DIM LGD_RESPMSG
DIM LGD_AMOUNT
DIM LGD_TID
DIM LGD_TIMESTAMP
DIM LGD_PAYTYPE
DIM LGD_PAYDATE
DIM LGD_HASHDATA
DIM LGD_FINANCECODE
DIM LGD_FINANCENAME
DIM LGD_FINANCEAUTHNUM
DIM LGD_CARDNUM
DIM LGD_CARDINSTALLMONTH
DIM LGD_CARDNOINTYN				'# 무이자할부여부(신용카드) - '1'이면 무이자할부 '0'이면 일반할부
DIM LGD_PCANCELFLAG				'# 0: 부분취소불가능,  1: 부분취소가능
DIM LGD_PCANCELSTR				'# 부분취소가능시는 "0" 으로 리턴
DIM LGD_ESCROWYN
DIM LGD_CASHRECEIPTNUM			'# 현금영수증 승인번호
DIM LGD_CASHRECEIPTSELFYN		'# 현금영수증자진발급제유무 Y: 자진발급제 적용, 그외 : 미적용
DIM LGD_CASHRECEIPTKIND			'# 현금영수증 종류 0: 소득공제용 , 1: 지출증빙용
DIM LGD_ACCOUNTNUM				'# 계좌번호(무통장입금)
DIM LGD_ACCOUNTOWNER			'# 계좌주명
DIM LGD_PAYER					'# 입금자명
DIM LGD_CASTAMOUNT				'# 입금총액(무통장입금)
DIM LGD_CASCAMOUNT				'# 현입금액(무통장입금)
DIM LGD_CASFLAG					'# 무통장입금 플래그(무통장입금) - 'R':계좌할당, 'I':입금, 'C':입금취소
DIM LGD_CASSEQNO				'# 입금순서(무통장입금)
DIM LGD_SAOWNER					'# 가상계좌 입금계좌주명.상점명이 디폴트로 리턴
DIM LGD_TELNO

DIM LGD_PRODUCTINFO
DIM LGD_BUYER
DIM LGD_BUYERID
DIM LGD_BUYERPHONE
DIM LGD_BUYERADDRESS
DIM LGD_BUYEREMAIL
DIM LGD_RECEIVER
DIM LGD_RECEIVERPHONE


'USafe 보증보험 관련
DIM USAFE_GuaranteeInsurance
DIM USAFE_GuaranteeInsuranceAgreement
DIM USAFE_JuminNumber
DIM USAFE_EmailFlag
DIM USAFE_SmsFlag


IF  xpay.TX() THEN
		'1)결제결과 화면처리(성공,실패 결과 처리를 하시기 바랍니다.)
		'# Response.Write("결제요청이 완료되었습니다. <br>")
		'# Response.Write("TX Response_code = " & xpay.resCode & "<br>")
		'# Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")

		'# Response.Write("거래번호 : " & xpay.Response("LGD_TID", 0) & "<br>")
		'# Response.Write("상점아이디 : " & xpay.Response("LGD_MID", 0) & "<br>")
		'# Response.Write("상점주문번호 : " & xpay.Response("LGD_OID", 0) & "<br>")
		'# Response.Write("결제금액 : " & xpay.Response("LGD_AMOUNT", 0) & "<br>")
		'# Response.Write("결과코드 : " & xpay.Response("LGD_RESPCODE", 0) & "<br>")
		'# Response.Write("결과메세지 : " & xpay.Response("LGD_RESPMSG", 0) & "<p>")

		'# Response.Write("[결제요청 결과 파라미터]<br>")

		LGD_RESPCODE				 = xpay.Response("LGD_RESPCODE", 0)
		LGD_RESPMSG					 = xpay.Response("LGD_RESPMSG", 0)
		LGD_AMOUNT					 = xpay.Response("LGD_AMOUNT", 0)
		LGD_TID						 = xpay.Response("LGD_TID", 0)
		LGD_OID						 = xpay.Response("LGD_OID", 0)
		LGD_TIMESTAMP				 = xpay.Response("LGD_TIMESTAMP", 0)
		LGD_PAYTYPE					 = xpay.Response("LGD_PAYTYPE", 0)
		LGD_PAYDATE					 = xpay.Response("LGD_PAYDATE", 0)
		LGD_HASHDATA				 = xpay.Response("LGD_HASHDATA", 0)
		LGD_FINANCECODE				 = xpay.Response("LGD_FINANCECODE", 0)
		LGD_FINANCENAME				 = xpay.Response("LGD_FINANCENAME", 0)
		LGD_FINANCEAUTHNUM			 = xpay.Response("LGD_FINANCEAUTHNUM", 0)
		LGD_CARDNUM					 = xpay.Response("LGD_CARDNUM", 0)
		LGD_CARDINSTALLMONTH		 = xpay.Response("LGD_CARDINSTALLMONTH", 0)
		LGD_CARDNOINTYN				 = xpay.Response("LGD_CARDNOINTYN", 0)
		LGD_PCANCELFLAG				 = xpay.Response("LGD_PCANCELFLAG", 0)
		LGD_PCANCELSTR				 = xpay.Response("LGD_PCANCELSTR", 0)
		LGD_ESCROWYN				 = xpay.Response("LGD_ESCROWYN", 0)
		LGD_CASHRECEIPTNUM			 = xpay.Response("LGD_CASHRECEIPTNUM", 0)
		LGD_CASHRECEIPTSELFYN 		 = xpay.Response("LGD_CASHRECEIPTSELFYN", 0)
		LGD_CASHRECEIPTKIND			 = xpay.Response("LGD_CASHRECEIPTKIND", 0)
		LGD_ACCOUNTNUM				 = xpay.Response("LGD_ACCOUNTNUM", 0)
		LGD_ACCOUNTOWNER			 = xpay.Response("LGD_ACCOUNTOWNER", 0)
		LGD_PAYER					 = xpay.Response("LGD_PAYER", 0)
		LGD_CASTAMOUNT				 = xpay.Response("LGD_CASTAMOUNT", 0)
		LGD_CASCAMOUNT				 = xpay.Response("LGD_CASCAMOUNT", 0)
		LGD_CASFLAG					 = xpay.Response("LGD_CASFLAG", 0)
		LGD_CASSEQNO				 = xpay.Response("LGD_CASSEQNO", 0)
		LGD_SAOWNER					 = xpay.Response("LGD_SAOWNER", 0)
		LGD_TELNO					 = xpay.Response("LGD_TELNO", 0)

		LGD_BUYER					 = xpay.Response("LGD_BUYER", 0)
		LGD_BUYERID					 = xpay.Response("LGD_BUYERID", 0)
		LGD_BUYERPHONE				 = xpay.Response("LGD_BUYERPHONE", 0)
		LGD_BUYERADDRESS			 = xpay.Response("LGD_BUYERADDRESS", 0)
		LGD_BUYEREMAIL				 = xpay.Response("LGD_BUYEREMAIL", 0)
		LGD_RECEIVER				 = xpay.Response("LGD_RECEIVER", 0)
		LGD_RECEIVERPHONE			 = xpay.Response("LGD_RECEIVERPHONE", 0)
		LGD_PRODUCTINFO				 = xpay.Response("LGD_PRODUCTINFO", 0)

		USAFE_GuaranteeInsurance			 = xpay.Response("USAFE_GuaranteeInsurance", 0)
		USAFE_GuaranteeInsuranceAgreement	 = xpay.Response("USAFE_GuaranteeInsuranceAgreement", 0)
		USAFE_JuminNumber					 = xpay.Response("USAFE_JuminNumber", 0)
		USAFE_EmailFlag						 = xpay.Response("USAFE_EmailFlag", 0)
		USAFE_SmsFlag						 = xpay.Response("USAFE_SmsFlag", 0)

		IF USAFE_GuaranteeInsurance			 = "" THEN USAFE_GuaranteeInsurance			 = "N"
		IF USAFE_GuaranteeInsuranceAgreement = "" THEN USAFE_GuaranteeInsuranceAgreement = "N"
		IF USAFE_EmailFlag					 = "" THEN USAFE_EmailFlag					 = "N"
		IF USAFE_SmsFlag					 = "" THEN USAFE_SmsFlag					 = "N"



		IF LGD_ESCROWYN = "" OR IsNull(LGD_ESCROWYN) THEN
				LGD_ESCROWYN = "N"
		END IF
		IF LGD_CASHRECEIPTNUM <> "" THEN
				ReceiptFlag = "Y"
		END IF

		SELECT CASE LGD_PAYTYPE
				CASE "SC0010" : PayType = "C"			'# 신용카드
				CASE "SC0030" : PayType = "B"			'# 계좌이체
				CASE "SC0040" : PayType = "V"			'# 가상계좌
				CASE "SC0060" : PayType = "M"			'# 휴대폰결제
		END SELECT

		'아래는 결제요청 결과 파라미터를 모두 찍어 줍니다.
		'# DIM itemCount
		'# DIM resCount
		'# itemCount	 = xpay.resNameCount
		'# resCount		 = xpay.resCount

		'# FOR i = 0 TO itemCount - 1
		'# 		itemName = xpay.ResponseName(i)
		'# 		Response.Write(itemName & "&nbsp: ")
		'# 		FOR j = 0 TO resCount - 1
		'# 				Response.Write(xpay.Response(itemName, j) & "<br>")
		'# 		NEXT
		'# NEXT
            
		'# Response.Write("<p>")
          
		SET oConn	= ConnectionOpen()	'//커넥션 생성
		SET oRs		= Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성

		IF Err.number <> 0 THEN
				isDBOK = false

				xpay.Rollback("상점 DB처리 실패로 인하여 Rollback 처리 [TID:" & xpay.Response("LGD_TID",0) & ",MID:" & xpay.Response("LGD_MID",0) & ",OID:" & xpay.Response("LGD_OID",0) & "]")
            			
				IF "0000" = xpay.resCode THEN
						Call SettleErrorLogWrite(LGD_OID, "Y", "PR03", "PayRes", "DB커넥션, 레코드셋 개체 생성 오류 / " & GetPayType(PayType) & " 취소 완료", Err.Description)
						Call AlertMessage2("주문 처리 도중 오류가 발생하여 결제를 취소하였습니다.[PR03]\n\n다시 주문 시도하여 주십시오.", "location.replace('/ASP/Mypage/OrderList.asp');")
				ELSE
						Call SettleErrorLogWrite(LGD_OID, "N", "PR04", "PayRes", "DB커넥션, 레코드셋 개체 생성 오류 / " & GetPayType(PayType) & " 취소 오류", Err.Description)
						Call AlertMessage2("주문 처리 도중 오류가 발생하였습니다.[PR04]\n\n결제 취소가 정상적으로 처리되지 않았습니다.\n\n관리자에게 문의 바랍니다.", "location.replace('/ASP/Mypage/OrderList.asp');")
				END IF
				Response.End
		END IF



		'#IF xpay.resCode = "0000" THEN
		IF LGD_RESPCODE = "0000" THEN

				DIM DBPayType
				Set oCmd = Server.CreateObject("ADODB.Command")
				WITH oCmd
						.ActiveConnection = oConn
						.CommandType = adCmdStoredProc
						.CommandText = "USP_Front_EShop_Order_Product_Cancel_Temp_Select_By_Idx"
						.Parameters.Append .CreateParameter("@Idx",		adInteger,	adParamInput,	,		Replace(LGD_OID, "OPC", ""))
				END WITH
				oRs.CursorLocation = adUseClient
				oRs.Open oCmd, , adOpenStatic, adLockReadOnly
				Set oCmd = Nothing
	
				IF Err.number <> 0 THEN

						oRs.Close
						SET oRs = Nothing
						oConn.Close
						SET oConn = Nothing
	
						isDBOK = false

						xpay.Rollback("상점 DB처리 실패로 인하여 Rollback 처리 [TID:" & xpay.Response("LGD_TID",0) & ",MID:" & xpay.Response("LGD_MID",0) & ",OID:" & xpay.Response("LGD_OID",0) & "]")
            			
						IF "0000" = xpay.resCode THEN
								Call SettleErrorLogWrite(LGD_OID, "Y", "PR09", "PayRes", "EShop_Order Select 오류 / " & GetPayType(PayType) & " 취소 완료", Err.Description)
								Call AlertMessage2("주문 처리 도중 오류가 발생하여 결제를 취소하였습니다.[PR09]\n\n다시 주문 시도하여 주십시오.", "location.replace('/ASP/Mypage/OrderList.asp');")
						ELSE
								Call SettleErrorLogWrite(LGD_OID, "N", "PR10", "PayRes", "EShop_Order Select 오류 / " & GetPayType(PayType) & " 취소 오류", Err.Description)
								Call AlertMessage2("주문 처리 도중 오류가 발생하였습니다.[PR10]\n\n결제 취소가 정상적으로 처리되지 않았습니다.\n\n관리자에게 문의 바랍니다.", "location.replace('/ASP/Mypage/OrderList.asp');")
						END IF
						Response.End
				END IF

				IF NOT oRs.EOF THEN
						SELECT CASE oRs("DelvFeeType")
								CASE "6" : DBPayType = "C"
								CASE "3" : DBPayType = "B"
								CASE ELSE : DBPayType = oRs("DelvFeeType")
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
						SET oRs = Nothing
						oConn.Close
						SET oConn = Nothing
	
						isDBOK = false

						xpay.Rollback("상점 DB처리 실패로 인하여 Rollback 처리 [TID:" & xpay.Response("LGD_TID",0) & ",MID:" & xpay.Response("LGD_MID",0) & ",OID:" & xpay.Response("LGD_OID",0) & "]")
            			
						IF "0000" = xpay.resCode THEN
								Call SettleErrorLogWrite(LGD_OID, "Y", "PR09", "PayRes", "EShop_Order Select 오류 / " & GetPayType(PayType) & " 취소 완료", Err.Description)
								Call AlertMessage2("배송비 결제 처리 도중 오류가 발생하여 결제를 취소하였습니다.[PR09]\n\n다시 주문 시도하여 주십시오.", "location.replace('/ASP/Mypage/OrderList.asp');")
						ELSE
								Call SettleErrorLogWrite(LGD_OID, "N", "PR10", "PayRes", "EShop_Order Select 오류 / " & GetPayType(PayType) & " 취소 오류", Err.Description)
								Call AlertMessage2("배송비 결제 처리 도중 오류가 발생하였습니다.[PR10]\n\n결제 취소가 정상적으로 처리되지 않았습니다.\n\n관리자에게 문의 바랍니다.", "location.replace('/ASP/Mypage/OrderList.asp');")
						END IF
						Response.End
				END IF
				oRs.Close


				'최종결제요청 결과 성공 DB처리
				'Response.Write("최종결제요청 결과 성공 DB처리하시기 바랍니다." & "<br>")
            	            	            	
				'최종결제요청 결과 성공 DB처리 실패시 Rollback 처리
				isDBOK = true 'DB처리 실패시 false로 변경해 주세요.
            	

				'-----------------------------------------------------------------------------------------------------------'
				'DB에 있는 결제수단과 PG사에서 넘어온 결제수단이 다르면 취소 START
				'-----------------------------------------------------------------------------------------------------------'	
				IF Trim(PayType) <> Trim(DBPayType) THEN
						Set oRs = Nothing
						oConn.Close
						Set oConn = Nothing

						isDBOK = false

						xpay.Rollback("상점 DB처리 실패로 인하여 Rollback 처리 [TID:" & xpay.Response("LGD_TID",0) & ",MID:" & xpay.Response("LGD_MID",0) & ",OID:" & xpay.Response("LGD_OID",0) & "]")
            		
						IF "0000" = xpay.resCode THEN
								Call SettleErrorLogWrite(LGD_OID, "Y", "PR13", "PayRes", "결제수단상이 결:" & GetPayType(PayType) & " / 주:" & GetPayType(DBPayType) & " 취소 완료", Err.Description)
								Call AlertMessage2("배송비 결제 처리 도중 오류가 발생하였습니다. 결제 자동취소가 정상적으로 완료 되었습니다. 다시 시도하여 주십시오.", "location.replace('/ASP/Mypage/OrderList.asp');")
								Response.End
						ELSE
								Call SettleErrorLogWrite(LGD_OID, "N", "PR14", "PayRes", "결제수단상이 결:" & GetPayType(PayType) & " / 주:" & GetPayType(DBPayType) & " 취소 오류", Err.Description)
								Call AlertMessage2("배송비 결제 처리 도중 오류가 발생하였습니다. 결제 자동취소가 정상적으로 처리되지 않았습니다. 관리자에게 문의 바랍니다.", "location.replace('/ASP/Mypage/OrderList.asp');")
								Response.End
						END IF
				END IF
				'-----------------------------------------------------------------------------------------------------------'
				'DB에 있는 결제수단과 PG사에서 넘어온 결제수단이 다르면 취소 START
				'-----------------------------------------------------------------------------------------------------------'	

				'-----------------------------------------------------------------------------------------------------------'
				'DB에 있는 금액과 PG사에서 넘어온 결재금액이 다르면 취소 START
				'-----------------------------------------------------------------------------------------------------------'	
				IF CDbl(LGD_AMOUNT) <> CDbl(DelvFee) THEN
						Set oRs = Nothing
						oConn.Close
						Set oConn = Nothing

						isDBOK = false

						xpay.Rollback("상점 DB처리 실패로 인하여 Rollback 처리 [TID:" & xpay.Response("LGD_TID",0) & ",MID:" & xpay.Response("LGD_MID",0) & ",OID:" & xpay.Response("LGD_OID",0) & "]")
            		
						IF "0000" = xpay.resCode THEN
								Call SettleErrorLogWrite(LGD_OID, "Y", "PR15", "PayRes", "결제금액상이 결:" & CDbl(LGD_AMOUNT) & " / 주:" & CDbl(DelvFee) & " / " & GetPayType(PayType) & " 취소 완료", Err.Description)
								Call AlertMessage2("배송비 결제 처리 도중 오류가 발생하였습니다. 결제 자동취소가 정상적으로 완료 되었습니다. 다시 시도하여 주십시오.", "location.replace('/ASP/Mypage/OrderList.asp');")
								Response.End
						ELSE
								Call SettleErrorLogWrite(LGD_OID, "N", "PR16", "PayRes", "결제금액상이 결:" & CDbl(LGD_AMOUNT) & " / 주:" & CDbl(DelvFee) & " / " & GetPayType(PayType) & " 취소 오류", Err.Description)
								Call AlertMessage2("배송비 결제 처리 도중 오류가 발생하였습니다. 결제 자동취소가 정상적으로 처리되지 않았습니다. 관리자에게 문의 바랍니다.", "location.replace('/ASP/Mypage/OrderList.asp');")
								Response.End
						END IF
				END IF
				'-----------------------------------------------------------------------------------------------------------'	
				'DB에 있는 금액과 PG사에서 넘어온 결재금액이 다르면 취소 END
				'-----------------------------------------------------------------------------------------------------------'


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

				IF Err.Number <> 0 THEN
						oConn.RollbackTrans

						Set oRs = Nothing
						oConn.Close
						Set oConn = Nothing

						isDBOK = false

						xpay.Rollback("상점 DB처리 실패로 인하여 Rollback 처리 [TID:" & xpay.Response("LGD_TID",0) & ",MID:" & xpay.Response("LGD_MID",0) & ",OID:" & xpay.Response("LGD_OID",0) & "]")
            		
						IF "0000" = xpay.resCode THEN
								Call SettleErrorLogWrite(LGD_OID, "Y", "PR17", "PayRes", "EShop_Order 결제정보 업데이트 오류 / " & GetPayType(PayType) & " 취소 완료", Err.Description)
								Call AlertMessage2("배송비 결제 처리 도중 오류가 발생하였습니다. 결제 자동취소가 정상적으로 완료 되었습니다. 다시 시도하여 주십시오.", "location.replace('/ASP/Mypage/OrderList.asp');")
								Response.End
						ELSE
								Call SettleErrorLogWrite(LGD_OID, "N", "PR18", "PayRes", "EShop_Order 결제정보 업데이트 오류 / " & GetPayType(PayType) & " 취소 오류", Err.Description)
								Call AlertMessage2("배송비 결제 처리 도중 오류가 발생하였습니다. 결제 자동취소가 정상적으로 처리되지 않았습니다. 관리자에게 문의 바랍니다.", "location.replace('/ASP/Mypage/OrderList.asp');")
								Response.End
						END IF
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

						isDBOK = false

						xpay.Rollback("상점 DB처리 실패로 인하여 Rollback 처리 [TID:" & xpay.Response("LGD_TID",0) & ",MID:" & xpay.Response("LGD_MID",0) & ",OID:" & xpay.Response("LGD_OID",0) & "]")
            		
						IF "0000" = xpay.resCode THEN
								Call SettleErrorLogWrite(LGD_OID, "Y", "PR35", "PayRes", "EShop_Order_Settle 입력 오류 / " & GetPayType(PayType) & " 취소 완료", Err.Description)
								Call AlertMessage2("배송비 결제 처리 도중 오류가 발생하였습니다. 결제 자동취소가 정상적으로 완료 되었습니다. 다시 시도하여 주십시오.", "location.replace('/ASP/Mypage/OrderList.asp');")
								Response.End
						ELSE
								Call SettleErrorLogWrite(LGD_OID, "Y", "PR35", "PayRes", "EShop_Order_Settle 입력 오류 / " & GetPayType(PayType) & " 취소 오류", Err.Description)
								Call AlertMessage2("배송비 결제 처리 도중 오류가 발생하였습니다. 결제 자동취소가 정상적으로 처리되지 않았습니다. 관리자에게 문의 바랍니다.", "location.replace('/ASP/Mypage/OrderList.asp');")
								Response.End
						END IF
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
						SET oRs = Nothing : oConn.Close : SET oConn = Nothing

						isDBOK = false

						xpay.Rollback("상점 DB처리 실패로 인하여 Rollback 처리 [TID:" & xpay.Response("LGD_TID",0) & ",MID:" & xpay.Response("LGD_MID",0) & ",OID:" & xpay.Response("LGD_OID",0) & "]")
            		
						IF "0000" = xpay.resCode THEN
								Call SettleErrorLogWrite(LGD_OID, "Y", "PR35", "PayRes", "IF_WMS_RETURNREQUEST_H 입력 오류 / " & GetPayType(PayType) & " 취소 완료", Err.Description)
								Call AlertMessage2("배송비 결제 처리 도중 오류가 발생하였습니다. 결제 자동취소가 정상적으로 완료 되었습니다. 다시 시도하여 주십시오.", "location.replace('/ASP/Mypage/OrderList.asp');")
								Response.End
						ELSE
								Call SettleErrorLogWrite(LGD_OID, "Y", "PR35", "PayRes", "IF_WMS_RETURNREQUEST_H 입력 오류 / " & GetPayType(PayType) & " 취소 오류", Err.Description)
								Call AlertMessage2("배송비 결제 처리 도중 오류가 발생하였습니다. 결제 자동취소가 정상적으로 처리되지 않았습니다. 관리자에게 문의 바랍니다.", "location.replace('/ASP/Mypage/OrderList.asp');")
								Response.End
						END IF
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

						isDBOK = false

						xpay.Rollback("상점 DB처리 실패로 인하여 Rollback 처리 [TID:" & xpay.Response("LGD_TID",0) & ",MID:" & xpay.Response("LGD_MID",0) & ",OID:" & xpay.Response("LGD_OID",0) & "]")
            		
						IF "0000" = xpay.resCode THEN
								Call SettleErrorLogWrite(LGD_OID, "Y", "PR35", "PayRes", "IF_WMS_RETURNREQUEST_D 입력 오류 / " & GetPayType(PayType) & " 취소 완료", Err.Description)
								Call AlertMessage2("배송비 결제 처리 도중 오류가 발생하였습니다. 결제 자동취소가 정상적으로 완료 되었습니다. 다시 시도하여 주십시오.", "location.replace('/ASP/Mypage/OrderList.asp');")
								Response.End
						ELSE
								Call SettleErrorLogWrite(LGD_OID, "Y", "PR35", "PayRes", "IF_WMS_RETURNREQUEST_D 입력 오류 / " & GetPayType(PayType) & " 취소 오류", Err.Description)
								Call AlertMessage2("배송비 결제 처리 도중 오류가 발생하였습니다. 결제 자동취소가 정상적으로 처리되지 않았습니다. 관리자에게 문의 바랍니다.", "location.replace('/ASP/Mypage/OrderList.asp');")
								Response.End
						END IF
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

						isDBOK = false

						xpay.Rollback("상점 DB처리 실패로 인하여 Rollback 처리 [TID:" & xpay.Response("LGD_TID",0) & ",MID:" & xpay.Response("LGD_MID",0) & ",OID:" & xpay.Response("LGD_OID",0) & "]")
            		
						IF "0000" = xpay.resCode THEN
								Call SettleErrorLogWrite(LGD_OID, "Y", "PR35", "PayRes", "EShop_Order_Product_Change_History 입력 오류 / " & GetPayType(PayType) & " 취소 완료", Err.Description)
								Call AlertMessage2("배송비 결제 처리 도중 오류가 발생하였습니다. 결제 자동취소가 정상적으로 완료 되었습니다. 다시 시도하여 주십시오.", "location.replace('/ASP/Mypage/OrderList.asp');")
								Response.End
						ELSE
								Call SettleErrorLogWrite(LGD_OID, "Y", "PR35", "PayRes", "EShop_Order_Product_Change_History 입력 오류 / " & GetPayType(PayType) & " 취소 오류", Err.Description)
								Call AlertMessage2("배송비 결제 처리 도중 오류가 발생하였습니다. 결제 자동취소가 정상적으로 처리되지 않았습니다. 관리자에게 문의 바랍니다.", "location.replace('/ASP/Mypage/OrderList.asp');")
								Response.End
						END IF
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



				IF isDBOK THEN
%>
						<script type="text/javascript">
							//alert("결제가 정상적으로 처리되었습니다.");
							location.replace('/ASP/Mypage/OrderList.asp');
						</script>
<%
						Response.End
				ELSE
						'# 여기는 들어올 경우가 없다(위에서 오류날 경우 바로 결제취소후 종료 시킨다)

           				'Response.Write("<p>")
           				'xpay.Rollback("상점 DB처리 실패로 인하여 Rollback 처리 [TID:" & xpay.Response("LGD_TID",0) & ",MID:" & xpay.Response("LGD_MID",0) & ",OID:" & xpay.Response("LGD_OID",0) & "]")
            		
						'Response.Write("TX Rollback Response_code = " & xpay.resCode & "<br>")
						'Response.Write("TX Rollback Response_msg = " & xpay.resMsg & "<p>")
            		
						IF "0000" = xpay.resCode THEN
								'#Response.Write("자동취소가 정상적으로 완료 되었습니다.<br>")
								Call AlertMessage2("주문 처리 도중 오류가 발생하였습니다. 결제 자동취소가 정상적으로 완료 되었습니다. 다시 시도하여 주십시오.", "location.replace('/ASP/Mypage/OrderList.asp');")
								Response.End
						ELSE
								'#Response.Write("자동취소가 정상적으로 처리되지 않았습니다.<br>")
								Call AlertMessage2("주문 처리 도중 오류가 발생하였습니다. 결제 자동취소가 정상적으로 처리되지 않았습니다. 관리자에게 문의 바랍니다.", "location.replace('/ASP/Mypage/OrderList.asp');")
								Response.End
						END IF
				END IF


		ELSE
				'결제결제요청 결과 실패 DB처리
				'#Response.Write("결제결제요청 결과 실패 DB처리하시기 바랍니다." & "<br>")
				'#Response.Write("TX Response_code = " & xpay.resCode & "<br>")
				'#Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")


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
				ELSE
						oConn.CommitTrans
				END IF
				'-----------------------------------------------------------------------------------------------------------'	
				'결제 정보 저장 End
				'-----------------------------------------------------------------------------------------------------------'	

				Set oRs = Nothing
				oConn.Close
				Set oConn = Nothing

				Call AlertMessage2("결제가 정상적으로 이루어지지 않았습니다.[Code:0001] 다시 시도하여 주십시오.", "location.replace('/ASP/Mypage/OrderList.asp');")
				Response.End
		END IF
ELSE
		'2)API 요청실패 화면처리
		'#Response.Write("결제요청이 실패하였습니다. <br>")
		'#Response.Write("TX Response_code = " & xpay.resCode & "<br>")
		'#Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")
            
		'결제요청 결과 실패 상점 DB처리
		'#Response.Write("결제결제요청 결과 실패 DB처리하시기 바랍니다." & "<br>")

		Call AlertMessage2("결제가 정상적으로 이루어지지 않았습니다.[Code:0002] 다시 시도하여 주십시오.", "location.replace('/ASP/Mypage/OrderList.asp');")
		Response.End
END IF
 

Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>
