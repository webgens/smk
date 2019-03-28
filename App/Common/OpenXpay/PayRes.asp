<%@ Language=VBScript codepage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************/
'PayRes.asp - 카드결제(안심결제) 결과 처리 및 리턴 / 가상계좌 리턴 페이지
'Date		: 2018.12.30
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
<body> <!-- oncontextmenu="return false" onselectstart="return false" ondragstart="return false">
	<table cellpadding="0" cellspacing="0" width="100%" height="100%">
		<tr>
			<td align="center" valign="middle"><img src="<%=HOME_URL%>/Images/loading.gif" width="100" alt="LOADING" /></td>
		</tr>
	</table>-->
</body>
</html>

<%
'Response.Flush


'# 결제 오류시 로그 데이터
SUB SettleErrorLogWrite(ByVal orderCode, ByVal cancelFlag, ByVal errCode, ByVal errPage, ByVal errMsg, ByVal errDesc)

		'ON ERROR RESUME NEXT

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




'ON ERROR RESUME NEXT

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

		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=주문 처리 도중 오류가 발생하였습니다.[PR01]<br />결제가 정상 처리 되지 않았습니다.<br />관리자에게 문의 바랍니다.&Script=APP_PopupHistoryBack();"
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

		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=주문 처리 도중 오류가 발생하였습니다.[PR02]<br />결제가 정상 처리 되지 않았습니다.<br />관리자에게 문의 바랍니다.&Script=APP_PopupHistoryBack();"
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



DIM OrderPrice
DIM OrderDate
DIM OrderTime
DIM OrderState
DIM SettleFlag
DIM SettleDate
DIM SettleTime
DIM CasFlag
DIM PayType

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
	
						Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=주문 처리 도중 오류가 발생하여 결제를 취소하였습니다.[PR03]<br />다시 주문 시도하여 주십시오.&Script=APP_PopupHistoryBack();"
				ELSE
						Call SettleErrorLogWrite(LGD_OID, "N", "PR04", "PayRes", "DB커넥션, 레코드셋 개체 생성 오류 / " & GetPayType(PayType) & " 취소 오류", Err.Description)
	
						Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=주문 처리 도중 오류가 발생하였습니다.[PR04]<br />결제 취소가 정상적으로 처리되지 않았습니다.<br />관리자에게 문의 바랍니다.&Script=APP_PopupHistoryBack();"
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
						.CommandText = "USP_Admin_EShop_Order_Select_By_OrderCode"
						.Parameters.Append .CreateParameter("@OrderCode",		adVarChar,	adParamInput,	20,		LGD_OID)
				END WITH
				oRs.Open oCmd, , adOpenStatic, adLockReadOnly
				Set oCmd = Nothing
	
				IF Err.number <> 0 THEN
						SET oRs = Nothing
						oConn.Close
						SET oConn = Nothing
	
						isDBOK = false

						xpay.Rollback("상점 DB처리 실패로 인하여 Rollback 처리 [TID:" & xpay.Response("LGD_TID",0) & ",MID:" & xpay.Response("LGD_MID",0) & ",OID:" & xpay.Response("LGD_OID",0) & "]")
            			
						IF "0000" = xpay.resCode THEN
								Call SettleErrorLogWrite(LGD_OID, "Y", "PR09", "PayRes", "EShop_Order Select 오류 / " & GetPayType(PayType) & " 취소 완료", Err.Description)

								Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=주문 처리 도중 오류가 발생하여 결제를 취소하였습니다.[PR09]<br />다시 주문 시도하여 주십시오.&Script=APP_PopupHistoryBack();"
						ELSE
								Call SettleErrorLogWrite(LGD_OID, "N", "PR10", "PayRes", "EShop_Order Select 오류 / " & GetPayType(PayType) & " 취소 오류", Err.Description)

								Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=주문 처리 도중 오류가 발생하였습니다.[PR10]<br />결제 취소가 정상적으로 처리되지 않았습니다.<br />관리자에게 문의 바랍니다.&Script=APP_PopupHistoryBack();"
						END IF
						Response.End
				END IF

				IF NOT oRs.EOF THEN
						U_NUM			= oRs("UserID")
						U_NAME			= oRs("OrderName")
						DBPayType		= oRs("PayType")
						OrderPrice		= oRs("OrderPrice")
						OrderDate		= oRs("OrderDate")
						OrderTime		= oRs("OrderTime")
				END IF
				oRs.Close


				IF PayType = "V" THEN
						OrderState		= "1"
						SettleFlag		= "N"
						SettleDate		= ""
						SettleTime		= ""
						CasFlag			= LGD_CASFLAG
				ELSE
						OrderState		= "3"
						SettleFlag		= "Y"
						SettleDate		= U_DATE
						SettleTime		= U_TIME
						CasFlag			= ""
				END IF


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

								Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=주문 처리 도중 오류가 발생하였습니다.<br />결제 자동취소가 정상적으로 완료 되었습니다.<br />다시 시도하여 주십시오.&Script=APP_PopupHistoryBack();"
								Response.End
						ELSE
								Call SettleErrorLogWrite(LGD_OID, "N", "PR14", "PayRes", "결제수단상이 결:" & GetPayType(PayType) & " / 주:" & GetPayType(DBPayType) & " 취소 오류", Err.Description)

								Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=주문 처리 도중 오류가 발생하였습니다.<br />결제 자동취소가 정상적으로 처리되지 않았습니다.<br />관리자에게 문의 바랍니다.<br />다시 시도하여 주십시오.&Script=APP_PopupHistoryBack();"
								Response.End
						END IF
				END IF
				'-----------------------------------------------------------------------------------------------------------'
				'DB에 있는 결제수단과 PG사에서 넘어온 결제수단이 다르면 취소 START
				'-----------------------------------------------------------------------------------------------------------'	

				'-----------------------------------------------------------------------------------------------------------'
				'DB에 있는 금액과 PG사에서 넘어온 결재금액이 다르면 취소 START
				'-----------------------------------------------------------------------------------------------------------'	
				IF CDbl(LGD_AMOUNT) <> CDbl(OrderPrice) THEN
						Set oRs = Nothing
						oConn.Close
						Set oConn = Nothing

						isDBOK = false

						xpay.Rollback("상점 DB처리 실패로 인하여 Rollback 처리 [TID:" & xpay.Response("LGD_TID",0) & ",MID:" & xpay.Response("LGD_MID",0) & ",OID:" & xpay.Response("LGD_OID",0) & "]")
            		
						IF "0000" = xpay.resCode THEN
								Call SettleErrorLogWrite(LGD_OID, "Y", "PR15", "PayRes", "결제금액상이 결:" & CDbl(LGD_AMOUNT) & " / 주:" & CDbl(OrderPrice) & " / " & GetPayType(PayType) & " 취소 완료", Err.Description)

								Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=주문 처리 도중 오류가 발생하였습니다.<br />결제 자동취소가 정상적으로 완료 되었습니다.<br />다시 시도하여 주십시오.&Script=APP_PopupHistoryBack();"
								Response.End
						ELSE
								Call SettleErrorLogWrite(LGD_OID, "N", "PR16", "PayRes", "결제금액상이 결:" & CDbl(LGD_AMOUNT) & " / 주:" & CDbl(OrderPrice) & " / " & GetPayType(PayType) & " 취소 오류", Err.Description)
	
								Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=주문 처리 도중 오류가 발생하였습니다.<br />결제 자동취소가 정상적으로 처리되지 않았습니다.<br />관리자에게 문의 바랍니다.&Script=APP_PopupHistoryBack();"
								Response.End
						END IF
				END IF
				'-----------------------------------------------------------------------------------------------------------'	
				'DB에 있는 금액과 PG사에서 넘어온 결재금액이 다르면 취소 END
				'-----------------------------------------------------------------------------------------------------------'


				oConn.BeginTrans	


				'-----------------------------------------------------------------------------------------------------------'	
				'USafe 보증보험 발급처리 시작
				'-----------------------------------------------------------------------------------------------------------'	
				IF USAFE_GuaranteeInsurance = "Y" THEN
						DIM wQuery
						DIM sQuery

						DIM USafeCom
						DIM UsafeResult
						DIM UsafeResultCode : UsafeResultCode = "0"
						DIM UsafeResultMsg

						SET USafeCom		= CreateObject( "USafeCom.guarantee.1")
						' Real
						USafeCom.Port		= 80
						USafeCom.Url		= "gateway.usafe.co.kr"
						USafeCom.CallForm	= "/esafe/guartrn.asp"

						'데이터 64Bit 암호화시 사용
						USafeCom.EncKey		= "uclick"						'널값인 경우 암호화 안됨

						'//주문정보 조회 시작
						wQuery	= "WHERE A.IsShowFlag = 'Y' AND A.SaleType = 'P' AND A.ProductType = 'P' AND A.OrderCode = '" & LGD_OID & "' "
						sQuery	= "ORDER BY A.Idx "

						Set oCmd = Server.CreateObject("ADODB.Command")
						WITH oCmd
								.ActiveConnection = oConn
								.CommandType = adCmdStoredProc
								.CommandText = "USP_Front_EShop_Order_Product_Select_For_Order_Detail"

								.Parameters.Append .CreateParameter("@WQUERY", adVarChar, adParamInput, 1000, wQuery)
								.Parameters.Append .CreateParameter("@SQUERY", adVarChar, adParamInput,  100, sQuery)
						END WITH
						oRs.CursorLocation = adUseClient
						oRs.Open oCmd, , adOpenStatic, adLockReadOnly
						Set oCmd = Nothing
						'//주문정보 조회 끝

						IF oRs.BOF AND oRs.EOF THEN
								USafeCom.goodsCount				 =  0
								USafeCom.AddGoods				 ""
								USafeCom.AddGoodsPrice			 ""
								USafeCom.AddGoodsCnt			 ""
						ELSE
								USafeCom.goodsCount				 =  1			'상품종류수에 맞게 아래 상품내역들을 맞춰주셔야 합니다.
								If oRs.RecordCount = 1 Then
										USafeCom.AddGoods		 oRs("ProductCD")
										USafeCom.AddGoodsPrice	 LGD_AMOUNT
										USafeCom.AddGoodsCnt	 oRs.RecordCount
								Else
										USafeCom.AddGoods		 oRs("ProductCD") & "외 " & oRs.RecordCount - 1 & " 품목"
										USafeCom.AddGoodsPrice	 LGD_AMOUNT
										USafeCom.AddGoodsCnt	 oRs.RecordCount
								End If
						END IF
						oRs.Close
					

						USafeCom.gubun				 =  "A0"
						UsafeCom.mallId				 =  USAFE_ID
						UsafeCom.oId				 =  LGD_OID										'// 상점의 주문번호
						UsafeCom.totalMoney			 =  LGD_AMOUNT
						UsafeCom.pId				 =  USAFE_JuminNumber
						IF PayType = "V" THEN
								UsafeCom.payMethod			 =  "CAS"										'//결제방식(가상계좌)
						ELSEIF PayType = "B" THEN
								UsafeCom.payMethod			 =  "BMC"										'//결제방식(계좌이체)
						END IF
						UsafeCom.payInfo1			 =  LGD_FINANCENAME
						UsafeCom.payInfo2			 =  LGD_ACCOUNTNUM
						UsafeCom.orderNm			 =  LGD_BUYER
						UsafeCom.orderHomeTel		 =  ""
						UsafeCom.orderHpTel			 =  LGD_BUYERPHONE
						UsafeCom.orderZip			 =  TRIM(MID(LGD_BUYERADDRESS, 1, 6))
						UsafeCom.orderAddress		 =  TRIM(MID(LGD_BUYERADDRESS, 7))
						UsafeCom.orderEmail			 =  LGD_BUYEREMAIL
						UsafeCom.acceptor			 =  LGD_RECEIVER
						UsafeCom.deliveryTel1		 =  LGD_RECEIVERPHONE
						UsafeCom.deliveryTel2		 =  ""
						UsafeCom.sign				 =  "Y" & USAFE_EmailFlag & USAFE_SmsFlag		'// 개인정보동의(1) Email수신동의(2) SMS수신동의(3)
						UsafeCom.serviceid			 =	""											'// 옵션(전자보증쇼핑몰관련)
						UsafeCom.catecode			 =	""											'//옵션(전자보증쇼핑몰관련)

						UsafeResult					 = UsafeCom.contractInsurance
						UsafeResultCode				 = Left( UsafeResult , 1 )
						UsafeResultMsg				 = Mid( UsafeResult , 3 )


						SET UsafeCom = Nothing


						IF CStr(UsafeResultCode) <> "0" THEN
								oConn.RollbackTrans
						
								'# 보증보험 로그 생성
								Set oCmd = Server.CreateObject("ADODB.Command")
								WITH oCmd
										.ActiveConnection = oConn
										.CommandType = adCmdStoredProc
										.CommandText = "USP_Admin_EShop_Usafe_Log_Insert"
										.Parameters.Append .CreateParameter("@OrderCode",					adVarChar,		adParamInput,	  20,		LGD_OID)
										.Parameters.Append .CreateParameter("@UsafeGubun",					adChar,			adParamInput,	   2,		"A0")
										.Parameters.Append .CreateParameter("@UsafeResultCode",				adVarChar,		adParamInput,	  50,		UsafeResultCode)
										.Parameters.Append .CreateParameter("@UsafeResultMsg",				adVarChar,		adParamInput,	1000,		Replace(UsafeResult, "'", ""))
										.Parameters.Append .CreateParameter("@U_MEMNUM",					adVarChar,		adParamInput,	  50,		U_NUM)
										.Parameters.Append .CreateParameter("@U_IP",						adVarChar,		adParamInput,	  15,		U_IP)

										.Execute, , adExecuteNoRecords
								END WITH
								Set oCmd = Nothing

								Set oRs = Nothing
								oConn.Close
								Set oConn = Nothing
						
						'#		resultMSG = "보증보험 발급에 실패하였습니다. 다시 주문하여 주십시오"
						'#		Response.Write resultMSG
						'#		Response.End

								isDBOK = false

								xpay.Rollback("상점 DB처리 실패로 인하여 Rollback 처리 [TID:" & xpay.Response("LGD_TID",0) & ",MID:" & xpay.Response("LGD_MID",0) & ",OID:" & xpay.Response("LGD_OID",0) & "]")
            		
								IF "0000" = xpay.resCode THEN
										Call SettleErrorLogWrite(LGD_OID, "Y", "PR05", "PayRes", "EShop_Order 보증보험 발급 오류 / " & GetPayType(PayType) & " 취소 완료", Err.Description)
	
										Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=주문 처리 도중 오류가 발생하였습니다.<br />결제 자동취소가 정상적으로 완료 되었습니다.<br />다시 시도하여 주십시오.&Script=APP_PopupHistoryBack();"
										Response.End
								ELSE
										Call SettleErrorLogWrite(LGD_OID, "N", "PR06", "PayRes", "EShop_Order 보증보험 발급 오류 / " & GetPayType(PayType) & " 취소 오류", Err.Description)
	
										Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=주문 처리 도중 오류가 발생하였습니다.<br />결제 자동취소가 정상적으로 처리되지 않았습니다.<br />관리자에게 문의 바랍니다.&Script=APP_PopupHistoryBack();"
										Response.End
								END IF
						END IF
						'# 보증보험 로그 생성
						Set oCmd = Server.CreateObject("ADODB.Command")
						WITH oCmd
								.ActiveConnection = oConn
								.CommandType = adCmdStoredProc
								.CommandText = "USP_Admin_EShop_Usafe_Log_Insert"
								.Parameters.Append .CreateParameter("@OrderCode",					adVarChar,		adParamInput,	  20,		LGD_OID)
								.Parameters.Append .CreateParameter("@UsafeGubun",					adChar,			adParamInput,	   2,		"A0")
								.Parameters.Append .CreateParameter("@UsafeResultCode",				adVarChar,		adParamInput,	  50,		UsafeResultCode)
								.Parameters.Append .CreateParameter("@UsafeResultMsg",				adVarChar,		adParamInput,	1000,		Replace(UsafeResult, "'", ""))
								.Parameters.Append .CreateParameter("@U_MEMNUM",					adVarChar,		adParamInput,	  50,		U_NUM)
								.Parameters.Append .CreateParameter("@U_IP",						adVarChar,		adParamInput,	  15,		U_IP)

								.Execute, , adExecuteNoRecords
						END WITH
						Set oCmd = Nothing


						'# 주문에 보증보험 결과 저장
						Set oCmd = Server.CreateObject("ADODB.Command")
						WITH oCmd
								.ActiveConnection = oConn
								.CommandType = adCmdStoredProc
								.CommandText = "USP_Front_EShop_Order_Update_For_USafe"
								.Parameters.Append .CreateParameter("@OrderCode",					adInteger,		adParamInput,	,		LGD_OID)
								.Parameters.Append .CreateParameter("@GuaranteeInsurance",			adChar,			adParamInput,	1,		USAFE_GuaranteeInsurance)
								.Parameters.Append .CreateParameter("@GuaranteeInsuranceGubun",		adChar,			adParamInput,	2,		"A0")
								.Parameters.Append .CreateParameter("@GuaranteeInsuranceResult",	adVarChar,		adParamInput,	100,	UsafeResultMsg)
								.Parameters.Append .CreateParameter("@U_MEMNUM",					adVarChar,		adParamInput,	20,		U_NUM)
								.Parameters.Append .CreateParameter("@U_IP",						adVarChar,		adParamInput,	15,		U_IP)

								.Execute, , adExecuteNoRecords
						END WITH
						Set oCmd = Nothing
							
						IF Err.number <> 0 THEN
								oConn.RollbackTrans
						
								Set oRs = Nothing
								oConn.Close
								Set oConn = Nothing
						
						'		resultMSG = "보증보험 발급 도중 에러가 발생하였습니다. 다시 주문하여 주십시오"
						'		Response.Write resultMSG
						'		Response.End

								isDBOK = false

								xpay.Rollback("상점 DB처리 실패로 인하여 Rollback 처리 [TID:" & xpay.Response("LGD_TID",0) & ",MID:" & xpay.Response("LGD_MID",0) & ",OID:" & xpay.Response("LGD_OID",0) & "]")
            		
								IF "0000" = xpay.resCode THEN
										Call SettleErrorLogWrite(LGD_OID, "Y", "PR07", "PayRes", "EShop_Order 보증보험정보 업데이트 오류 / " & GetPayType(PayType) & " 취소 완료", Err.Description)
	
										Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=주문 처리 도중 오류가 발생하였습니다.<br />결제 자동취소가 정상적으로 완료 되었습니다.<br />다시 시도하여 주십시오.&Script=APP_PopupHistoryBack();"
										Response.End
								ELSE
										Call SettleErrorLogWrite(LGD_OID, "N", "PR08", "PayRes", "EShop_Order 보증보험정보 업데이트 오류 / " & GetPayType(PayType) & " 취소 오류", Err.Description)
	
										Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=주문 처리 도중 오류가 발생하였습니다.<br />결제 자동취소가 정상적으로 처리되지 않았습니다.<br />관리자에게 문의 바랍니다.&Script=APP_PopupHistoryBack();"
										Response.End
								END IF
						END IF
				END IF
				'-----------------------------------------------------------------------------------------------------------'	
				'USafe 보증보험 발급처리 끝
				'-----------------------------------------------------------------------------------------------------------'	



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

						Set oRs = Nothing
						oConn.Close
						Set oConn = Nothing

						isDBOK = false

						xpay.Rollback("상점 DB처리 실패로 인하여 Rollback 처리 [TID:" & xpay.Response("LGD_TID",0) & ",MID:" & xpay.Response("LGD_MID",0) & ",OID:" & xpay.Response("LGD_OID",0) & "]")
            		
						IF "0000" = xpay.resCode THEN
								Call SettleErrorLogWrite(LGD_OID, "Y", "PR17", "PayRes", "EShop_Order 결제정보 업데이트 오류 / " & GetPayType(PayType) & " 취소 완료", Err.Description)
	
								Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=주문 처리 도중 오류가 발생하였습니다.<br />결제 자동취소가 정상적으로 완료 되었습니다.<br />다시 시도하여 주십시오.&Script=APP_PopupHistoryBack();"
								Response.End
						ELSE
								Call SettleErrorLogWrite(LGD_OID, "N", "PR18", "PayRes", "EShop_Order 결제정보 업데이트 오류 / " & GetPayType(PayType) & " 취소 오류", Err.Description)
	
								Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=주문 처리 도중 오류가 발생하였습니다.<br />결제 자동취소가 정상적으로 처리되지 않았습니다.<br />관리자에게 문의 바랍니다.&Script=APP_PopupHistoryBack();"
								Response.End
						END IF
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
	
								Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=주문 처리 도중 오류가 발생하였습니다.<br />결제 자동취소가 정상적으로 완료 되었습니다.<br />다시 시도하여 주십시오.&Script=APP_PopupHistoryBack();"
								Response.End
						ELSE
								Call SettleErrorLogWrite(LGD_OID, "Y", "PR35", "PayRes", "EShop_Order_Settle 입력 오류 / " & GetPayType(PayType) & " 취소 오류", Err.Description)
	
								Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=주문 처리 도중 오류가 발생하였습니다.<br />결제 자동취소가 정상적으로 처리되지 않았습니다.<br />관리자에게 문의 바랍니다.&Script=APP_PopupHistoryBack();"
								Response.End
						END IF
				END IF
				'-----------------------------------------------------------------------------------------------------------'	
				'결제 정보 저장 End
				'-----------------------------------------------------------------------------------------------------------'	

				'-----------------------------------------------------------------------------------------------------------'	
				'ERP 전송용 I/F 주문 생성 START
				'-----------------------------------------------------------------------------------------------------------'	
				' 가상계좌 결제가 아닐 경우(가상계좌는 입금완료시 처리)
				IF PayType <> "V" THEN

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


														isDBOK = false

														xpay.Rollback("상점 DB처리 실패로 인하여 Rollback 처리 [TID:" & xpay.Response("LGD_TID",0) & ",MID:" & xpay.Response("LGD_MID",0) & ",OID:" & xpay.Response("LGD_OID",0) & "]")
            		
														IF "0000" = xpay.resCode THEN
																Call SettleErrorLogWrite(LGD_OID, "Y", "PR37", "PayRes", "IF_ONLINE_ORDER 입력 오류 / " & GetPayType(PayType) & " 취소 완료", Err.Description)
	
																Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=주문 처리 도중 오류가 발생하였습니다.<br />결제 자동취소가 정상적으로 완료 되었습니다.<br />다시 시도하여 주십시오.&Script=APP_PopupHistoryBack();"
																Response.End
														ELSE
																Call SettleErrorLogWrite(LGD_OID, "Y", "PR38", "PayRes", "IF_ONLINE_ORDER 입력 오류 / " & GetPayType(PayType) & " 취소 오류", Err.Description)
	
																Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=주문 처리 도중 오류가 발생하였습니다.<br />결제 자동취소가 정상적으로 처리되지 않았습니다.<br />관리자에게 문의 바랍니다.&Script=APP_PopupHistoryBack();"
																Response.End
														END IF
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

				'-----------------------------------------------------------------------------------------------------------'	
				'문자발송 시작
				'-----------------------------------------------------------------------------------------------------------'	
				'# Server.Execute("/Common/SMS/OrderSmsSend.asp")
				DIM SmsCode
				IF PayType = "V" THEN
						SmsCode		= "ORD_S100"		'# 입금대기
				ELSE
						SmsCode		= "ORD_S300"		'# 주문완료
				END IF
				SET oCmd = Server.CreateObject("ADODB.Command")
				WITH oCmd
						.ActiveConnection	 = oConn
						.CommandType		 = adCmdStoredProc
						.CommandText		 = "USP_Admin_EShop_Order_Sms_Send"

						.Parameters.Append .CreateParameter("@OrderCode",	 adVarChar,	 adParamInput,   20,	 LGD_OID)
						.Parameters.Append .CreateParameter("@OPIdx",		 adInteger,	 adParamInput,     ,	 0)
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
				Server.Execute("/Common/Mail/OrderMailSend.asp")
				'-----------------------------------------------------------------------------------------------------------'	
				'메일발송 끝
				'-----------------------------------------------------------------------------------------------------------'	



				IF isDBOK THEN
						Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=&Msg=&Script=APP_PopupHistoryBack_Move('/ASP/Order/OrderComplete.asp?OrderCode=" & LGD_OID & "');"
						Response.End
%>
						<script type="text/javascript">
							APP_PopupHistoryBack_Move("/ASP/Order/OrderComplete.asp?OrderCode=<%=LGD_OID%>");
							//location.replace("/ASP/Order/OrderComplete.asp?OrderCode=<%=LGD_OID%>");
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
								Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=주문 처리 도중 오류가 발생하였습니다.<br />결제 자동취소가 정상적으로 완료 되었습니다.<br />다시 시도하여 주십시오.&Script=APP_PopupHistoryBack();"
								Response.End
						ELSE
								Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=주문 처리 도중 오류가 발생하였습니다.<br />결제 자동취소가 정상적으로 처리되지 않았습니다.<br />관리자에게 문의 바랍니다.&Script=APP_PopupHistoryBack();"
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

				Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=결제가 정상적으로 이루어지지 않았습니다.[Code:0001]<br />다시 시도하여 주십시오.&Script=APP_PopupHistoryBack();"
				Response.End
		END IF
ELSE
		'2)API 요청실패 화면처리
		'#Response.Write("결제요청이 실패하였습니다. <br>")
		'#Response.Write("TX Response_code = " & xpay.resCode & "<br>")
		'#Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")
            
		'결제요청 결과 실패 상점 DB처리
		'#Response.Write("결제결제요청 결과 실패 DB처리하시기 바랍니다." & "<br>")
	
		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=결제가 정상적으로 이루어지지 않았습니다.[Code:0002]<br />다시 시도하여 주십시오.&Script=APP_PopupHistoryBack();"
		Response.End
END IF
 

Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>
