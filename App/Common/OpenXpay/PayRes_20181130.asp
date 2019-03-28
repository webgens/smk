<%@ Language=VBScript codepage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************/
'PayRes.asp - 카드결제(안심결제) 결과 처리 및 리턴 / 가상계좌 리턴 페이지
'Date		: 2018.11.29
'Update	: 
'/****************************************************************************************/

'//페이지 응답헤더 설정------------------------------------------------------
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'//-------------------------------------------------------------------------------
%>

<!-- #include Virtual = "/ADO/ADODBCommon.asp" -->
<!-- #include Virtual = "/Common/Common.asp" -->
<!-- #include Virtual = "/Common/SetInfo.asp" -->

<%
'If U_ID <> "distance" then
	ON ERROR RESUME NEXT
'End If 

'/*
' * [최종결제요청 페이지(STEP2-2)]
' *
' * LG유플러스으로 부터 내려받은 LGD_PAYKEY(인증Key)를 가지고 최종 결제요청.(파라미터 전달시 POST를 사용하세요)
' */
DIM configPath
configPath = "C:/LGDacom"  'LG유플러스에서 제공한 환경파일("/conf/lgdacom.conf, /conf/mall.conf") 위치 지정.  
'configPath = "C:/lgdacom"  'LG유플러스에서 제공한 환경파일("/conf/lgdacom.conf, /conf/mall.conf") 위치 지정.  

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
xpay.Init configPath, CST_PLATFORM

xpay.Init_TX(LGD_MID)
xpay.Set "LGD_TXNAME", "PaymentByKey"
xpay.Set "LGD_PAYKEY", LGD_PAYKEY


    
DIM oConn						'# ADODB Connection 개체
DIM oRs							'# ADODB Recordset 개체
DIM oCmd						'# ADODB Command 개체


SET oConn	= ConnectionOpen()	'//커넥션 생성
SET oRs		= Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성







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

		'# USAFE_GuaranteeInsurance				 = Request("USAFE_GuaranteeInsurance")
		'# USAFE_GuaranteeInsuranceAgreement		 = Request("USAFE_GuaranteeInsuranceAgreement")
		'# USAFE_JuminNumber						 = Request("USAFE_JuminNumber")
		'# USAFE_EmailFlag							 = Request("USAFE_EmailFlag")
		'# USAFE_SmsFlag							 = Request("USAFE_SmsFlag")

		'# IF USAFE_GuaranteeInsurance = "" THEN USAFE_GuaranteeInsurance = "N"
		'# IF USAFE_GuaranteeInsuranceAgreement = "" THEN USAFE_GuaranteeInsuranceAgreement = "N"
		'# IF USAFE_EmailFlag = "" THEN USAFE_EmailFlag = "N"
		'# IF USAFE_SmsFlag = "" THEN USAFE_SmsFlag = "N"



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
	
				IF NOT oRs.EOF THEN
						U_ID			= oRs("UserID")
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
				IF PayType <> DBPayType THEN
						Set oRs = Nothing
						oConn.Close
						Set oConn = Nothing

						isDBOK = false

						xpay.Rollback("상점 DB처리 실패로 인하여 Rollback 처리 [TID:" & xpay.Response("LGD_TID",0) & ",MID:" & xpay.Response("LGD_MID",0) & ",OID:" & xpay.Response("LGD_OID",0) & "]")
            		
						IF "0000" = xpay.resCode THEN
								Call AlertMessage("주문 처리 도중 오류가 발생하였습니다. 결제 자동취소가 정상적으로 완료 되었습니다. 다시 시도하여 주십시오.", "parent.parent.location.reload();")
								Response.End
						ELSE
								Call AlertMessage("주문 처리 도중 오류가 발생하였습니다. 결제 자동취소가 정상적으로 처리되지 않았습니다. 관리자에게 문의 바랍니다.", "parent.parent.location.reload();")
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
								Call AlertMessage("주문 처리 도중 오류가 발생하였습니다. 결제 자동취소가 정상적으로 완료 되었습니다. 다시 시도하여 주십시오.", "parent.parent.location.reload();")
								Response.End
						ELSE
								Call AlertMessage("주문 처리 도중 오류가 발생하였습니다. 결제 자동취소가 정상적으로 처리되지 않았습니다. 관리자에게 문의 바랍니다.", "parent.parent.location.reload();")
								Response.End
						END IF
				END IF
				'-----------------------------------------------------------------------------------------------------------'	
				'DB에 있는 금액과 PG사에서 넘어온 결재금액이 다르면 취소 END
				'-----------------------------------------------------------------------------------------------------------'


				oConn.BeginTrans	


				'# IF LGD_PAYTYPE = "SC0040" AND LGD_CASFLAG = "R" THEN			'# 가상계좌결제시 계좌할당
						'-----------------------------------------------------------------------------------------------------------'	
						'USafe 보증보험 발급처리 시작
						'-----------------------------------------------------------------------------------------------------------'	
						'-----------------------------------------------------------------------------------------------------------'	
						'USafe 보증보험 발급처리 끝
						'-----------------------------------------------------------------------------------------------------------'	
				'# END IF


				'-----------------------------------------------------------------------------------------------------------'	
				'EShop_Order  업데이트 START
				'-----------------------------------------------------------------------------------------------------------'	
				Set oCmd = Server.CreateObject("ADODB.Command")
				WITH oCmd
						.ActiveConnection = oConn
						.CommandType = adCmdStoredProc
						.CommandText = "USP_Front_EShop_Order_Update_SettleState"
						.Parameters.Append .CreateParameter("@OrderCode",			adVarChar,	adParamInput,	20,		LGD_OID)
						.Parameters.Append .CreateParameter("@OrderState",			adChar,		adParamInput,	 1,		OrderState)
						.Parameters.Append .CreateParameter("@SettleFlag",			adChar,		adParamInput,	 1,		SettleFlag)
						.Parameters.Append .CreateParameter("@SettleDate",			adChar,		adParamInput,	 8,		SettleDate)
						.Parameters.Append .CreateParameter("@SettleTime",			adChar,		adParamInput,	 6,		SettleTime)
						.Parameters.Append .CreateParameter("@ReceiptFlag",			adChar,		adParamInput,	 1,		ReceiptFlag)
						.Parameters.Append .CreateParameter("@ReceiptKind",			adChar,		adParamInput,	 1,		LGD_CASHRECEIPTKIND)
						.Parameters.Append .CreateParameter("@EscrowFlag",			adChar,		adParamInput,	 1,		LGD_ESCROWYN)
						.Parameters.Append .CreateParameter("@CasFlag",				adChar,		adParamInput,	 1,		CasFlag)
						.Parameters.Append .CreateParameter("@PayType",				adChar,		adParamInput,	 1,		PayType)
						.Parameters.Append .CreateParameter("@UpdateID",			adVarChar,	adParamInput,	20,		U_NUM)
						.Parameters.Append .CreateParameter("@UpdateIP",			adVarChar,	adParamInput,	15,		U_IP)
			
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
								Call AlertMessage("주문 처리 도중 오류가 발생하였습니다. 결제 자동취소가 정상적으로 완료 되었습니다. 다시 시도하여 주십시오.", "parent.parent.location.reload();")
								Response.End
						ELSE
								Call AlertMessage("주문 처리 도중 오류가 발생하였습니다. 결제 자동취소가 정상적으로 처리되지 않았습니다. 관리자에게 문의 바랍니다.", "parent.parent.location.reload();")
								Response.End
						END IF
				END IF
				'-----------------------------------------------------------------------------------------------------------'	
				'EShop_Order  업데이트 End
				'-----------------------------------------------------------------------------------------------------------'	

				'-----------------------------------------------------------------------------------------------------------'	
				'구매에 사용한 포인트/슈즈상품권 사용 시작
				'USP_Front_EShop_Order_Update_SettleState 에서 처리한다.
				'-----------------------------------------------------------------------------------------------------------'	
				'# SET oCmd = Server.CreateObject("ADODB.Command")
				'# WITH oCmd
				'# 		.ActiveConnection	 = oConn
				'# 		.CommandType		 = adCmdStoredProc
				'# 		.CommandText		 = "USP_Front_EShop_Order_Product_Select_By_OrderCode"
				'# 
				'# 		.Parameters.Append .CreateParameter("@OrderCode",	adVarChar,	adParamInput, 20,	LGD_OID)
				'# END WITH
				'# oRs.CursorLocation = adUseClient
				'# oRs.Open oCmd, , adOpenStatic, adLockReadOnly
				'# SET oCmd = Nothing
				'# 
				'# IF NOT oRs.EOF THEN
				'# 		Do Until oRs.EOF
				'# 				'-----------------------------------------------------------------------------------------------------------'
				'# 				'포인트 사용처리 시작
				'# 				' PCode : 201(주문사용)
				'# 				'-----------------------------------------------------------------------------------------------------------'
				'# 				IF CDbl(oRs("UsePointPrice")) > 0 THEN
				'# 						DIM RestUsePoint
				'# 						DIM UsePoint
				'# 
				'# 						RestUsePoint	= CDbl(oRs("UsePointPrice"))
				'# 
				'# 						'# 포인트 잔여목록
				'# 						Set oCmd = Server.CreateObject("ADODB.Command")
				'# 						WITH oCmd
				'# 								.ActiveConnection = oConn
				'# 								.CommandType = adCmdStoredProc
				'# 								.CommandText = "USP_Front_EShop_Member_Point_Select_For_Use"
				'# 								.Parameters.Append .CreateParameter("@MemberNum",		 adInteger,	 adParamInput, ,	 U_NUM)
				'# 						END WITH
				'# 						oRs1.CursorLocation = adUseClient
				'# 						oRs1.Open oCmd, , adOpenStatic, adLockReadOnly
				'# 						Set oCmd = Nothing
				'# 
				'# 						IF NOT oRs1.EOF THEN
				'# 								Do Until oRs1.EOF OR RestUsePoint <= 0
				'# 
				'# 										IF CDbl(oRs1("RestPoint")) >= RestUsePoint THEN
				'# 												UsePoint		= RestUsePoint
				'# 										ELSE
				'# 												UsePoint		= CDbl(oRs1("RestPoint"))
				'# 										END IF
				'# 
				'# 										RestUsePoint	= RestUsePoint - UsePoint
				'# 
				'# 										'# 주문 사용포인트 등록처리
				'# 										Set oCmd = Server.CreateObject("ADODB.Command")
				'# 										WITH oCmd
				'# 												.ActiveConnection = oConn
				'# 												.CommandType = adCmdStoredProc
				'# 												.CommandText = "USP_Front_EShop_Member_Point_Use_Insert"
				'# 
				'# 												.Parameters.Append .CreateParameter("@PointIdx",	adInteger,	 adParamInput,    ,		oRs1("Idx"))
				'# 												.Parameters.Append .CreateParameter("@PCode",		adChar,		 adParamInput,   3,		"201")
				'# 												.Parameters.Append .CreateParameter("@UsePoint",	adCurrency,	 adParamInput,    ,		UsePoint)
				'# 												.Parameters.Append .CreateParameter("@OrderCode",	adVarChar,	 adParamInput,  20,		oRs("OrderCode"))
				'# 												.Parameters.Append .CreateParameter("@OPIdx_Org",	adInteger,	 adParamInput,    ,		oRs("OPIdx_Org"))
				'# 												.Parameters.Append .CreateParameter("@CreateID",	adVarChar,	 adParamInput,  20,		U_NUM)
				'# 												.Parameters.Append .CreateParameter("@CreateIP",	adVarChar,	 adParamInput,  15,		U_IP)
				'# 
				'# 												.Execute, , adExecuteNoRecords
				'# 										END WITH
				'# 										Set oCmd = Nothing
				'# 
				'# 										IF Err.Number <> 0 THEN
				'# 												oConn.RollbackTrans
				'# 
				'# 												oRs1.Close : oRs.Close
				'# 												Set oRs1 = Nothing : Set oRs = Nothing
				'# 												oConn.Close : Set oConn = Nothing
				'# 
				'# 												Call AlertMessage("포인트적용 처리 중 오류가 발생하였습니다. [13041]", "")
				'# 												Response.End
				'# 										END IF
				'# 
				'# 										oRs1.MoveNext
				'# 								Loop 
				'# 						ELSE
				'# 								oConn.RollbackTrans
				'# 
				'# 								oRs1.Close : oRs.Close
				'# 								Set oRs1 = Nothing : Set oRs = Nothing
				'# 								oConn.Close : Set oConn = Nothing
				'# 
				'# 								Call AlertMessage("사용할 수 있는 잔여 포인트가 없습니다. [13048]", "")
				'# 								Response.End
				'# 						END IF
				'# 						oRs1.Close
				'# 
				'# 						'# 주문 사용포인트 등록처리
				'# 						Set oCmd = Server.CreateObject("ADODB.Command")
				'# 						WITH oCmd
				'# 								.ActiveConnection = oConn
				'# 								.CommandType = adCmdStoredProc
				'# 								.CommandText = "USP_Front_EShop_Member_Point_Insert"
				'# 
				'# 								.Parameters.Append .CreateParameter("@MemberNum",	adInteger,	 adParamInput,    ,		U_NUM)
				'# 								.Parameters.Append .CreateParameter("@PCode",		adChar,		 adParamInput,   3,		"201")
				'# 								.Parameters.Append .CreateParameter("@AddPoint",	adCurrency,	 adParamInput,    ,		CDbl(oRs("UsePointPrice")) * -1)
				'# 								.Parameters.Append .CreateParameter("@Memo",		adVarChar,	 adParamInput, 300,		"주문시 사용")
				'# 								.Parameters.Append .CreateParameter("@OrderCode",	adVarChar,	 adParamInput,  20,		oRs("OrderCode"))
				'# 								.Parameters.Append .CreateParameter("@OPIdx_Org",	adInteger,	 adParamInput,    ,		oRs("OPIdx_Org"))
				'# 								.Parameters.Append .CreateParameter("@AvailableDT",	adVarChar,	 adParamInput,  10,		Null)
				'# 								.Parameters.Append .CreateParameter("@CreateID",	adVarChar,	 adParamInput,  20,		U_NUM)
				'# 								.Parameters.Append .CreateParameter("@CreateIP",	adVarChar,	 adParamInput,  15,		U_IP)
				'# 
				'# 								.Execute, , adExecuteNoRecords
				'# 						END WITH
				'# 						Set oCmd = Nothing
				'# 
				'# 						IF Err.Number <> 0 THEN
				'# 								oConn.RollbackTrans
				'# 
				'# 								oRs.Close
				'# 								Set oRs1 = Nothing : Set oRs = Nothing
				'# 								oConn.Close : Set oConn = Nothing
				'# 
				'# 								Call AlertMessage("포인트적용 처리 중 오류가 발생하였습니다. [13041]", "")
				'# 								Response.End
				'# 						END IF
				'# 				END IF
				'# 				'-----------------------------------------------------------------------------------------------------------'
				'# 				'포인트 사용처리 끝
				'# 				'-----------------------------------------------------------------------------------------------------------'
				'# 
				'# 				'-----------------------------------------------------------------------------------------------------------'
				'# 				' 슈즈상품권 사용처리 시작
				'# 				' PCode : 211(주문사용)
				'# 				'-----------------------------------------------------------------------------------------------------------'
				'# 				IF CDbl(oRs("UseScashPrice")) > 0 THEN
				'# 						DIM RestUseScash
				'# 						DIM UseScash
				'# 
				'# 						RestUseScash	= CDbl(oRs("UseScashPrice"))
				'# 
				'# 						'# 슈즈상품권 잔여목록
				'# 						Set oCmd = Server.CreateObject("ADODB.Command")
				'# 						WITH oCmd
				'# 								.ActiveConnection = oConn
				'# 								.CommandType = adCmdStoredProc
				'# 								.CommandText = "USP_Front_EShop_Member_SCash_Select_For_Use"
				'# 								.Parameters.Append .CreateParameter("@MemberNum",		 adInteger,	 adParamInput, ,	 U_NUM)
				'# 						END WITH
				'# 						oRs1.CursorLocation = adUseClient
				'# 						oRs1.Open oCmd, , adOpenStatic, adLockReadOnly
				'# 						Set oCmd = Nothing
				'# 
				'# 						IF NOT oRs1.EOF THEN
				'# 								Do Until oRs1.EOF OR RestUseScash <= 0
				'# 
				'# 										IF CDbl(oRs1("RestScash")) >= RestUseScash THEN
				'# 												UseScash		= RestUseScash
				'# 										ELSE
				'# 												UseScash		= CDbl(oRs1("RestScash"))
				'# 										END IF
				'# 
				'# 										RestUseScash	= RestUseScash - UseScash
				'# 
				'# 										'# 주문 사용슈즈상품권 등록처리
				'# 										Set oCmd = Server.CreateObject("ADODB.Command")
				'# 										WITH oCmd
				'# 												.ActiveConnection = oConn
				'# 												.CommandType = adCmdStoredProc
				'# 												.CommandText = "USP_Front_EShop_Member_SCash_Use_Insert"
				'# 
				'# 												.Parameters.Append .CreateParameter("@SCashIdx",	adInteger,	 adParamInput,    ,		oRs1("Idx"))
				'# 												.Parameters.Append .CreateParameter("@SCode",		adChar,		 adParamInput,   3,		"211")
				'# 												.Parameters.Append .CreateParameter("@UseSCash",	adCurrency,	 adParamInput,    ,		UseScash)
				'# 												.Parameters.Append .CreateParameter("@OrderCode",	adVarChar,	 adParamInput,  20,		oRs("OrderCode"))
				'# 												.Parameters.Append .CreateParameter("@OPIdx_Org",	adInteger,	 adParamInput,    ,		oRs("OPIdx_Org"))
				'# 												.Parameters.Append .CreateParameter("@CreateID",	adVarChar,	 adParamInput,  20,		U_NUM)
				'# 												.Parameters.Append .CreateParameter("@CreateIP",	adVarChar,	 adParamInput,  15,		U_IP)
				'# 
				'# 												.Execute, , adExecuteNoRecords
				'# 										END WITH
				'# 										Set oCmd = Nothing
				'# 
				'# 										IF Err.Number <> 0 THEN
				'# 												oConn.RollbackTrans
				'# 
				'# 												oRs1.Close : oRs.Close
				'# 												Set oRs1 = Nothing : Set oRs = Nothing
				'# 												oConn.Close : Set oConn = Nothing
				'# 
				'# 												Call AlertMessage("슈즈상품권적용 처리 중 오류가 발생하였습니다. [13051]", "")
				'# 												Response.End
				'# 										END IF
				'# 
				'# 										oRs1.MoveNext
				'# 								Loop 
				'# 						ELSE
				'# 								oConn.RollbackTrans
				'# 
				'# 								oRs1.Close : oRs.Close
				'# 								Set oRs1 = Nothing : Set oRs = Nothing
				'# 								oConn.Close : Set oConn = Nothing
				'# 
				'# 								Call AlertMessage("사용할 수 있는 잔여 슈즈상품권이 없습니다. [13058]", "")
				'# 								Response.End
				'# 						END IF
				'# 						oRs1.Close
				'# 
				'# 						'# 주문 사용슈즈상품권 등록처리
				'# 						Set oCmd = Server.CreateObject("ADODB.Command")
				'# 						WITH oCmd
				'# 								.ActiveConnection = oConn
				'# 								.CommandType = adCmdStoredProc
				'# 								.CommandText = "USP_Front_EShop_Member_Scash_Insert"
				'# 
				'# 								.Parameters.Append .CreateParameter("@MemberNum",	adInteger,	 adParamInput,    ,		U_NUM)
				'# 								.Parameters.Append .CreateParameter("@SCode",		adChar,		 adParamInput,   3,		"211")
				'# 								.Parameters.Append .CreateParameter("@AddSCash",	adCurrency,	 adParamInput,    ,		CDbl(oRs("UseScashPrice")) * -1)
				'# 								.Parameters.Append .CreateParameter("@Memo",		adVarChar,	 adParamInput, 300,		"주문시 사용")
				'# 								.Parameters.Append .CreateParameter("@OrderCode",	adVarChar,	 adParamInput,  20,		oRs("OrderCode"))
				'# 								.Parameters.Append .CreateParameter("@OPIdx_Org",	adInteger,	 adParamInput,    ,		oRs("OPIdx_Org"))
				'# 								.Parameters.Append .CreateParameter("@CPNo",		adVarChar,	 adParamInput,  20,		"")
				'# 								.Parameters.Append .CreateParameter("@AvailableDT",	adVarChar,	 adParamInput,  10,		Null)
				'# 								.Parameters.Append .CreateParameter("@CreateID",	adVarChar,	 adParamInput,  20,		U_NUM)
				'# 								.Parameters.Append .CreateParameter("@CreateIP",	adVarChar,	 adParamInput,  15,		U_IP)
				'# 
				'# 								.Execute, , adExecuteNoRecords
				'# 						END WITH
				'# 						Set oCmd = Nothing
				'# 
				'# 						IF Err.Number <> 0 THEN
				'# 								oConn.RollbackTrans
				'# 
				'# 								oRs.Close
				'# 								Set oRs1 = Nothing : Set oRs = Nothing
				'# 								oConn.Close : Set oConn = Nothing
				'# 
				'# 								Call AlertMessage("슈즈상품권적용 처리 중 오류가 발생하였습니다. [13051]", "")
				'# 								Response.End
				'# 						END IF
				'# 				END IF
				'# 				'-----------------------------------------------------------------------------------------------------------'
				'# 				' 슈즈상품권 사용처리 끝
				'# 				'-----------------------------------------------------------------------------------------------------------'
				'# 
				'# 
				'# 				oRs.MoveNext
				'# 		Loop 
				'# END IF
				'# oRs.Close
				'-----------------------------------------------------------------------------------------------------------'	
				'구매에 사용한 포인트/슈즈상품권 사용 끝
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
								Call AlertMessage("주문 처리 도중 오류가 발생하였습니다. 카드 자동취소가 정상적으로 완료 되었습니다. 다시 시도하여 주십시오.", "parent.parent.location.reload();")
								Response.End
						ELSE
								Call AlertMessage("주문 처리 도중 오류가 발생하였습니다. 카드 자동취소가 정상적으로 처리되지 않았습니다. 관리자에게 문의 바랍니다.", "parent.parent.location.reload();")
								Response.End
						END IF
				END IF
				'-----------------------------------------------------------------------------------------------------------'	
				'결제 정보 저장 End
				'-----------------------------------------------------------------------------------------------------------'	

				oConn.CommitTrans

				'-----------------------------------------------------------------------------------------------------------'	
				'문자발송 시작
				'-----------------------------------------------------------------------------------------------------------'	
				'Server.Execute("/Common/SMS/SendSMS.asp")
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
							parent.parent.location.href = "/ASP/Order/OrderComplete.asp?OrderCode=<%=LGD_OID%>";
							self.close();
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
								Call AlertMessage("주문 처리 도중 오류가 발생하였습니다. 카드 자동취소가 정상적으로 완료 되었습니다. 다시 시도하여 주십시오.", "parent.parent.location.reload();")
								Response.End
						ELSE
								'#Response.Write("자동취소가 정상적으로 처리되지 않았습니다.<br>")
								Call AlertMessage("주문 처리 도중 오류가 발생하였습니다. 카드 자동취소가 정상적으로 처리되지 않았습니다. 관리자에게 문의 바랍니다.", "parent.parent.location.reload();")
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

				Call AlertMessage("결제가 정상적으로 이루어지지 않았습니다.[Code:0001]\n\n다시 시도하여 주십시오.", "parent.parent.location.reload();")
				Response.End
		END IF
ELSE
		'2)API 요청실패 화면처리
		'#Response.Write("결제요청이 실패하였습니다. <br>")
		'#Response.Write("TX Response_code = " & xpay.resCode & "<br>")
		'#Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")
            
		'결제요청 결과 실패 상점 DB처리
		'#Response.Write("결제결제요청 결과 실패 DB처리하시기 바랍니다." & "<br>")

		Call AlertMessage("결제가 정상적으로 이루어지지 않았습니다.[Code:0002]\n\n다시 시도하여 주십시오.", "parent.parent.location.reload();")
		Response.End
END IF
 

Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>
