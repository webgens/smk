<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'NonDepositOrderCancelOk.asp - 입금전 주문취소 처리
'Date		: 2019.01.02
'Update	: 
'/****************************************************************************************'

'//페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'//---------------------------------------------------------------------------------------'
%>

<!-- #include Virtual = "/ADO/ADODBCommon.asp" -->
<!-- #include Virtual = "/Common/Common.asp" -->
<!-- #include Virtual = "/Common/SetInfo.asp" -->

<%
IF U_NUM = "" AND N_NAME = "" THEN
		Response.Write "LOGIN|||||로그인 정보가 없습니다. 다시 로그인하여 주십시오."
		Response.End
END IF

'/****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn											'# ADODB Connection 개체
DIM oRs												'# ADODB Recordset 개체
DIM oRs1											'# ADODB Recordset 개체
DIM oCmd											'# ADODB Command 개체

DIM i
DIM j
DiM x
DIM y

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절

DIM OrderCode

DIM PayType
DIM SettlePrice

DIM GuaranteeInsurance
DIM GuaranteeInsuranceGubun

DIM LGD_MID
DIM LGD_TID
DIM LGD_CANCELREASON
DIM LGD_CANCELREQUESTER
DIM LGD_CANCELREQUESTERIP

DIM LGD_RESPCODE
DIM LGD_RESPMSG

DIM Result
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


OrderCode		= sqlFilter(Request("OrderCode"))


IF OrderCode = "" THEN
		Response.Write "FAIL|||||취소할 입력정보가 부족합니다."
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


SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성



'# 주문정보 체크
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Order_Select_For_OrderInfo"

		.Parameters.Append .CreateParameter("@OrderCode",	adVarchar, adParaminput,	20,		OrderCode)
		.Parameters.Append .CreateParameter("@UserID",		adVarchar, adParaminput,	20,		U_NUM)
		.Parameters.Append .CreateParameter("@OrderName",	adVarChar, adParamInput,	50,		N_NAME)
		.Parameters.Append .CreateParameter("@OrderHp",		adVarChar, adParamInput,	20,		N_HP)
		.Parameters.Append .CreateParameter("@OrderEmail",	adVarChar, adParamInput,	50,		N_EMAIL)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		'# 미입금 상태가 아니면 취소불가
		IF oRs("SettleFlag") = "Y" THEN
				oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "FAIL|||||결제완료된 주문으로 주문취소할 수 없습니다."
				Response.End

		'# 가상계좌 발급 상태가 아니면 취소불가
		ELSEIF oRs("CasFlag") <> "R" THEN
				oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "FAIL|||||미결제 주문만 주문취소가 가능합니다."
				Response.End
		END IF

		LGD_TID		= oRs("LGD_TID")
		PayType		= oRs("PayType")
		SettlePrice	= oRs("OrderPrice") + oRs("DeliveryPrice")

		GuaranteeInsurance		 = oRs("GuaranteeInsurance")
		GuaranteeInsuranceGubun	 = oRs("GuaranteeInsuranceGubun")
ELSE
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||취소할 주문내역이 없습니다."
		Response.End
END IF
oRs.Close


'# 주문상품 상태 체크
wQuery = ""
wQuery = wQuery & "WHERE A.IsShowFlag = 'Y' AND A.ProductType IN ('P','O') "
wQuery = wQuery & "AND A.OrderCode = '" & OrderCode & "' "

sQuery = "ORDER BY A.OPIdx_Group, A.OPIdx_Org"

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
		Do Until oRs.EOF
				'# 미입금 상태가 아니면 취소불가
				IF oRs("OrderState") <> "1" THEN
						oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
						Response.Write "FAIL|||||주문취소할 수 없는 상태의 상품 있습니다."
						Response.End
				END IF

				oRs.MoveNext
		Loop
ELSE
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||취소할 주문내역이 없습니다."
		Response.End
END IF
oRs.Close



oConn.BeginTrans


'-----------------------------------------------------------------------------------------------------------'	
'입금전 주문취소  업데이트 START
'2. 주문 상품 정보 테이블에 주문 상태 정보 Update
'3. 가상계좌 발급 반납일 경우 처리
'	3-1. 쿠폰 환원 처리 Upudate
'	3-2. 포인트, 슈즈상품권 환원 처리
'		3-2-1. 포인트 환원 처리
'			3-2-1-1. 포인트 사용 삭제 처리
'			3-2-1-2. 포인트 사용이력 삭제 처리 시작
'				3-2-1-2-1. 회원포인트 사용이력 삭제
'				3-2-1-2-2. 회원포인트 사용차감
'			3-2-1-3. 회원정보 포인트 누적처리
'		3-2-2. 슈즈상품권 환원 처리 시작
'			3-2-2-1. 슈즈상품권 사용 삭제 처리
'			3-2-2-2. 슈즈상품권 사용이력 삭제 처리
'				3-2-2-2-1. 회원슈즈상품권 사용이력 삭제
'				3-2-2-2-2. 회원슈즈상품권 사용차감
'			3-2-2-3. 회원정보 슈즈상품권 누적처리
'	3-3. 임직원쿠폰 사용 처리 Upudate
'	3-4. 주문 상품 재고 Upudate
'-----------------------------------------------------------------------------------------------------------'	
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Front_EShop_Order_Update_For_NonDeposit_OrderCancel"
		.Parameters.Append .CreateParameter("@OrderCode",			adVarChar,	adParamInput,	 20,		OrderCode)
		.Parameters.Append .CreateParameter("@UpdateNM",			adVarChar,	adParamInput,	100,		U_NAME)
		.Parameters.Append .CreateParameter("@UpdateID",			adVarChar,	adParamInput,	 20,		U_NUM)
		.Parameters.Append .CreateParameter("@UpdateIP",			adVarChar,	adParamInput,	 15,		U_IP)
			
		.Execute, , adExecuteNoRecords
END WITH
Set oCmd = Nothing

IF Err.number <> 0 THEN
		oConn.RollbackTrans

		SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||주문취소 중 오류가 발생하였습니다.[1]"
		Response.End
END IF
'-----------------------------------------------------------------------------------------------------------'	
'EShop_Order  업데이트 End
'-----------------------------------------------------------------------------------------------------------'	




'ON ERROR RESUME NEXT
'-----------------------------------------------------------------------------------------------------------'	
'가상계좌 반납 Start
'-----------------------------------------------------------------------------------------------------------'	
IF IsNull(LGD_TID) = false AND LGD_TID <> "" THEN

        '/*
        ' * [결제취소 요청 페이지]
        ' *
        ' * LG유플러스으로 부터 내려받은 거래번호(LGD_TID)를 가지고 취소 요청을 합니다.(파라미터 전달시 POST를 사용하세요)
        ' * (승인시 LG유플러스으로 부터 내려받은 PAYKEY와 혼동하지 마세요.)
        ' */

        'CST_PLATFORM         = trim(request("CST_PLATFORM"))        ' LG유플러스 결제서비스 선택(test:테스트, service:서비스)
        'CST_MID              = trim(request("CST_MID"))             ' LG유플러스으로 부터 발급받으신 상점아이디를 입력하세요.
                                                                    ' 테스트 아이디는 't'를 제외하고 입력하세요.
        IF PAY_PLATFORM = "test" THEN                               ' 상점아이디(자동생성)
				LGD_MID = "t" & CST_MID
        ELSE
				LGD_MID = CST_MID
        END IF
        '#LGD_TID               = trim(request("LGD_TID"))          ' LG유플러스으로 부터 내려받은 거래번호(LGD_TID)
        LGD_CANCELREASON        = "주문취소"                        ' 취소사유
        IF N_NAME = "" THEN
				LGD_CANCELREQUESTER     = U_NAME                            ' 취소요청자
        ELSE
				LGD_CANCELREQUESTER     = N_NAME                            ' 취소요청자
        END IF
        LGD_CANCELREQUESTERIP   = U_IP                              ' 취소요청IP
    
	    ' ※ 중요
	    ' 환경설정 파일의 경우 반드시 외부에서 접근이 가능한 경로에 두시면 안됩니다.
	    ' 해당 환경파일이 외부에 노출이 되는 경우 해킹의 위험이 존재하므로 반드시 외부에서 접근이 불가능한 경로에 두시기 바랍니다. 
	    ' 예) [Window 계열] C:\inetpub\wwwroot\lgdacom -- 절대불가(웹 디렉토리)
        DIM configPath
        configPath = "C:/LGDacom"  'LG유플러스에서 제공한 환경파일("/conf/lgdacom.conf, /conf/mall.conf") 위치 지정.  

        DIM xpay				' 결제요청 API 객체

        Set xpay = CreateObject("XPayClientCOM.XPayClient")
        xpay.Init configPath, PAY_PLATFORM
        xpay.Init_TX(LGD_MID)

        xpay.Set "LGD_TXNAME",				"Settlement"
        xpay.Set "LGD_TID",					LGD_TID
        xpay.Set "LGD_CANCELREASON",		LGD_CANCELREASON
        xpay.Set "LGD_CANCELREQUESTER",		LGD_CANCELREQUESTER
        xpay.Set "LGD_CANCELREQUESTERIP",	LGD_CANCELREQUESTERIP
 

        '/*
        ' * 1. 결제취소 요청 결과처리
        ' *
        ' * 취소결과 리턴 파라미터는 연동메뉴얼을 참고하시기 바랍니다.
	    ' *
	    ' * [[[중요]]] 고객사에서 정상취소 처리해야할 응답코드
	    ' * 1. 신용카드 : 0000, AV11  
	    ' * 2. 계좌이체 : 0000, RF00, RF10, RF09, RF15, RF19, RF23, RF25 (환불진행중 응답-> 환불결과코드.xls 참고)
	    ' * 3. 나머지 결제수단의 경우 0000(성공) 만 취소성공 처리
	    ' *
        ' */

        IF xpay.TX() THEN
				'1)결제취소결과 화면처리(성공,실패 결과 처리를 하시기 바랍니다.)
				'Response.Write("결제취소 요청이 완료되었습니다. <br>")
				'Response.Write("TX Response_code = " & xpay.resCode & "<br>")
				'Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")

				LGD_RESPCODE				 = xpay.Response("LGD_RESPCODE", 0)
				LGD_RESPMSG					 = xpay.Response("LGD_RESPMSG", 0)

				IF LGD_RESPCODE = "0000" THEN
						'-----------------------------------------------------------------------------------------------------------'	
						'결제 정보 저장 START
						'-----------------------------------------------------------------------------------------------------------'
						Set oCmd = Server.CreateObject("ADODB.Command")
						WITH oCmd
								.ActiveConnection = oConn
								.CommandType = adCmdStoredProc
								.CommandText = "USP_Front_EShop_Order_Settle_Cancel_Insert"
								.Parameters.Append .CreateParameter("@OrderCode",					adVarChar,	adParamInput,	 20,	OrderCode)
								.Parameters.Append .CreateParameter("@LGD_RESPCODE",				adVarChar,	adParamInput,	  4,	LGD_RESPCODE)
								.Parameters.Append .CreateParameter("@LGD_RESPMSG",					adVarChar,	adParamInput,	512,	LGD_RESPMSG)
								.Parameters.Append .CreateParameter("@LGD_AMOUNT",					adVarChar,	adParamInput,	 12,	SettlePrice)
								.Parameters.Append .CreateParameter("@LGD_MID",						adVarChar,	adParamInput,	 15,	LGD_MID)
								.Parameters.Append .CreateParameter("@LGD_TID",						adVarChar,	adParamInput,	 24,	LGD_TID)
								.Parameters.Append .CreateParameter("@LGD_OID",						adVarChar,	adParamInput,	 64,	OrderCode)
								.Parameters.Append .CreateParameter("@LGD_TIMESTAMP",				adVarChar,	adParamInput,	 14,	U_DATE & U_TIME)
								.Parameters.Append .CreateParameter("@LGD_PAYTYPE",					adVarChar,	adParamInput,	  6,	"SC0040")
								.Parameters.Append .CreateParameter("@LGD_RFBANKCODE",				adVarChar,	adParamInput,	  2,	"")
								.Parameters.Append .CreateParameter("@LGD_RFACCOUNTNUM",			adVarChar,	adParamInput,	 20,	"")
								.Parameters.Append .CreateParameter("@LGD_RFCUSTOMERNAME",			adVarChar,	adParamInput,	 40,	"")
								.Parameters.Append .CreateParameter("@LGD_RFPHONE",					adVarChar,	adParamInput,	 20,	"")
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

								xpay.Rollback("상점 DB처리 실패로 인하여 Rollback 처리 [TID:" & LGD_TID & ",MID:" & LGD_MID & ",OID:" & OrderCode & "]")
            		
								IF "0000" = xpay.resCode THEN
										Call SettleErrorLogWrite(LGD_OID, "Y", "PR11", "OrderCancelOk", "EShop_Order_Settle_Cancel 입력 오류 / " & GetPayType(PayType) & " 취소 완료", Err.Description)
										Response.Write "FAIL|||||주문취소 처리 도중 오류가 발생하였습니다.[11]"
										Response.End
								ELSE
										Call SettleErrorLogWrite(LGD_OID, "Y", "PR12", "OrderCancelOk", "EShop_Order_Settle_Cancel 입력 오류 / " & GetPayType(PayType) & " 취소 오류", Err.Description)
										Response.Write "FAIL|||||주문취소 처리 도중 오류가 발생하였습니다.[12]"
										Response.End
								END IF
						END IF
						'-----------------------------------------------------------------------------------------------------------'	
						'결제 정보 저장 End
						'-----------------------------------------------------------------------------------------------------------'	
				ELSE
						'-----------------------------------------------------------------------------------------------------------'	
						'결제 정보 저장 START
						'-----------------------------------------------------------------------------------------------------------'
						Set oCmd = Server.CreateObject("ADODB.Command")
						WITH oCmd
								.ActiveConnection = oConn
								.CommandType = adCmdStoredProc
								.CommandText = "USP_Front_EShop_Order_Settle_Cancel_Insert"
								.Parameters.Append .CreateParameter("@OrderCode",					adVarChar,	adParamInput,	 20,	OrderCode)
								.Parameters.Append .CreateParameter("@LGD_RESPCODE",				adVarChar,	adParamInput,	  4,	LGD_RESPCODE)
								.Parameters.Append .CreateParameter("@LGD_RESPMSG",					adVarChar,	adParamInput,	512,	LGD_RESPMSG)
								.Parameters.Append .CreateParameter("@LGD_AMOUNT",					adVarChar,	adParamInput,	 12,	SettlePrice)
								.Parameters.Append .CreateParameter("@LGD_MID",						adVarChar,	adParamInput,	 15,	LGD_MID)
								.Parameters.Append .CreateParameter("@LGD_TID",						adVarChar,	adParamInput,	 24,	LGD_TID)
								.Parameters.Append .CreateParameter("@LGD_OID",						adVarChar,	adParamInput,	 64,	OrderCode)
								.Parameters.Append .CreateParameter("@LGD_TIMESTAMP",				adVarChar,	adParamInput,	 14,	U_DATE & U_TIME)
								.Parameters.Append .CreateParameter("@LGD_PAYTYPE",					adVarChar,	adParamInput,	  6,	"SC0040")
								.Parameters.Append .CreateParameter("@LGD_RFBANKCODE",				adVarChar,	adParamInput,	  2,	"")
								.Parameters.Append .CreateParameter("@LGD_RFACCOUNTNUM",			adVarChar,	adParamInput,	 20,	"")
								.Parameters.Append .CreateParameter("@LGD_RFCUSTOMERNAME",			adVarChar,	adParamInput,	 40,	"")
								.Parameters.Append .CreateParameter("@LGD_RFPHONE",					adVarChar,	adParamInput,	 20,	"")
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

								Response.Write "FAIL|||||주문취소 처리 도중 오류가 발생하였습니다.[41]"
								Response.End
						END IF
						'-----------------------------------------------------------------------------------------------------------'	
						'결제 정보 저장 End
						'-----------------------------------------------------------------------------------------------------------'	

				END IF

        ELSE
				'2)API 요청 실패 화면처리
				'Response.Write("결제취소 요청이 실패하였습니다. <br>")
				'Response.Write("TX Response_code = " & xpay.resCode & "<br>")
				'Response.Write("TX Response_msg = " & xpay.resMsg & "<p>")

				oConn.RollbackTrans

				Set oRs = Nothing
				oConn.Close
				Set oConn = Nothing

				Response.Write "FAIL|||||주문취소 처리 도중 오류가 발생하였습니다.[01]"
				Response.End
        END IF

END IF
'-----------------------------------------------------------------------------------------------------------'	
'가상계좌 반납 End
'-----------------------------------------------------------------------------------------------------------'	

oConn.CommitTrans

'-----------------------------------------------------------------------------------------------------------'	
'# 보증보험 취소 Start
'-----------------------------------------------------------------------------------------------------------'	
IF GuaranteeInsurance = "Y" AND GuaranteeInsuranceGubun = "A0" THEN

		DIM USafeCom
		DIM USafeComResult
		DIM UsafeComResultCode
		DIM UsafeComResultMsg
		Set USafeCom		= CreateObject( "USafeCom.guarantee.1"  )

		USafeCom.Port		= 80
		USafeCom.Url		= "gateway.usafe.co.kr"
		USafeCom.CallForm	= "/esafe/guartrn.asp"

		'데이터 64Bit 암호화시 사용
		USafeCom.EncKey		= "uclick"

	
		'///////////////////////////////////////////////////////////////////////////
		USafeCom.gubun 		= "B0"	                         
		USafeCom.mallId		= USAFE_ID
		USafeCom.oId		= OrderCode	' 상점의 주문번호
		'// 테스트를 위해 코딩 end
		'///////////////////////////////////////////////////////////////////////////

		USafeComResult		= USafeCom.cancelInsurance
		UsafeComResultCode	= Left( USafeComResult , 1 )
		UsafeComResultMsg	= Mid( USafeComResult , 3 )


		'# 보증보험 로그 생성
		Set oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection = oConn
				.CommandType = adCmdStoredProc
				.CommandText = "USP_Admin_EShop_Usafe_Log_Insert"
				.Parameters.Append .CreateParameter("@OrderCode",					adVarChar,		adParamInput,	  20,		OrderCode)
				.Parameters.Append .CreateParameter("@UsafeGubun",					adChar,			adParamInput,	   2,		"B0")
				.Parameters.Append .CreateParameter("@UsafeResultCode",				adVarChar,		adParamInput,	  50,		UsafeComResultCode)
				.Parameters.Append .CreateParameter("@UsafeResultMsg",				adVarChar,		adParamInput,	1000,		Replace(USafeComResult, "'", ""))
				.Parameters.Append .CreateParameter("@U_MEMNUM",					adVarChar,		adParamInput,	  50,		U_NUM)
				.Parameters.Append .CreateParameter("@U_IP",						adVarChar,		adParamInput,	  15,		U_IP)

				.Execute, , adExecuteNoRecords
		END WITH
		Set oCmd = Nothing
END IF
'-----------------------------------------------------------------------------------------------------------'	
'# 보증보험 취소 End
'-----------------------------------------------------------------------------------------------------------'	



Response.Write "OK|||||"


Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>