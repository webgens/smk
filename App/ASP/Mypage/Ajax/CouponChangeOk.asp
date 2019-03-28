<%@ Language=VBScript codepage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'CouponChangeOk.asp - 쿠폰 전환 배포 처리
'Date		: 2019.01.06
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
<!-- #include Virtual = "/Common/CheckID_Ajax.asp" -->

<%
'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn							'# ADODB Connection 개체
DIM oRs								'# ADODB Recordset 개체
DIM oCmd							'# ADODB Command 개체

DIM wQuery							'# WHERE 절
DIM sQuery							'# SORT 절

DIM i
DIM x
DIM y

DIM CouponNum

DIM CouponEventIdx
DIM B_CouponNum
DIM Seq
DIM Idx
DIM B_RegFlag
DIM B_DelFlag
DIM MemberNum
DIM MemberName
DIM ReceiveIP
DIM ReceiveDT


DIM EventName
DIM CouponIdx
DIM EventSDate
DIM EventEDate
DIM A_CouponType
DIM A_CouponNum
DIM CouponCnt
DIM ExchangeCnt
DIM A_UseFlag
DIM A_DelFlag

DIM PCFlag
DIM C_CouponType
DIM CouponName
DIM DistributeSDate
DIM DistributeEDate
DIM UseDateType
DIM UseSDate
DIM UseEDate
DIM UseDay
DIM DeliveryCouponFlag
DIM ApplyPriceType
DIM Discount
DIM MoneyType
DIM LimitDiscountFlag
DIM LimitDiscount
DIM LimitPriceType
DIM LimitPrice
DIM LimitDistributeFlag
DIM LimitDistributeCount
DIM DistributeType
DIM DuplicateFlag
DIM C_UseFlag
DIM C_DelFlag
DIM DownCount

DIM CEC_Idx
DIM Complete

DIM StartDT
DIM EndDT

DIM CouponMemberIdx

DIM ToDay : ToDay = R_YEAR & R_MONTH & R_DAY & R_HOUR & R_MIN & R_SEC
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'
	
	
CouponNum		 = sqlFilter(Request("CouponNum"))
IF CouponNum = "" THEN
		Response.Write "FAIL|||||쿠폰번호를 입력하여 주십시오."
		Response.End
END IF


SET oConn				 = ConnectionOpen()							'# 커넥션 생성
SET oRs					 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성





'# 바우처 쿠폰번호 정보
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Coupon_Event_CouponNum_Select_By_CouponNum"

		.Parameters.Append .CreateParameter("@CouponNum", adVarChar, adParamInput, 20, CouponNum)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN

		CouponEventIdx			 = oRs("CouponEventIdx")
		B_CouponNum				 = oRs("CouponNum")
		Seq						 = oRs("Seq")
		Idx						 = oRs("Idx")
		B_RegFlag				 = oRs("RegFlag")
		B_DelFlag				 = oRs("DelFlag")
		MemberNum				 = oRs("MemberNum")
		MemberName				 = oRs("MemberName")
		ReceiveIP				 = oRs("ReceiveIP")
		ReceiveDT				 = oRs("ReceiveDT")

ELSE
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||없는 쿠폰 번호 입니다.[01]"
		Response.End
END IF
oRs.Close


'# 바우처 정보
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Coupon_Event_Select_By_Idx"

		.Parameters.Append .CreateParameter("@Idx", adInteger, adParamInput, , CouponEventIdx)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN

		EventName				 = oRs("EventName")
		CouponIdx				 = oRs("CouponIdx")
		EventSDate				 = oRs("EventSDate")
		EventEDate				 = oRs("EventEDate")
		A_CouponType			 = oRs("CouponType")
		A_CouponNum				 = oRs("CouponNum")
		CouponCnt				 = oRs("CouponCnt")
		ExchangeCnt				 = oRs("ExchangeCnt")
		A_UseFlag				 = oRs("UseFlag")
		A_DelFlag				 = oRs("DelFlag")

ELSE
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||없는 쿠폰 입니다.[01]"
		Response.End
END IF
oRs.Close

IF A_UseFlag = "N" THEN
		SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||발급 중지된 쿠폰 입니다."
		Response.End
END IF

IF A_DelFlag = "Y" THEN
		SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||없는 쿠폰 입니다.[02]"
		Response.End
END IF

IF CDate(EventSDate) > Date THEN
		SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||" & EventSDate & " 일부터 교환 가능합니다."
		Response.End
END IF

IF CDate(EventEDate) < Date THEN
		SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||쿠폰 교환 기간이 종료 되었습니다."
		Response.End
END IF





'# 쿠폰 정보
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Coupon_Select_By_Idx"
	
		.Parameters.Append .CreateParameter("@Idx", adInteger, adParamInput, , CouponIdx)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN

		PCFlag					 = oRs("PCFlag")
		C_CouponType			 = oRs("CouponType")
		CouponName				 = oRs("CouponName")
		DistributeSDate			 = oRs("DistributeSDate")
		DistributeEDate			 = oRs("DistributeEDate")
		UseDateType				 = oRs("UseDateType")
		UseSDate				 = oRs("UseSDate")
		UseEDate				 = oRs("UseEDate")
		UseDay					 = oRs("UseDay")
		DeliveryCouponFlag		 = oRs("DeliveryCouponFlag")
		ApplyPriceType			 = oRs("ApplyPriceType")
		Discount				 = oRs("Discount")
		MoneyType				 = oRs("MoneyType")
		LimitDiscountFlag		 = oRs("LimitDiscountFlag")
		LimitDiscount			 = oRs("LimitDiscount")
		LimitPriceType			 = oRs("LimitPriceType")
		LimitPrice				 = oRs("LimitPrice")
		LimitDistributeFlag		 = oRs("LimitDistributeFlag")
		LimitDistributeCount	 = oRs("LimitDistributeCount")
		DistributeType			 = oRs("DistributeType")
		DuplicateFlag			 = oRs("DuplicateFlag")
		C_UseFlag				 = oRs("UseFlag")
		C_DelFlag				 = oRs("DelFlag")
		DownCount				 = oRs("DownCount")

ELSE
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||없는 쿠폰 입니다.[02]"
		Response.End
END IF
oRs.Close







'# 쿠폰번호타입 (1:고정[CoponNum컬럼값] / 2:테이블[EShop_Coupon_Event_CouponNum]쿠번 번호)
IF A_CouponType = "1" THEN

		'# 고정 쿠폰번호가 일치하는지 체크
		IF CouponNum <> A_CouponNum THEN
				SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "FAIL|||||일치하는 쿠폰번호가 없습니다."
				Response.End
		END IF

		'# 받은 쿠폰인지 체크
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Coupon_Event_CouponNum_Select_For_Receive"

				.Parameters.Append .CreateParameter("@CouponEventIdx",	 adInteger, adParamInput,   , CouponEventIdx)
				.Parameters.Append .CreateParameter("@CouponNum",		 adVarChar, adParamInput, 20, CouponNum)
				.Parameters.Append .CreateParameter("@MemberNum",		 adInteger, adParamInput,   , U_NUM)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		SET oCmd = Nothing

		IF NOT oRs.EOF THEN
				oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "FAIL|||||회원님은 이미 교환 받으셨습니다."
				Response.End
		END IF
		oRs.Close


		'# 잔여 수량 체크
		IF CDbl(CouponCnt) - CDbl(ExchangeCnt) < 1 THEN
				SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "FAIL|||||쿠폰이 모두 소진되었습니다."
				Response.End
		END IF



		'# 교환 안된 바우처 쿠폰 번호
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Coupon_Event_CouponNum_Select_For_Not_Receive_Top1"
	
				.Parameters.Append .CreateParameter("@CouponEventIdx",	 adInteger, adParamInput,   , CouponEventIdx)
				.Parameters.Append .CreateParameter("@CouponNum",		 adVarChar, adParamInput, 20, CouponNum)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		SET oCmd = Nothing

		IF NOT oRs.EOF THEN
				CEC_Idx		 = oRs("Idx")
		ELSE
				oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "FAIL|||||쿠폰이 모두 소진되었습니다."
				Response.End
		END IF
		oRs.Close



		oConn.BeginTrans




		'# 일단 바우처 쿠폰 선점 업데이트
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Coupon_Event_CouponNum_Update_For_Receive"
	
				.Parameters.Append .CreateParameter("@Idx",			 adInteger, adParamInput,   , CEC_Idx)
				.Parameters.Append .CreateParameter("@MemberNum",	 adInteger, adParamInput, 20, U_NUM)
				.Parameters.Append .CreateParameter("@MemberName",	 adVarChar, adParamInput, 50, U_NAME)
				.Parameters.Append .CreateParameter("@UpdateID",	 adVarChar, adParamInput, 20, U_NUM)
				.Parameters.Append .CreateParameter("@UpdateIP",	 adVarChar, adParamInput, 15, U_IP)
				.Parameters.Append .CreateParameter("@Complete",	 adInteger, adParamoutput)

				.Execute, , adExecuteNoRecords
				Complete = .Parameters("@Complete").Value
		END WITH
		SET oCmd = Nothing


		IF Complete = 1 THEN
				SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "FAIL|||||다시 시도하여 주십시오."
				Response.End
		END IF


		IF UseDateType = "U" THEN
				StartDT	 = U_DATE & R_HOUR & "0000"
				EndDT	 = "99999999999999"
		ELSEIF UseDateType = "P" THEN
				StartDT	 = U_DATE & R_HOUR & "0000"
				EndDT	 = UseEDate
		ELSE
				StartDT	 = U_DATE & "000000"
				EndDT	 = REPLACE(DATEADD("d", UseDay, Date), "-", "") & "240000"
		END IF

				
		'# 쿠폰발급
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Coupon_Member_Insert_For_Return_Idx"
	
				.Parameters.Append .CreateParameter("@MemberNum",	 adInteger,	 adParamInput,   , U_NUM)
				.Parameters.Append .CreateParameter("@CouponIdx",	 adBigInt,	 adParamInput,   , CouponIdx)
				.Parameters.Append .CreateParameter("@StartDT",		 adVarChar,	 adParamInput, 14, StartDT)
				.Parameters.Append .CreateParameter("@EndDT",		 adVarChar,	 adParamInput, 14, EndDT)
				.Parameters.Append .CreateParameter("@CreateID",	 adVarChar,	 adParamInput, 20, U_NUM)
				.Parameters.Append .CreateParameter("@CreateIP",	 adVarChar,	 adParamInput, 15, U_IP)
				.Parameters.Append .CreateParameter("@Idx",			 adInteger, adParamoutput)

				.Execute, , adExecuteNoRecords
				CouponMemberIdx = .Parameters("@Idx").Value
		END WITH
		SET oCmd = Nothing

	
		'# 일단 바우처 쿠폰에 발급 쿠폰 일련번호 업데이트
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Coupon_Event_CouponNum_Update_For_CouponMemberIdx"
	
				.Parameters.Append .CreateParameter("@Idx",				 adInteger, adParamInput,   , CEC_Idx)
				.Parameters.Append .CreateParameter("@CouponMemberIdx",	 adInteger, adParamInput,   , CouponMemberIdx)

				.Execute, , adExecuteNoRecords
		END WITH
		SET oCmd = Nothing


		oConn.CommitTrans


'# 2 : 쿠폰번호타입 (1:고정[CoponNum컬럼값] / 2:테이블[EShop_Coupon_Event_CouponNum]쿠번 번호)
ELSE
		

		IF B_DelFlag = "Y" THEN
				SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "FAIL|||||없는 쿠폰 번호 입니다.[02]"
				Response.End
		END IF

		IF B_RegFlag = "Y" THEN
				SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "FAIL|||||사용된 쿠폰 번호 입니다."
				Response.End
		END IF


	
		oConn.BeginTrans




		'# 일단 바우처 쿠폰 선점 업데이트
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Coupon_Event_CouponNum_Update_For_Receive"
	
				.Parameters.Append .CreateParameter("@Idx",			 adInteger, adParamInput,   , Idx)
				.Parameters.Append .CreateParameter("@MemberNum",	 adInteger, adParamInput, 20, U_NUM)
				.Parameters.Append .CreateParameter("@MemberName",	 adVarChar, adParamInput, 50, U_NAME)
				.Parameters.Append .CreateParameter("@UpdateID",	 adVarChar, adParamInput, 20, U_NUM)
				.Parameters.Append .CreateParameter("@UpdateIP",	 adVarChar, adParamInput, 15, U_IP)
				.Parameters.Append .CreateParameter("@Complete",	 adInteger, adParamoutput)

				.Execute, , adExecuteNoRecords
				Complete = .Parameters("@Complete").Value
		END WITH
		SET oCmd = Nothing


		IF Complete = 1 THEN
				SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "FAIL|||||다시 시도하여 주십시오."
				Response.End
		END IF


		IF UseDateType = "U" THEN
				StartDT	 = U_DATE & R_HOUR & "0000"
				EndDT	 = "99999999999999"
		ELSEIF UseDateType = "P" THEN
				StartDT	 = U_DATE & R_HOUR & "0000"
				EndDT	 = UseEDate
		ELSE
				StartDT	 = U_DATE & "000000"
				EndDT	 = REPLACE(DATEADD("d", UseDay, Date), "-", "") & "240000"
		END IF

				
		'# 쿠폰발급
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Coupon_Member_Insert_For_Return_Idx"
	
				.Parameters.Append .CreateParameter("@MemberNum",	 adInteger,	 adParamInput,   , U_NUM)
				.Parameters.Append .CreateParameter("@CouponIdx",	 adBigInt,	 adParamInput,   , CouponIdx)
				.Parameters.Append .CreateParameter("@StartDT",		 adVarChar,	 adParamInput, 14, StartDT)
				.Parameters.Append .CreateParameter("@EndDT",		 adVarChar,	 adParamInput, 14, EndDT)
				.Parameters.Append .CreateParameter("@CreateID",	 adVarChar,	 adParamInput, 20, U_NUM)
				.Parameters.Append .CreateParameter("@CreateIP",	 adVarChar,	 adParamInput, 15, U_IP)
				.Parameters.Append .CreateParameter("@Idx",			 adInteger, adParamoutput)

				.Execute, , adExecuteNoRecords
				CouponMemberIdx = .Parameters("@Idx").Value
		END WITH
		SET oCmd = Nothing

	
		'# 일단 바우처 쿠폰에 발급 쿠폰 일련번호 업데이트
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Coupon_Event_CouponNum_Update_For_CouponMemberIdx"
	
				.Parameters.Append .CreateParameter("@Idx",				 adInteger, adParamInput,   , Idx)
				.Parameters.Append .CreateParameter("@CouponMemberIdx",	 adInteger, adParamInput,   , CouponMemberIdx)

				.Execute, , adExecuteNoRecords
		END WITH
		SET oCmd = Nothing


		oConn.CommitTrans

END IF




	
SET oRs = Nothing
oConn.Close
SET oConn = Nothing



Response.Write "OK|||||"
%>