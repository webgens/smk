<%@ Language=VBScript codepage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'CouponDownLoad.asp - 쿠폰 다운받기
'Date		: 2018.12.07
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

DIM Idx


DIM PCFlag
DIM CouponType
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
DIM UseFlag
DIM DelFlag
DIM DownCount

DIM StartDT
DIM EndDT

DIM ToDay : ToDay = R_YEAR & R_MONTH & R_DAY & R_HOUR & R_MIN & R_SEC
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'
	
	
Idx				 = sqlFilter(Request("Idx"))
IF Idx = "" THEN
		Response.Write "FAIL|||||받을 쿠폰 정보가 없습니다."
		Response.End
END IF


SET oConn				 = ConnectionOpen()							'# 커넥션 생성
SET oRs					 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성





'# 쿠폰정보
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_Coupon_Select_By_Idx"

		.Parameters.Append .CreateParameter("@Idx",	 adInteger,	 adParamInput, , Idx)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		PCFlag					 = oRs("PCFlag")
		CouponType				 = oRs("CouponType")
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
		UseFlag					 = oRs("UseFlag")
		DelFlag					 = oRs("DelFlag")
		DownCount				 = oRs("DownCount")
ELSE
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||없는 쿠폰 정보 입니다.[01]"
		Response.End
END IF
oRs.Close


'# 삭제된 쿠폰
IF DelFlag = "Y" THEN
		SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||없는 쿠폰 정보 입니다.[02]"
		Response.End
END IF

	
'# 배포가 종료된 쿠폰
IF UseFlag = "N" THEN
		SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||배포 종료된 쿠폰 정보 입니다."
		Response.End
END IF

	
'# 배포기간이 종료된 경우
IF CStr(DistributeEDate) < CStr(ToDay) THEN
		SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||배포 종료된 쿠폰 정보 입니다."
		Response.End
END IF

	
'# 배포기간이 아직 안된 경우
IF CStr(DistributeSDate) > CStr(ToDay) THEN
		SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||없는 쿠폰 정보 입니다.[03]"
		Response.End
END IF


'# 배포타입이 다운로드 쿠폰이 아닐 경우
IF DistributeType <> "D" THEN
		SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||배포가 불가능한 쿠폰 정보 입니다."
		Response.End
END IF



'# 중복 발행 불가 쿠폰일 경우
IF DuplicateFlag = "N" THEN

		'# 받은 쿠폰인지 체크
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Coupon_Member_Select_For_Recevie_Check"
	
				.Parameters.Append .CreateParameter("@MemberNum", adInteger,	 adParamInput, , U_NUM)
				.Parameters.Append .CreateParameter("@CouponIdx", adBigInt,		 adParamInput, , Idx)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		SET oCmd = Nothing

		IF NOT oRs.EOF THEN
				IF oRs("ReceiveFlag") = "N" THEN
						Response.Write "FAIL|||||해당 쿠폰 하나만 받을 수 있는 쿠폰입니다.<br><br>다시 로그인 하시면 마이페이지에서 확인 하실 수 있습니다."
				ELSE
						Response.Write "FAIL|||||해당 쿠폰 하나만 받을 수 있는 쿠폰입니다.<br><br>" & U_NAME & "님은 " & LEFT(oRs("CreateDT"), 10) & "에 쿠폰을 받으셨습니다."
				END IF
				oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.End
		END IF
		oRs.Close

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

						
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Coupon_Member_Insert"
	
		.Parameters.Append .CreateParameter("@MemberNum",	 adInteger,	 adParamInput,   , U_NUM)
		.Parameters.Append .CreateParameter("@CouponIdx",	 adBigInt,	 adParamInput,   , Idx)
		.Parameters.Append .CreateParameter("@StartDT",		 adVarChar,	 adParamInput, 14, StartDT)
		.Parameters.Append .CreateParameter("@EndDT",		 adVarChar,	 adParamInput, 14, EndDT)
		.Parameters.Append .CreateParameter("@CreateID",	 adVarChar,	 adParamInput, 20, U_NUM)
		.Parameters.Append .CreateParameter("@CreateIP",	 adVarChar,	 adParamInput, 15, U_IP)

		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing


	
SET oRs = Nothing
oConn.Close
SET oConn = Nothing



Response.Write "OK|||||"
%>