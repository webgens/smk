<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'OrderSheetUseCouponCheck.asp - 주문서 쿠폰 사용 체크처리 페이지
'Date		: 2018.12.28
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
<!-- #include Virtual = "/Common/CheckID_Ajax.asp" -->

<%
'/****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn											'# ADODB Connection 개체
DIM oRs												'# ADODB Recordset 개체
DIM oCmd											'# ADODB Command 개체

DIM i
DIM j
DiM x
DIM y

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절

DIM OrderSheetIdx
DIM MemberCouponIdx
DIM ApplyFlag

DIM TagPrice
DIM SalePrice
DIM DCRate
DIM UseCouponPrice
DIM UseScashPrice
DIM UsePointPrice
DIM DiscountPrice
DIM OrderPrice

DIM CouponIdx
DIM CouponName
DIM CouponType
DIM LimitPriceType
DIM LimitPrice
DIM MoneyType
DIM Discount
DIM ApplyPriceType
DIM LimitDiscountFlag
DIM LimitDiscount
DIM DuplicateUseFlag

DIM TotalDiscountPrice	: TotalDiscountPrice	= 0
DIM CouponIdxs			: CouponIdxs			= ""
DIM DuplicateUseFlags	: DuplicateUseFlags		= ""
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


OrderSheetIdx		= sqlFilter(Request("OrderSheetIdx"))
MemberCouponIdx		= sqlFilter(Request("MemberCouponIdx"))
ApplyFlag			= sqlFilter(Request("ApplyFlag"))


IF OrderSheetIdx = "" THEN
		Response.Write "FAIL|||||입력정보가 부족합니다."
		Response.End
END IF




SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성



'# 1. 주문서  구매 상품 정보 체크
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_OrderSheet_Select_By_Idx"

		.Parameters.Append .CreateParameter("@CartID",	adVarChar,	adParamInput, 20,		U_CARTID)
		.Parameters.Append .CreateParameter("@Idx",		adInteger,	adParamInput,   ,		OrderSheetIdx)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		IF oRs("SalePriceType") = "2" THEN
				oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "FAIL|||||임직원 판매가 구매시 쿠폰을 사용하실 수 없습니다."
				Response.End
		END IF

		TagPrice			= oRs("TagPrice")
		IF oRs("SalePriceType") = "2" THEN
				SalePrice			= oRs("EmployeeSalePrice")
				DCRate				= oRs("EmployeeDCRate")
		ELSE
				SalePrice			= oRs("SalePrice")
				DCRate				= oRs("DCRate")
		END IF
		UseCouponPrice		= oRs("UseCouponPrice")
		UsePointPrice		= oRs("UsePointPrice")
		UseScashPrice		= oRs("UseScashPrice")
ELSE
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||선택한 주문상품이 주문서에 없습니다."
		Response.End
END IF
oRs.Close


'# 2. 주문금액을 판매가로 먼저 셋팅
OrderPrice = CDbl(SalePrice)
'# 포인트사용금액과 슈즈상품권사용금액을 빼준다
OrderPrice = OrderPrice - CDbl(UsePointPrice) - CDbl(UseScashPrice)



oConn.BeginTrans



'# 3. 사용 쿠폰 정보 임시 테이블 초기화(삭제)
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Front_EShop_OrderSheet_UseCoupon_Temp_Delete"

		.Parameters.Append .CreateParameter("@MemberNum",		 adInteger,	 adParamInput, ,	 U_NUM)
		.Parameters.Append .CreateParameter("@OrderSheetIdx",	 adInteger,	 adParamInput, ,	 OrderSheetIdx)

		.Execute, , adExecuteNoRecords
END WITH
Set oCmd = Nothing

IF Err.Number <> 0 THEN
		oConn.RollbackTrans

		SET oRs = Nothing
		oConn.Close
		SET oConn = Nothing

		Response.Write "FAIL|||||쿠폰적용 처리 중 오류가 발생하였습니다.[1]"
		Response.End
END IF



IF MemberCouponIdx <> "" THEN
		MemberCouponIdx		= Split(MemberCouponIdx, ",")

		'# 4. 회원에 배포되어 있고 이상품에 사용 가능한 쿠폰인지 체크
		FOR i = 0 TO UBound(MemberCouponIdx)

		'Response.Write "MemberCouponIdx(" & i & ") = " & MemberCouponIdx(i) & "<br>"

				'# 4-1. 사용가능한 쿠폰인지 체크
				Set oCmd = Server.CreateObject("ADODB.Command")
				WITH oCmd
						.ActiveConnection = oConn
						.CommandType = adCmdStoredProc
						.CommandText = "USP_Front_EShop_Coupon_Member_Select_By_Idx"
						.Parameters.Append .CreateParameter("@MemberNum",			adInteger,	 adParamInput,   ,	U_NUM)
						.Parameters.Append .CreateParameter("@MemberCouponIdx",		adInteger,	 adParamInput,   ,	MemberCouponIdx(i))
				END WITH
				oRs.Open oCmd, , adOpenStatic, adLockReadOnly
				Set oCmd = Nothing

				IF NOT oRs.EOF THEN
						CouponName		= oRs("CouponName")

						IF oRs("ReceiveFlag") <> "Y" THEN
								oConn.RollbackTrans

								oRs.Close
								Set oRs = Nothing
								oConn.Close
								Set oConn = Nothing

								Response.Write "FAIL|||||[" & CouponName & "] 쿠폰은 배포되지 않은 쿠폰입니다. 사용하실 수 없습니다."
								Response.End

						ELSEIF oRs("UseFlag") = "Y" THEN
								oConn.RollbackTrans

								oRs.Close
								Set oRs = Nothing
								oConn.Close
								Set oConn = Nothing

								Response.Write "FAIL|||||[" & CouponName & "] 쿠폰은 이미 사용된 쿠폰입니다. 사용하실 수 없습니다."
								Response.End

						ELSEIF oRs("CollectFlag") = "Y" THEN
								oConn.RollbackTrans

								oRs.Close
								Set oRs = Nothing
								oConn.Close
								Set oConn = Nothing

								Response.Write "FAIL|||||[" & CouponName & "] 쿠폰은 회수된 쿠폰입니다. 사용하실 수 없습니다."
								Response.End

						ELSEIF oRs("StartDT") > U_DATE & LEFT(U_TIME,4) THEN
								oConn.RollbackTrans

								oRs.Close
								Set oRs = Nothing
								oConn.Close
								Set oConn = Nothing

								Response.Write "FAIL|||||[" & CouponName & "] 쿠폰은 아직 사용하실 수 없습니다."
								Response.End

						ELSEIF oRs("EndDT") < U_DATE & LEFT(U_TIME,4) THEN
								oConn.RollbackTrans

								oRs.Close
								Set oRs = Nothing
								oConn.Close
								Set oConn = Nothing

								Response.Write "FAIL|||||[" & CouponName & "] 쿠폰은 사용기한이 지났습니다. 사용하실 수 없습니다."
								Response.End

						ELSEIF oRs("MobileFlag") <> "Y" THEN
								oConn.RollbackTrans

								oRs.Close
								Set oRs = Nothing
								oConn.Close
								Set oConn = Nothing

								Response.Write "FAIL|||||[" & CouponName & "] 쿠폰은 모바일웹에서 사용불가한 쿠폰입니다."
								Response.End

						ELSEIF oRs("DeliveryCouponFlag") = "Y" THEN
								oConn.RollbackTrans

								oRs.Close
								Set oRs = Nothing
								oConn.Close
								Set oConn = Nothing

								Response.Write "FAIL|||||[" & CouponName & "] 쿠폰은 주문시 사용불가한 쿠폰입니다."
								Response.End
						END IF
				ELSE
						oConn.RollbackTrans

						oRs.Close
						Set oRs = Nothing
						oConn.Close
						Set oConn = Nothing

						Response.Write "FAIL|||||없는 쿠폰입니다. 사용하실 수 없습니다."
						Response.End
				END IF
				oRs.Close


				'# 4-2. 해당 주문중에 다른 상품에 사용한 쿠폰인지 체크
				Set oCmd = Server.CreateObject("ADODB.Command")
				WITH oCmd
						.ActiveConnection = oConn
						.CommandType = adCmdStoredProc
						.CommandText = "USP_Front_EShop_OrderSheet_UseCoupon_Select_For_Used_Check"
						.Parameters.Append .CreateParameter("@CartID",				adVarChar,	 adParamInput, 20,	U_CARTID)
						.Parameters.Append .CreateParameter("@OrderSheetIdx",		adInteger,	 adParamInput,   ,	OrderSheetIdx)
						.Parameters.Append .CreateParameter("@MemberCouponIdx",		adInteger,	 adParamInput,   ,	MemberCouponIdx(i))
				END WITH
				oRs.Open oCmd, , adOpenStatic, adLockReadOnly
				Set oCmd = Nothing

				IF NOT oRs.EOF THEN
						oConn.RollbackTrans

						oRs.Close
						Set oRs = Nothing
						oConn.Close
						Set oConn = Nothing

						Response.Write "FAIL|||||[" & CouponName & "] 쿠폰은 다른 상품에 적용한 쿠폰입니다."
						Response.End
				END IF
				oRs.Close


				'# 4-3. 해당상품에 적용가능 여부
				Set oCmd = Server.CreateObject("ADODB.Command")
				WITH oCmd
						.ActiveConnection = oConn
						.CommandType = adCmdStoredProc
						.CommandText = "USP_Front_EShop_Coupon_Member_Select_For_UseCheck"
						.Parameters.Append .CreateParameter("@MemberNum",			adInteger,	 adParamInput,   ,	 U_NUM)
						.Parameters.Append .CreateParameter("@OrderSheetIdx",		adInteger,	 adParamInput,   ,	 OrderSheetIdx)
						.Parameters.Append .CreateParameter("@MemberCouponIdx",		adInteger,	 adParamInput,   ,	 MemberCouponIdx(i))
				END WITH
				oRs.Open oCmd, , adOpenStatic, adLockReadOnly
				Set oCmd = Nothing
	

				'# 4-3-1. 해당상품에 적용가능한 쿠폰인 경우
				IF NOT oRs.EOF THEN

						CouponIdx			= oRs("CouponIdx")
						CouponName			= oRs("CouponName")
						CouponType			= oRs("CouponType")
						LimitPriceType		= oRs("LimitPriceType")
						LimitPrice			= oRs("LimitPrice")
						MoneyType			= oRs("MoneyType")
						Discount			= oRs("Discount")
						ApplyPriceType		= oRs("ApplyPriceType")
						LimitDiscountFlag	= oRs("LimitDiscountFlag")
						LimitDiscount		= oRs("LimitDiscount")
						DuplicateUseFlag	= oRs("DuplicateUseFlag")

						'# 임직원 쿠폰일 경우 다른 쿠폰과 사용 금지
						IF CouponType = "99" AND UBound(MemberCouponIdx) > 0 THEN
								oConn.RollbackTrans

								oRs.Close
								Set oRs = Nothing
								oConn.Close
								Set oConn = Nothing

								Response.Write "FAIL|||||[" & CouponName & "] 쿠폰은 다른 쿠폰과 중복 사용할 수 없습니다."
								Response.End
						END IF

						'# 판매가 제한
						IF LimitPriceType = "W" THEN
								IF CDbl(SalePrice) < CDbl(LimitPrice) THEN
										oConn.RollbackTrans

										oRs.Close
										Set oRs = Nothing
										oConn.Close
										Set oConn = Nothing

										Response.Write "FAIL|||||[" & CouponName & "] 쿠폰은 " & FormatNumber(LimitPrice,0) & "원이상 상품에만 사용할 수 있습니다."
										Response.End
								END IF

						'# 할인율 제한
						ELSEIF LimitPriceType = "P" THEN
								IF CDbl(DCRate) > (100 - CDbl(LimitPrice)) THEN
										oConn.RollbackTrans

										oRs.Close
										Set oRs = Nothing
										oConn.Close
										Set oConn = Nothing

										Response.Write "FAIL|||||[" & CouponName & "] 쿠폰은 " & FormatNumber(100 - LimitPrice,0) & "% 미만 할인된 상품에만 사용할 수 있습니다."
										Response.End
								END IF
						END IF


						'# 할인금액 계산
						IF MoneyType = "W" THEN
								DiscountPrice	 = Discount
						ELSE
								IF ApplyPriceType = "T" THEN
										DiscountPrice	 = Round(TagPrice * CDbl(Discount) / 1000) * 10
								ELSE
										DiscountPrice	 = Round(SalePrice * CDbl(Discount) / 1000) * 10
								END IF
						END IF

						'# 최대할인금액 적용
						IF LimitDiscountFlag = "Y" AND CDbl(DiscountPrice) > CDbl(LimitDiscount) THEN
								DiscountPrice	 = LimitDiscount
						END IF

						'# 최소 주문금액 체크
						IF DiscountPrice > (OrderPrice - MALL_MIN_ORDERPRICE) THEN
								DiscountPrice	= OrderPrice - MALL_MIN_ORDERPRICE
						END IF

						'# 할인금액이 0보다 작을 경우
						IF DiscountPrice <= 0 THEN
								oConn.RollbackTrans

								oRs.Close
								Set oRs = Nothing
								oConn.Close
								Set oConn = Nothing

								Response.Write "FAIL|||||쿠폰할인금액이 0 보다 작아 쿠폰을 사용할 수 없습니다."
								Response.End
						END IF

						'# 총 할인금액
						TotalDiscountPrice		= TotalDiscountPrice + DiscountPrice

						OrderPrice				 = OrderPrice - DiscountPrice


						'# 똑같은 쿠폰이 여러장인지 체크
						'# 상품 상세에서 쿠폰 여러 장 다운 후 여러장을 한 상품에 사용할 경우 제거
						IF InStr(CouponIdxs, "," & CouponIdx & ",") THEN
								oConn.RollbackTrans

								oRs.Close
								Set oRs = Nothing
								oConn.Close
								Set oConn = Nothing

								Response.Write "FAIL|||||[" & CouponName & "] 쿠폰은 하나만 사용 가능합니다."
								Response.End
						END IF


						'# 중복 사용 가능 쿠폰인지 체크
						'# 해당 쿠폰이 쇼핑몰쿠폰 중복 불가능 쿠폰이면, 다른 쇼핑몰쿠폰이 사용되었는지 체크해서 제거
						IF DuplicateUseFlag = "N" AND CouponIdxs <> "" THEN
								oConn.RollbackTrans

								oRs.Close
								Set oRs = Nothing
								oConn.Close
								Set oConn = Nothing

								Response.Write "FAIL|||||[" & CouponName & "] 쿠폰은 다른 쿠폰과 같이 사용 하실 수 없습니다."
								Response.End
						END IF
	

						'# 중복 사용 가능 쿠폰인지 체크
						'# 쇼핑몰 중복 불가능 쿠폰과 같이 사용하는 경우
						IF DuplicateUseFlag = "Y" AND InStr(DuplicateUseFlags, "N") THEN
								oConn.RollbackTrans

								oRs.Close
								Set oRs = Nothing
								oConn.Close
								Set oConn = Nothing

								Response.Write "FAIL|||||[" & CouponName & "] 쿠폰은 단독사용 쿠폰과 같이 사용 하실 수 없습니다."
								Response.End
						END IF


						CouponIdxs			= CouponIdxs		& "," & CouponIdx			& ","
						DuplicateUseFlags	= DuplicateUseFlags & "," & DuplicateUseFlag	& ","


						Set oCmd = Server.CreateObject("ADODB.Command")
						WITH oCmd
								.ActiveConnection = oConn
								.CommandType = adCmdStoredProc
								.CommandText = "USP_Front_EShop_OrderSheet_UseCoupon_Temp_Insert"

								.Parameters.Append .CreateParameter("@MemberNum",				adInteger,	 adParamInput,    ,		U_NUM)
								.Parameters.Append .CreateParameter("@OrderSheetIdx",			adInteger,	 adParamInput,    ,		OrderSheetIdx)
								.Parameters.Append .CreateParameter("@MemberCouponIdx",			adInteger,	 adParamInput,    ,		MemberCouponIdx(i))
								.Parameters.Append .CreateParameter("@CouponName",				adVarChar,	 adParamInput, 100,		CouponName)
								.Parameters.Append .CreateParameter("@MoneyType",				adChar,		 adParamInput,   1,		MoneyType)
								.Parameters.Append .CreateParameter("@DiscountPrice",			adInteger,	 adParamInput,    ,		DiscountPrice)
								.Parameters.Append .CreateParameter("@DiscountRate",			adDouble,	 adParamInput,    ,		FormatNumber( (DiscountPrice/TagPrice)*100, 3 ))
								.Parameters.Append .CreateParameter("@DiscountRateOriginal",	adDouble,	 adParamInput,    ,		Discount)
								.Parameters.Append .CreateParameter("@ApplyPriceType",			adChar,		 adParamInput,   1,		ApplyPriceType)
								.Parameters.Append .CreateParameter("@CouponIdx",				adInteger,	 adParamInput,    ,		CouponIdx)
								.Parameters.Append .CreateParameter("@CouponType",				adChar,		 adParamInput,   2,		CouponType)
								.Parameters.Append .CreateParameter("@DuplicateUseFlag",		adChar,		 adParamInput,   1,		DuplicateUseFlag)

								.Execute, , adExecuteNoRecords
						END WITH
						Set oCmd = Nothing

						IF Err.Number <> 0 THEN
								oConn.RollbackTrans

								oRs.Close
								SET oRs = Nothing
								oConn.Close
								SET oConn = Nothing

								Response.Write "FAIL|||||쿠폰적용 처리 중 오류가 발생하였습니다.[2]"
								Response.End
						END IF

				'# 4-3-2. 해당상품에 적용가능한 쿠폰이 아닌 경우
				ELSE
					oConn.RollbackTrans

					oRs.Close
					Set oRs = Nothing
					oConn.Close
					Set oConn = Nothing

					Response.Write "FAIL|||||[" & CouponName & "] 쿠폰은 해당 상품에 적용가능한 쿠폰이 아닙니다."
					Response.End
				END IF
				oRs.Close

		NEXT
END IF

'# 쿠폰 유효성 체크일 경우
IF ApplyFlag = "" OR ApplyFlag = "N" THEN
		Response.Write "OK|||||" & FormatNumber(TotalDiscountPrice,0) & "|" & FormatNumber(UsePointPrice,0) & "|" & FormatNumber(UseScashPrice,0) & "|" & FormatNumber(OrderPrice,0) & "|"

'# 쿠폰 적용일 경우
ELSE
		'# 사용 쿠폰 정보 임시 테이블
		Set oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection = oConn
				.CommandType = adCmdStoredProc
				.CommandText = "USP_Front_EShop_OrderSheet_UseCoupon_Temp_For_Move_EShop_OrderSheet_UseCoupon"
				.Parameters.Append .CreateParameter("@MemberNum",		 adInteger,	 adParamInput, ,	 U_NUM)
				.Parameters.Append .CreateParameter("@OrderSheetIdx",	 adInteger,	 adParamInput, ,	 OrderSheetIdx)

				.Execute, , adExecuteNoRecords
		END WITH
		Set oCmd = Nothing

		IF Err.Number <> 0 THEN
				oConn.RollbackTrans

				SET oRs = Nothing
				oConn.Close
				SET oConn = Nothing

				Response.Write "FAIL|||||쿠폰적용 처리 중 오류가 발생하였습니다.[3]"
				Response.End
		END IF


		Response.Write "OK|||||0|0|0|0|"
END IF


oConn.CommitTrans


Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>