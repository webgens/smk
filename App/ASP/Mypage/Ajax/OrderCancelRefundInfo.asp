<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'OrderCancelRefundInfo.asp - 주문취소시 환불예산금액 폼 페이지
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
DIM OPIdx

DIM PayType
DIM EscrowFlag
DIM TotalSettlePrice

DIM OrderPrice			: OrderPrice		= 0
DIM SalePrice			: SalePrice			= 0
DIM DeliveryPrice		: DeliveryPrice		= 0
DIM AddDeliveryPrice	: AddDeliveryPrice	= 0

DIM DiscountPrice		: DiscountPrice		= 0
DIM UseCouponPrice		: UseCouponPrice	= 0
DIM UsePointPrice		: UsePointPrice		= 0
DIM UseScashPrice		: UseScashPrice		= 0

DIM RefundPrice			: RefundPrice		= 0

DIM SetStandardPrice	: SetStandardPrice	= 0
DIM SetDeliveryPrice	: SetDeliveryPrice	= 0
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


OrderCode			= sqlFilter(Request("OrderCode"))
OPIdx				= sqlFilter(Request("OPIdx"))




IF OrderCode = "" THEN
		Response.Write "FAIL|||||선택한 주문번호가 없습니다."
		Response.End
END IF



SET oConn	= ConnectionOpen()	'//커넥션 생성
SET oRs		= Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성
SET oRs1	= Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성


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
		PayType				= oRs("PayType")
		EscrowFlag			= oRs("EscrowFlag")
		TotalSettlePrice	= oRs("OrderPrice") + oRs("DeliveryPrice")

		'# SalePrice		= oRs("SalePrice")
		'# DeliveryPrice	= oRs("DeliveryPrice")
		'# OrderPrice		= SalePrice + DeliveryPrice
		'# 
		'# UseCouponPrice	= oRs("UseCouponPrice")
		'# UsePointPrice	= oRs("UsePointPrice")
		'# UseScashPrice	= oRs("UseScashPrice")
		'# DiscountPrice	= UseCouponPrice + UsePointPrice + UseScashPrice
END IF
oRs.Close


'# 취소상품 금액
IF OPIdx <> "" THEN
		wQuery = ""
		wQuery = wQuery & "WHERE A.OrderCode = '" & OrderCode & "' "
		wQuery = wQuery & "AND A.ProductType = 'P' "
		wQuery = wQuery & "AND A.Idx IN (" & OPIdx & ") "


		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Order_Product_Select_For_Vendor_TotalPrice"

				.Parameters.Append .CreateParameter("@WQUERY",		adVarchar, adParaminput,	1000,	wQuery)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		SET oCmd = Nothing

		IF NOT oRs.EOF THEN
				Do Until oRs.EOF
						SalePrice		= SalePrice			+ oRs("SalePrice")
						UseCouponPrice	= UseCouponPrice	+ oRs("UseCouponPrice")
						UsePointPrice	= UsePointPrice		+ oRs("UsePointPrice")
						UseScashPrice	= UseScashPrice		+ oRs("UseScashPrice")
						DiscountPrice	= DiscountPrice		+ UseCouponPrice + UsePointPrice + UseScashPrice

						'# 배송비 설정 가져오기
						SET oCmd = Server.CreateObject("ADODB.Command")
						WITH oCmd
								.ActiveConnection	 = oConn
								.CommandType		 = adCmdStoredProc
								.CommandText		 = "USP_Admin_EShop_Store_Select_By_ShopCD"

								.Parameters.Append .CreateParameter("@ShopCD",		adChar, adParaminput,	6,	oRs("Vendor"))
						END WITH
						oRs1.CursorLocation = adUseClient
						oRs1.Open oCmd, , adOpenStatic, adLockReadOnly
						SET oCmd = Nothing

						IF NOT oRs1.EOF THEN
								SetStandardPrice	= oRs1("StandardPrice")
								SetDeliveryPrice	= oRs1("DeliveryPrice")
						END IF
						oRs1.Close

						'# 배송비 환불금액 계산
						wQuery = ""
						wQuery = wQuery & "WHERE A.OrderCode = '" & OrderCode & "' "
						wQuery = wQuery & "AND A.ProductType = 'P' "
						wQuery = wQuery & "AND A.DelvType = 'P' "
						wQuery = wQuery & "AND A.OrderState <> 'C' "
						wQuery = wQuery & "AND A.Vendor = '" & oRs("Vendor") & "' "
						wQuery = wQuery & "AND A.Idx NOT IN (" & OPIdx & ") "

						SET oCmd = Server.CreateObject("ADODB.Command")
						WITH oCmd
								.ActiveConnection	 = oConn
								.CommandType		 = adCmdStoredProc
								.CommandText		 = "USP_Front_EShop_Order_Product_Select_For_TotalPrice"

								.Parameters.Append .CreateParameter("@WQUERY",		adVarChar, adParaminput,	1000,	wQuery)
						END WITH
						oRs1.CursorLocation = adUseClient
						oRs1.Open oCmd, , adOpenStatic, adLockReadOnly
						SET oCmd = Nothing

						IF NOT oRs1.EOF THEN
								'# 남아있는 일반택배 주문상품이 없을 경우 결제한 배송비 환불
								IF CDbl(oRs1("SalePrice")) = 0 THEN
										DeliveryPrice	= CDbl(DeliveryPrice) + CDbl(oRs("DeliveryPrice"))

								'# 남아있는 주문상품 금액이 배송비 기준금액 미만이면 결제금액에서 차감할 추가 배송비 적용
								ELSEIF CDbl(oRs1("SalePrice")) < CDbl(SetStandardPrice) THEN
										'# 최초 주문시 무료배송이었을 경우 환불금액에서 차감할 배송비 적용
										IF CDbl(oRs("DeliveryPrice")) = 0 THEN
												AddDeliveryPrice	= CDbl(AddDeliveryPrice) + CDbl(SetDeliveryPrice)
										END IF
								END IF
						ELSE
								DeliveryPrice	= CDbl(DeliveryPrice) + CDbl(oRs("DeliveryPrice"))
						END IF
						oRs1.Close

						oRs.MoveNext
				Loop
		END IF
		oRs.Close
END IF

OrderPrice		= SalePrice + DeliveryPrice
RefundPrice		= SalePrice + DeliveryPrice - DiscountPrice - AddDeliveryPrice

Response.Write "OK|||||"
%>					
                                <li class="detailList">
                                    <div class="tit">결제금액</div>
                                    <div class="cont">
                                        <span class="general"><em class="strong"><%=FormatNumber(OrderPrice, 0)%></em>원</span>
                                    </div>
                                </li>
                                <li class="detailList">
                                    <div class="tit">차감/할인</div>
                                    <div class="cont">
                                        <span class="general"><span class="art-1">쿠폰할인</span><span><%=FormatNumber(UseCouponPrice, 0)%>원</span></span>
                                        <span class="general"><span class="art-1">포인트</span><span><%=FormatNumber(UsePointPrice, 0)%>원</span></span>
                                        <span class="general"><span class="art-1">슈즈상품권</span><span><%=FormatNumber(UseScashPrice, 0)%>원</span></span>
                                        <span class="general"><span class="art-1">추가 배송비</span><span><%=FormatNumber(AddDeliveryPrice, 0)%>원</span></span>
                                    </div>
                                </li>
                                <li class="detailList">
                                    <div class="tit">
										<em class="strong">환불 예정
											<%
											IF PayType = "C" OR PayType = "M" THEN
													SELECT CASE PayType
														CASE "C" : Response.Write "(신용카드 취소)"
														CASE "M" : Response.Write "(모바일결제 취소)"
													END SELECT
											ELSEIF PayType = "N" THEN
													Response.Write "(네이버페이 취소)"
											ELSE
													Response.Write "(환불계좌 환불)"
											END IF
											%>
										</em>
                                    </div>
                                    <div class="cont">
                                        <span class="general ty-red"><em class="strong"><%=FormatNumber(RefundPrice, 0)%></em>원</span>
										<input type="hidden" name="RefundPrice" value="<%=RefundPrice%>" />
                                    </div>
                                </li>
<%
Set oRs1 = Nothing
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>