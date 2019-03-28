<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'OrderReturnRefundInfo.asp - 주문반품시 환불예산금액 폼 페이지
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
DIM DelvFeeType

DIM PayType
DIM EscrowFlag
DIM TotalSettlePrice
DIM DelvType
DIM DelvFee

DIM Vendor
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
DelvFeeType			= sqlFilter(Request("DelvFeeType"))




IF OrderCode = "" OR OPIdx = "" THEN
		Response.Write "FAIL|||||선택한 주문번호가 없습니다."
		Response.End
END IF



SET oConn	= ConnectionOpen()	'//커넥션 생성
SET oRs		= Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성


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

ELSE
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||반품할 주문내역이 없습니다."
		Response.End
END IF
oRs.Close



wQuery = ""
wQuery = wQuery & "WHERE A.IsShowFlag = 'Y' AND A.ProductType = 'P' "
wQuery = wQuery & "AND A.OrderCode = '" & OrderCode & "' "
wQuery = wQuery & "AND A.Idx = " & OPIdx & " "
IF U_NUM <> "" THEN
		wQuery = wQuery & "AND B.UserID = '" & U_NUM & "' "
ELSE
		wQuery = wQuery & "AND B.OrderName = '" & N_NAME & "' AND B.OrderHp = '" & N_HP & "' AND B.OrderEmail = '" & N_EMAIL & "' "
END IF

sQuery = ""

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
		DelvType		= oRs("DelvType")
		DelvFee			= oRs("VendorDeliveryPrice")		'# 반품 배송비
		Vendor			= oRs("Vendor")
END IF
oRs.Close


'# 반품상품 금액
wQuery = ""
wQuery = wQuery & "WHERE A.OrderCode = '" & OrderCode & "' "
wQuery = wQuery & "AND A.ProductType = 'P' "
wQuery = wQuery & "AND A.Idx = " & OPIdx & " "


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
		SalePrice		= oRs("SalePrice")
		UseCouponPrice	= oRs("UseCouponPrice")
		UsePointPrice	= oRs("UsePointPrice")
		UseScashPrice	= oRs("UseScashPrice")
		DiscountPrice	= UseCouponPrice + UsePointPrice + UseScashPrice

		'# 일반택배 주문일 경우 배송비 결제내역이 없으면 추가배송비 산정
		IF DelvType = "P" AND (oRs("OrderCnt_P") > 1 OR oRs("DeliveryPrice") = 0) THEN
				AddDeliveryPrice	= DelvFee
		END IF
END IF
oRs.Close


OrderPrice		= SalePrice
RefundPrice		= SalePrice - DiscountPrice

'# 배송비 환불금액에서 차감일 경우 배송비 환불금액 계산
IF DelvFeeType = "5" THEN
		DiscountPrice	= DiscountPrice + DelvFee + AddDeliveryPrice
		RefundPrice		= RefundPrice - DelvFee - AddDeliveryPrice
END IF



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
										<%IF DelvFeeType = "5" AND AddDeliveryPrice > 0 THEN%>
                                        <span class="general"><span class="art-1">왕복배송비</span><span><%=FormatNumber(DelvFee + AddDeliveryPrice, 0)%>원</span></span>
										<%ELSEIF DelvFeeType = "5" AND AddDeliveryPrice = 0 THEN%>
                                        <span class="general"><span class="art-1">반품배송비</span><span><%=FormatNumber(DelvFee, 0)%>원</span></span>
										<%END IF%>
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
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>