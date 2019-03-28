<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'OrderSheetUsePoint.asp - 주문서 포인트 사용 폼 페이지
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
DIM oRs1											'# ADODB Recordset 개체
DIM oCmd											'# ADODB Command 개체

DIM i
DIM j
DiM x
DIM y

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절

DIM OrderSheetIdx

DIM ProductCode
DIM ProductName
DIM SizeCD
DIM ProdCD
DIM ColorCD
DIM BrandName
DIM ProductImage
DIM OrderCnt
DIM TagPrice
DIM SalePriceType
DIM SalePrice
DIM DCRate
DIM UseCouponPrice
DIM UsePointPrice
DIM UseScashPrice

DIM TotalPoint				: TotalPoint			= 0
DIM TotalUsePointPrice		: TotalUsePointPrice	= 0
DIM UsablePointPrice
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


OrderSheetIdx			= sqlFilter(Request("OrderSheetIdx"))




IF OrderSheetIdx = "" THEN
		Response.Write "FAIL|||||상품이 없습니다."
		Response.End
END IF


SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성


'# 주문서 상품 내역
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
				Response.Write "FAIL|||||임직원 판매가 구매시 포인트를 사용하실 수 없습니다."
				Response.End
		END IF

		ProductCode		= oRs("ProductCode")
		ProductName		= oRs("ProductName")
		BrandName		= oRs("BrandName")
		SizeCD			= oRs("SizeCD")
		ProdCD			= oRs("ProdCD")
		ColorCD			= oRs("ColorCD")
		ProductImage	= oRs("ProductImage")
		OrderCnt		= oRs("OrderCnt")
		TagPrice		= oRs("TagPrice")
		SalePriceType	= oRs("SalePriceType")
		IF SalePriceType = "2" THEN
				SalePrice			= oRs("EmployeeSalePrice")
				DCRate				= oRs("EmployeeDCRate")
		ELSE
				SalePrice			= oRs("SalePrice")
				DCRate				= oRs("DCRate")
		END IF
		UseCouponPrice	= oRs("UseCouponPrice")
		UsePointPrice	= oRs("UsePointPrice")
		UseScashPrice	= oRs("UseScashPrice")

		IF ProductImage = "" THEN
				ProductImage	= "/Images/180_noimage.png"
		END IF
ELSE
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||선택한 상품이 없습니다."
		Response.End
END IF
oRs.Close


'# 보유 포인트
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_Member_Select_By_MemberNum"

		.Parameters.Append .CreateParameter("@MemberNum",	adInteger,	adParamInput, ,		U_NUM)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		TotalPoint		= oRs("Point")
END IF
oRs.Close


'# 주문서에서 사용한 총 포인트
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_OrderSheet_Select_For_UseDiscount_By_CartID"

		.Parameters.Append .CreateParameter("@CartID",	adVarChar,	adParamInput, 20,		U_CARTID)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		TotalUsePointPrice	= oRs("TotalUsePointPrice")
END IF
oRs.Close

'# 사용가능 포인트
UsablePointPrice	= CDbl(TotalPoint) - CDbl(TotalUsePointPrice) + CDbl(UsePointPrice)

'# 상품별 결제 최소금액
DIM MinOrderPrice
MinOrderPrice		= CDbl(SalePrice) - CDbl(UseCouponPrice) - CDbl(UseScashPrice) - CDbl(MALL_MIN_ORDERPRICE)
IF UsablePointPrice > MinOrderPrice THEN
		UsablePointPrice	= MinOrderPrice
END IF


Response.Write "OK|||||"
%>					
        <div class="area-pop" id="UsePoint">
            <div class="full">
                <div class="tit-pop">
                    <p class="tit">포인트 사용</p>
                    <button type="button" onclick="closePop('DimDepth1')" class="btn-hide-pop">닫기</button>
                </div>

                <div class="container-pop">
                    <div class="contents">
                        <div class="inquiry">
                            <div class="hold-wrap">
                                <p>보유 포인트</p>
                                <strong><%=FormatNumber(TotalPoint,0)%>P</strong>
                            </div>
                            <div class="hold-wrap able">
                                <p>적용가능 포인트</p>
                                <strong><%=FormatNumber(UsablePointPrice,0)%>P</strong>
                            </div>
                            <div class="inf-type1">
                                <p class="tit">포인트는 최소 5,000P부터 현금처럼 사용 가능합니다.</p>
                            </div>
                        </div>
                        <div class="usage">
                            <span class="input">
								<input type="hidden" name="UsablePointPrice" value="<%=UsablePointPrice%>" />
								<input type="text" name="UsePointPrice" value="<%=FormatNumber(UsePointPrice,0)%>" />
                                <span class="point">P</span>
                            </span>
                            <div class="fieldset">
                                <button type="button" onclick="usePointAll()">모두 사용</button>
                            </div>
                        </div>
                    </div>
                    <div class="btns">
                        <button type="button" onclick="usePoint(<%=OrderSheetIdx%>)" class="button ty-red">적용</button>
                    </div>
                </div>
            </div>
        </div>

<%
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>