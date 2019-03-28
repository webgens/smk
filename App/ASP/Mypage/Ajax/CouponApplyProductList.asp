<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'CouponApplyProductList.asp - 쿠폰적용상품 리스트
'Date		: 2019.01.25
'Update		: 
'*****************************************************************************************'

'# 페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'-----------------------------------------------------------------------------------------'

%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->
<!-- #include virtual="/Common/CheckID.asp" -->

<%

'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn						'# ADODB Connection 개체
DIM oRs							'# ADODB Recordset 개체
DIM oRs1						'# ADODB Recordset 개체
DIM oCmd						'# ADODB Command 개체

DIM i
DIM j
DIM x
DIM y

DIM wQuery						'# WHERE 절
DIM sQuery						'# SORT 절

DIM Page
DIM PageSize : PageSize = 10
DIM RecCnt
DIM PageCnt

Dim StoreProcName

Dim Idx
Dim ISTopN
Dim ProductImage
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

Page			 = sqlFilter(Request("Page"))
Idx				 = sqlFilter(Request("Idx"))
ISTopN			 = sqlFilter(Request("ISTopN"))

IF Page			 = "" THEN Page			 = 1


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성

wQuery = "WHERE A.MemberNum = " & U_NUM & "  AND A.Idx = " & Idx & " AND C.SaleState = 'Y' "
sQuery = "ORDER BY C.ProductCode DESC "

If IsTopN = "Y" Then
	StoreProcName = "USP_Mobile_EShop_Coupon_Member_Select_For_Apply_ProductList_HistoryBack"
Else
	StoreProcName = "USP_Mobile_EShop_Coupon_Member_Select_For_Apply_ProductList"
End If

SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = StoreProcName

		.Parameters.Append .CreateParameter("@PAGE",		 adInteger, adParaminput,		, Page)
		.Parameters.Append .CreateParameter("@PAGE_SIZE",	 adInteger, adParaminput,		, PageSize)
		.Parameters.Append .CreateParameter("@wQuery",		 adVarchar, adParaminput, 1000	, wQuery)
		.Parameters.Append .CreateParameter("@sQuery",		 adVarchar, adParaminput, 100	, sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

RecCnt	 = oRs(0)

PageCnt	 = FIX((RecCnt+(PageSize-1))/PageSize)

SET oRs = oRs.NextrecordSet

If oRs.EOF Then
%>
						<li class="no-products" style="width:95%;">
							<p>해당하는 상품이 없습니다.</p>
						</li>
<%
Else
	Do While Not oRs.EOF
		ProductImage = oRs("ProductImage")
		IF ProductImage = "" THEN ProductImage = "/Images/180_noimage.png"
%>
                        <li>
                            <a href="/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>" class="listitems" onclick="pushHash();">
                                <div class="badgegroup">
									<%=ProductBadgeNew(oRs("ProductCode"), oRs("DiscountRate"), oRs("ReserveFlag"), oRs("OPOFlag"), oRs("PickupFlag"), oRs("GiftCnt"), oRs("CouponIdx"))%>
                                </div>
                                <div class="thumbnail"><img src="<%=ProductImage%>" alt="<%=oRs("ProductName")%>"></div>
                                <p class="brand-name"><%=oRs("BrandName")%> /<%=oRs("ProductCode")%></p>
                                <h1 class="product-name pname"><%=oRs("ProductName")%></h1>
                                <p class="price"><strong><%=FormatNumber(oRs("SalePrice"), 0)%></strong>원</p>
                            </a>
							<a nohref class="listitems">
                                <p class="optional-info">
                                    <button type="button" class="btn-size" onclick="SizeLayerOpen('<%=oRs("ProductCode")%>');">SIZE</button>
                                    <span class="icon ico-fav"><%=FormatNumber(oRs("WishCnt"), 0)%></span>
                                    <span class="icon ico-cmt"><%=FormatNumber(oRs("ReviewCnt"), 0)%></span>
                                </p>
							</a>
                        </li>
<%
		oRs.MoveNext
	Loop
End If
oRs.Close

Response.Write "|||||" & RecCnt & "|||||" & PageCnt

SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>
