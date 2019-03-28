<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'Main_BestBrandsProductList.asp - 메인페이지 BEST BRANDS 상품 리스트
'Date		: 2018.12.21
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

DIM Idx
DIM BrandCode
	
DIM DCRate
DIM Badge
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'

	
Idx				 = sqlFilter(Request("Idx"))
BrandCode		 = sqlFilter(Request("BrandCode"))

	

SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성





Response.Write "OK|||||"



IF U_NUM = "" THEN U_NUM = 0



SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Mobile_EShop_MainBanner_Product_Select_Top4_By_MainBannerIdx"

		.Parameters.Append .CreateParameter("@MainBannerIdx",	 adInteger, adParamInput,  , Idx)
		.Parameters.Append .CreateParameter("@MemberNum",		 adInteger, adParamInput,  , U_NUM)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing
																
IF NOT oRs.EOF THEN
%>
								<ul class="listview">
<%
		Do Until oRs.EOF
%>
									<li>
										<a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>')" class="listitems">
											<div class="badgegroup">
												<%=ProductBadgeNew(oRs("ProductCode"), oRs("DiscountRate"), oRs("ReserveFlag"), oRs("OPOFlag"), oRs("PickupFlag"), oRs("GiftCnt"), oRs("CouponIdx"))%>
											</div>
											<div class="thumbnail"><img src="<%=oRs("ImageUrl_0320")%>" alt="<%=REPLACE(oRs("ProductName"), """", "")%>"></div>
											<p class="brand-name"><%=oRs("BrandName")%></p>
											<h1 class="product-name pname"><%=oRs("ProductName")%></h1>
											<p class="price"><strong><%=FormatNumber(oRs("SalePrice"), 0)%></strong>원</p>
										</a>
										<a nohref class="listitems">
											<p class="optional-info">
												<button type="button" class="btn-size" onclick="SizeLayerOpen('<%=oRs("ProductCode")%>');">SIZE</button>
												<span class="icon ico-fav"><%=oRs("WishCnt")%></span>
												<span class="icon ico-cmt"><%=oRs("ReviewCnt")%></span>
											</p>
										</a>
									</li>
<%
				oRs.MoveNext
		Loop
%>
								</ul>
<%
END IF
oRs.Close



SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>
