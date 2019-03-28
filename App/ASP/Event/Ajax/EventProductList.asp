<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'EventProductList.asp - 이벤트 상품 리스트
'Date		: 2019.01.12
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

<%
'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn						'# ADODB Connection 개체
DIM oRs							'# ADODB Recordset 개체
DIM oCmd						'# ADODB Command 개체

DIM i
DIM j
DIM x
DIM y

DIM wQuery						'# WHERE 절
DIM sQuery						'# SORT 절

DIM Page
DIM PageSize : PageSize = 12
DIM RecCnt
DIM PageCnt

Dim EventIDX
Dim CategoryIDX

Dim CateName
Dim ProductCount
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

EventIDX			 = sqlFilter(Request("EventIDX"))
CategoryIDX			 = sqlFilter(Request("CategoryIDX"))

SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성

SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Event_Category_Select_By_IDX"
		.Parameters.Append .CreateParameter("@IDX", adInteger, adParamInput, , CategoryIDX)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

CateName = oRs("CateName")
ProductCount = oRs("ProductCount")

oRs.Close
	
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Mobile_EShop_Event_Category_Product_Select_By_EventIDX_CategoryIDX"
		.Parameters.Append .CreateParameter("@EventIDX", adInteger, adParamInput, , EventIDX)
		.Parameters.Append .CreateParameter("@CategoryIDX", adInteger, adParamInput, , CategoryIDX)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

Response.Write CateName & " (" & oRs.RecordCount & ")|||||"

If Not oRs.EOF Then
	Do While Not oRs.EOF
%>
									<li>
										<a href="javascript:void(0)" class="listitems" onclick="APP_GoUrl('/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>')">
											<div class="badgegroup">
												<%=ProductBadgeNew(oRs("ProductCode"), oRs("DiscountRate"), oRs("ReserveFlag"), oRs("OPOFlag"), oRs("PickupFlag"), oRs("GiftCnt"), oRs("CouponIdx"))%>
											</div>
											<div class="thumbnail"><img src="<%=oRs("ImageUrl")%>" alt="<%=oRs("ProductName")%>"></div>
											<p class="brand-name"><%=oRs("BrandName")%></p>
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
%>


<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>