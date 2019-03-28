<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'Main_NowBestProductList.asp - 메인페이지 지금뜨는상품 상품 리스트
'Date		: 2018.12.23
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
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'

	
Idx				 = sqlFilter(Request("Idx"))

	

SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성





Response.Write "OK|||||"



SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_MainBanner_Product_Select_Top3_By_MainBannerIdx"

		.Parameters.Append .CreateParameter("@MainBannerIdx", adInteger, adParamInput, 20, Idx)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing
																
IF NOT oRs.EOF THEN
%>
							<div class="swiper-container ranking-slider NBProductList">
								<ol class="swiper-wrapper">
<%
		i = 1
		Do Until oRs.EOF
%>
									<li class="swiper-slide">
										<a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>')" class="listitems">
											<div class="thumbnail"><img src="<%=oRs("ImageUrl_0320")%>" style="width:100%;" alt="<%=oRs("ProductCD")%>"></div>
											<p class="brand-name"><%=oRs("BrandName")%></p>
											<h1 class="product-name pname"><%=oRs("ProductName")%></h1>
											<p class="price"><strong><%=FormatNumber(oRs("SalePrice"), 0)%></strong>원</p>
										</a>
									</li>
<%
				oRs.MoveNext
				i = i + 1
		Loop
%>
								</ol>
							</div>

							<script type="text/javascript">
								$(function () {
									/* 상품목록 랭킹 슬라이드 init */
									var NBProductListSlider = new Swiper('.NBProductList', {
										slidesPerView: 'auto',
										spaceBetween: 20,
										//centeredSlides: true,
										observer: true,
										observeParents: true,
										on: {
											observerUpdate: true
										}
									});
								});
							</script>
<%
END IF
oRs.Close



SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>
