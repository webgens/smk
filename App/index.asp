<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>

<%
'*****************************************************************************************'
'Index.asp - 메인페이지
'Date		: 2018.10.29
'Update		: 
'*****************************************************************************************'

'# 페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'-----------------------------------------------------------------------------------------'

'# 페이지 코드----------------------------------------------------------------------------'
DIM PageCode1, PageCode2, PageCode3, PageCode4
PageCode1 = "00"
PageCode2 = "00"
PageCode3 = "00"
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>
jhgfjhgfjghfjghf
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
	
DIM LinkFunction
DIM ToDay : ToDay = R_YEAR & R_MONTH & R_DAY & R_HOUR & R_MIN
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>


<!-- #include virtual="/INC/Header.asp" -->
<!-- #include virtual="/INC/TopMain.asp" -->

	<!-- Main -->
	<main id="container" class="container">
		<div class="content">
			<div class="slider-for">
				<section class="main-contents">
					<article class="main-style1">
<%
'# 메인 비쥬얼 배너
wQuery = "WHERE BCode = '01' AND DelFlag = 'N' AND StartDT <= '" & ToDay & "' AND EndDT >= '" & ToDay & "' "
sQuery = "ORDER BY ReserveMainFlag DESC, DisplayNum ASC, Idx DESC "


SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_MainBanner_Select_For_Ing"

		.Parameters.Append .CreateParameter("@wQuery",		 adVarchar, adParaminput, 1000	, wQuery)
		.Parameters.Append .CreateParameter("@sQuery",		 adVarchar, adParaminput, 100	, sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
%>
						<div class="main-slider">
							<div class="swiper-container main-swiper">
								<ul class="swiper-wrapper">
<%
		Do Until oRs.EOF
				IF oRs("NewBrowserFlag") = "Y" THEN
						LinkFunction = "openExternal"
				ELSE
						LinkFunction = "APP_GoUrl"
				END IF
%>
									<li class="swiper-slide">
										<a href="javascript:void(0)" onclick="<%=LinkFunction%>('<%=oRs("LinkUrl")%>')" class="listitems">
											<div class="thumbnail">
												<img src="<%=oRs("MobileImage1")%>" alt="<%=REPLACE(oRs("Title"), """", "")%>">
											</div>
										</a>
									</li>
<%
				oRs.MoveNext
		Loop
%>
								</ul>
								<div class="swiper-pagination"></div>
							</div>
						</div>
<%
END IF
oRs.Close
%>
						<div id="tabs" class="tab btn-area" data-use="">
							<ul class="tab-selector main-tab-btn">
								<li data-scode0="01" class="part-2 active"><a href="javascript:void(0);" data-target="tabs-col1">BEST SELLER</a></li>
								<li data-scode0="02" class="part-2"><a href="javascript:void(0);" data-target="tabs-col1">NEW ARRIVALS</a></li>
							</ul>
							<div id="tabs-col1" class="tab-panel active">
								<ul class="main-category-btn">
									<li data-scode1="00" class="sub-part-2 active"><a href="javascript:void(0);">ALL</a></li>
									<li data-scode1="02" class="sub-part-2"><a href="javascript:void(0);">WOMEN</a></li>
									<li data-scode1="01" class="sub-part-2"><a href="javascript:void(0);">MEN</a></li>
									<li data-scode1="03" class="sub-part-2"><a href="javascript:void(0);">KIDS</a></li>
									<li data-scode1="04" class="sub-part-2"><a href="javascript:void(0);">ACC</a></li>
								</ul>

								<div id="BestNArrivalsProductList" class="wrap-item-list"></div>
							</div>
						</div>
					</article>



					<article class="main-style2">

						<div class="best-brands">
<%
'# BEST BRANDS
wQuery = "WHERE A.BCode = '02' AND A.DelFlag = 'N' AND A.StartDT <= '" & ToDay & "' AND A.EndDT >= '" & ToDay & "' "
sQuery = "ORDER BY A.ReserveMainFlag DESC, A.DisplayNum ASC, A.Idx DESC "


SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_MainBanner_Select_Top5_For_BestBrands_Ing"

		.Parameters.Append .CreateParameter("@wQuery",		 adVarchar, adParaminput, 1000	, wQuery)
		.Parameters.Append .CreateParameter("@sQuery",		 adVarchar, adParaminput, 100	, sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

DIM MobileImage1_02
DIM BrandName_02

IF NOT oRs.EOF THEN
		MobileImage1_02	 = oRs("MobileImage1")
		BrandName_02	 = oRs("BrandName")
%>
							<div class="tit-area">
								<p class="section-tit">BEST BRANDS</p>
								<a href="javascript:void(0)" class="tit-badge ViewAll">VIEW ALL</a>
							</div>
							<div class="btn-area">
								<ul class="brand-btn">
<%
		i = 1
		Do Until oRs.EOF	
%>
									<li <%IF i = 1 THEN%>class="active"<%END IF%>><a href="javascript:void(0)" data-num="<%=oRs("Idx")%>" data-brandcode="<%=oRs("BrandCode")%>" data-brandname="<%=oRs("BrandName")%>" data-img="<%=oRs("MobileImage1")%>" class="BestBrands"><%=oRs("BrandName")%></a></li>
<%
				oRs.MoveNext
				i = i + 1
		Loop
%>
								</ul>
							</div>
							<div class="bg-area">
								<img id="BestBrandsVisual" src="<%=MobileImage1_02%>" alt="<%=BrandName_02%>">
							</div>
							<div id="BestBrandsProductList" class="wrap-item-list">
								<!--
								<div class="buttongroup">
									<button type="button" class="button is-expand">
										<span class="icon is-right is-arrow-d2">더보기</span>
									</button>

									<span class="pagination">
										<span class="current">1</span>/<span class="all">7</span>
									</span>
								</div>
								-->
							</div>
<%
END IF
oRs.Close



wQuery = "WHERE BCode = '03' AND DelFlag = 'N' AND StartDT <= '" & ToDay & "' AND EndDT >= '" & ToDay & "' "
sQuery = "ORDER BY ReserveMainFlag DESC, DisplayNum ASC, Idx DESC "


SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_MainBanner_Select_For_Ing"

		.Parameters.Append .CreateParameter("@wQuery",		 adVarchar, adParaminput, 1000	, wQuery)
		.Parameters.Append .CreateParameter("@sQuery",		 adVarchar, adParaminput, 100	, sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

DIM Flag_03 : Flag_03 = "N"

IF NOT oRs.EOF THEN
		Flag_03 = "Y"
%>
							<div class="ad-event">
								<div class="swiper-container evt-slider">
									<ul class="swiper-wrapper">
<%
		Do Until oRs.EOF
				IF oRs("NewBrowserFlag") = "Y" THEN
						LinkFunction = "openExternal"
				ELSE
						LinkFunction = "LinkgoUrl"
				END IF
%>
										<li class="swiper-slide">
											<a href="javascript:void(0)" onclick="<%=LinkFunction%>('<%=oRs("LinkUrl")%>')" class="listitems">
												<div class="thumbnail">
													<img src="<%=oRs("MobileImage1")%>" alt="<%=REPLACE(oRs("Title"), """", "")%>">
												</div>
											</a>
										</li>
<%
				oRs.MoveNext
		Loop
%>
									</ul>
									<div class="swiper-pagination"></div>
								</div>
							</div>
<%
END IF
oRs.Close
%>
						</div>


<%
wQuery = "WHERE BCode = '04' AND DelFlag = 'N' AND StartDT <= '" & ToDay & "' AND EndDT >= '" & ToDay & "' "
sQuery = "ORDER BY ReserveMainFlag DESC, DisplayNum ASC, Idx DESC "


SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_MainBanner_Select_For_MDChoice_Ing"

		.Parameters.Append .CreateParameter("@wQuery",		 adVarchar, adParaminput, 1000	, wQuery)
		.Parameters.Append .CreateParameter("@sQuery",		 adVarchar, adParaminput, 100	, sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing


IF NOT oRs.EOF THEN
%>
						<div class="md-choice">
							<div class="tit-area">
								<p class="section-tit">MD CHOICE</p>
							</div>
							<div class="swiper-container md-swiper">
								<ul class="swiper-wrapper">
<%
		Do Until oRs.EOF
				IF oRs("NewBrowserFlag") = "Y" THEN
						LinkFunction = "openExternal"
				ELSE
						LinkFunction = "APP_GoUrl"
				END IF
%>
									<li class="swiper-slide">
										<a href="javascript:void(0)" onclick="APP_GoUrl('<%=oRs("LinkUrl")%>')">
											<img src="<%=oRs("MobileImage1")%>" alt="<%=REPLACE(oRs("Title"), """", "")%>">
											<p><%=oRs("DisplayTitle1")%><br><%=oRs("DisplayTitle2")%></p>
											<span><%=oRs("SubTitle1")%><br><%=oRs("SubTitle2")%></span>
										</a>
									</li>
<%
				oRs.MoveNext
		Loop	
%>
								</ul>
							</div>
						</div>
<%
END IF
oRs.Close





wQuery = "WHERE BCode = '05' AND DelFlag = 'N' AND StartDT <= '" & ToDay & "' AND EndDT >= '" & ToDay & "' "
sQuery = "ORDER BY ReserveMainFlag DESC, DisplayNum ASC, Idx DESC "


SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_MainBanner_Select_For_Ing"

		.Parameters.Append .CreateParameter("@wQuery",		 adVarchar, adParaminput, 1000	, wQuery)
		.Parameters.Append .CreateParameter("@sQuery",		 adVarchar, adParaminput, 100	, sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing
	
DIM Title_05
DIM MobileImage1_05
DIM DisplayTitle1_05
DIM DisplayTitle2_05
IF NOT oRs.EOF THEN
		Title_05			 = REPLACE(oRs("Title"), """", "")
		MobileImage1_05		 = oRs("MobileImage1")
		DisplayTitle1_05	 = oRs("DisplayTitle1")
		DisplayTitle2_05	 = oRs("DisplayTitle2")
%>
						<div class="hot-item">
							<div class="tit-area">
								<p class="section-tit">지금 뜨는 #상품</p>
							</div>
							<div class="tab-area">
								<img id="MobileImage1_05" src="<%=MobileImage1_05%>" alt="<%=Title_05%>">
								<div class="tab-btn">
<%
		i = 1
		Do Until oRs.EOF	
%>
									<a href="javascript:void(0)" data-num="<%=oRs("Idx")%>" data-title="<%=REPLACE(oRs("TItle"), """", "")%>" data-mobileimage1="<%=oRs("MobileImage1")%>" data-displaytitle1="<%=oRs("DisplayTitle1")%>" data-displaytitle2="<%=oRs("DisplayTitle2")%>" class="NBProduct<%IF i = 1 THEN%> active<%END IF%>">#<%=oRs("SubTitle1")%></a>
<%
				oRs.MoveNext
				i = i + 1
		Loop
%>
								</div>
								<p id="DisplayTitle1_05"><%=DisplayTitle1_05%><br><%=DisplayTitle2_05%></p>
							</div>

							<div id="NowBestProductList"></div>
						</div>
<%
END IF
oRs.Close




wQuery = "WHERE BCode = '06' AND DelFlag = 'N' AND StartDT <= '" & ToDay & "' AND EndDT >= '" & ToDay & "' "
sQuery = "ORDER BY ReserveMainFlag DESC, DisplayNum ASC, Idx DESC "


SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_MainBanner_Select_For_Ing"

		.Parameters.Append .CreateParameter("@wQuery",		 adVarchar, adParaminput, 1000	, wQuery)
		.Parameters.Append .CreateParameter("@sQuery",		 adVarchar, adParaminput, 100	, sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing


REDIM Idx_06(oRs.RecordCount)
REDIM Title_06(oRs.RecordCount)
REDIM MobileImage1_06(oRs.RecordCount)
REDIM MobileImage2_06(oRs.RecordCount)
REDIM DisplayTitle1_06(oRs.RecordCount)
REDIM DisplayTitle2_06(oRs.RecordCount)
REDIM DisplayTitle3_06(oRs.RecordCount)

IF NOT oRs.EOF THEN
		i = 1
		Do Until oRs.EOF
				Idx_06(i)			 = oRs("Idx")
				Title_06(i)			 = REPLACE(oRs("Title"), """", "")
				MobileImage1_06(i)	 = oRs("MobileImage1")
				MobileImage2_06(i)	 = oRs("MobileImage2")
				DisplayTitle1_06(i)	 = oRs("DisplayTitle1")
				DisplayTitle2_06(i)	 = oRs("DisplayTitle2")
				DisplayTitle3_06(i)	 = oRs("DisplayTitle3")

				oRs.MoveNext
				i = i + 1
		Loop
END IF
oRs.Close



IF UBound(Idx_06) > 0 THEN
%>
						<div class="style-people">
							<div class="tit-area">
								<p class="section-tit">STYLE PEOPLE</p>
							</div>
							<div></div>
							<div class="swiper-container style-swiper">
								<ul class="swiper-wrapper">
<%
		FOR i = 1 TO UBound(Idx_06)
%>
									<li class="swiper-slide <%IF i = 1 THEN%>active<%END IF%>">
										<a data-num="<%=Idx_06(i)%>" href="javascript:get_StylePeopleProductList(<%=Idx_06(i)%>);" class="StylePeopleIcon"><img src="<%=MobileImage1_06(i)%>" alt="<%=Title_06(i)%>"></a>
									</li>
<%
		NEXT	
%>
								</ul>
							</div>
							<div class="style-bg" id="StylePeopleProductList">

							</div>

						</div>
<%
END IF







wQuery = "WHERE BCode = '07' AND DelFlag = 'N' AND StartDT <= '" & ToDay & "' AND EndDT >= '" & ToDay & "' "
sQuery = "ORDER BY ReserveMainFlag DESC, DisplayNum ASC, Idx DESC "


SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_MainBanner_Select_For_Ing"

		.Parameters.Append .CreateParameter("@wQuery",		 adVarchar, adParaminput, 1000	, wQuery)
		.Parameters.Append .CreateParameter("@sQuery",		 adVarchar, adParaminput, 100	, sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing


IF NOT oRs.EOF THEN
%>
						<div class="season-now">
							<div class="tit-area">
								<p class="section-tit">SEASON NOW</p>
							</div>
							<ul>
<%
		Do Until oRs.EOF
				IF oRs("NewBrowserFlag") = "Y" THEN
						LinkFunction = "openExternal"
				ELSE
						LinkFunction = "APP_GoUrl"
				END IF
%>
								<li>
									<img src="<%=oRs("PCImage1")%>" alt="<%=REPLACE(oRs("Title"), """", "")%>">
									<div class="txt">
										<p><%=oRs("DisplayTitle1")%></p>
										<a href="javascript:void(0)" onclick="<%=LinkFunction%>('<%=oRs("LinkUrl")%>')">VIEW MORE</a>
									</div>
								</li>
<%
				oRs.MoveNext
		Loop
%>
							</ul>
						</div>
<%
END IF
oRs.Close
%>


<%
wQuery = "WHERE A.DelFlag = 'N' AND C.DCode = 'M' "

SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_Product_Review_Display_Select_By_wQuery"

		.Parameters.Append .CreateParameter("@wQuery",		 adVarchar, adParaminput, 1000	, wQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

If Not oRs.EOF Then
%>					


						<div class="best-review">
							<div class="tit-area">
								<p class="section-tit">BEST REVIEW</p>
							</div>
							<div class="swiper-container review-swiper">
								<ul class="swiper-wrapper">
								<% Do While Not oRs.EOF %>
									<li class="swiper-slide">
										<img src="/Upload/Community/ProductReview/<%=oRs("ThumbNameImage")%>" alt="">
										<div class="txt">
											<strong><%=MaskUserID(oRs("UserID"))%></strong>
											<p><%=ReplaceDetails(oRs("Contents"))%></p>
											<p class="star-score">
												<span class="point val<%=Replace(FormatNumber(oRs("AvgGrade"), 1), ".", "")%>"></span>
												<!-- 평점에 해당하는 값을 닷(.) 제외하고 val40 같은 형식으로 클래스 부여 (3.5점이면 val35) -->
											</p>
											<a class="more" href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>')">
												+<span class="hidden">더보기</span>
											</a>
										</div>
									</li>
								<%
									oRs.MoveNext
								Loop
								%>

								</ul>
								<div class="swiper-pagination"></div>
							</div>
						</div>

<%
End If
oRs.Close
%>





<%
wQuery = "WHERE BCode = '08' AND DelFlag = 'N' AND StartDT <= '" & ToDay & "' AND EndDT >= '" & ToDay & "' "
sQuery = "ORDER BY ReserveMainFlag DESC, DisplayNum ASC, Idx DESC "


SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_MainBanner_Select_For_ShoeMarker_Ing"

		.Parameters.Append .CreateParameter("@wQuery",		 adVarchar, adParaminput, 1000	, wQuery)
		.Parameters.Append .CreateParameter("@sQuery",		 adVarchar, adParaminput, 100	, sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing


IF NOT oRs.EOF AND oRs.RecordCount = 9 THEN
%>
						<div class="shoemarker">
							<div class="tit-area">
								<p class="section-tit">#SHOEMARKER</p>
								<a href="javascript:openExternal('https://www.instagram.com/shoemarker_official/')" class="tit-badge">FOLLOW!</a>
							</div>
							<div class="sns">
<%
		i = 1
		Do Until oRs.EOF
%>
								<a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>')">
									<%IF i = 1 THEN%>
									<img src="<%=oRs("MobileImage2")%>" alt="<%=REPLACE(oRs("Title"), """", "")%>">
									<%ELSE%>
									<img src="<%=oRs("MobileImage1")%>" alt="<%=REPLACE(oRs("Title"), """", "")%>">
									<%END IF%>
									<p class="btn">+<span class="hidden">sns로 이동</span></p>
								</a>
<%
				oRs.MoveNext
				i = i + 1
		Loop
%>
							</div>
						</div>
					</article>
<%
END IF
oRs.Close
%>
				</section>
				<!--
				<section>2</section>
				<section>3</section>
				<section>4</section>
				<section>5</section>
				<section>6</section>
				-->
			</div>
		</div>
	</main>




<%IF Flag_03 ="Y" THEN%>
<script type="text/javascript">
	/*
	$(function () {
		var evtSlide = new Swiper('.evt-slider', {
			slidesPerView: 'auto',
			spaceBetween: 5,
			centeredSlides: true,
			observer: true,
			observeParents: true,
			pagination: {
				el: '.swiper-pagination',
				clickable: true
			},
		});
	});
	*/
</script>
<%END IF%>

<script type="text/javascript">
	$(function () {
		var mainSwiper = new Swiper('.main-swiper', {
			slidesPerView: 1,
			loop: true,
			spaceBetween: 5,
			centeredSlides: true,
            autoplay: {
               delay: 5000,  
            },
            autoplayDisableOnInteraction: true,
			observer: true,
			observeParents: true,
			pagination: {
				el: '.swiper-pagination',
				clickable: true
			},
		});

		/* BEST BRANDS PRODUCT LIST */
		$(".part-2").each(function () {
			if ($(this).hasClass("active")) {
				var sCode0 = $(this).data("scode0");
				var sCode1 = "";
				$(".sub-part-2").each(function () {
					if ($(this).hasClass("active")) {
						sCode1 = $(this).data("scode1");
					}
				});

				if (sCode1 == "") {
					$(".sub-part-2").eq(0).addClass("active");
				}
					
				get_BestNArrivalsProductList(sCode0, sCode1);
			}
		});

		/* BEST SELLER/NEW ARRIVALS TAB CLICK */
		$(".part-2").click(function () {
			$(".part-2").removeClass("active");
			$(this).addClass("active");

			var sCode0 = $(this).data("scode0");
			var sCode1 = "";
			$(".sub-part-2").each(function () {
				if ($(this).hasClass("active")) {
					sCode1 = $(this).data("scode1");
				}
			});

			if (sCode1 == "") {
				$(".sub-part-2").eq(0).addClass("active");
			}
					
			get_BestNArrivalsProductList(sCode0, sCode1);
		});

		/* BEST SELLER/NEW ARRIVALS SUB TAB CLICK */
		$(".sub-part-2").click(function () {
			$(".sub-part-2").removeClass("active");
			$(this).addClass("active");

			var sCode0 = "";
			var sCode1 = $(this).data("scode1");
			$(".part-2").each(function () {
				if ($(this).hasClass("active")) {
					sCode0 = $(this).data("scode0");
				}
			});
					
			if (sCode0 == "") {
				$(".part-2").eq(0).addClass("active");
			}
					
			get_BestNArrivalsProductList(sCode0, sCode1);
		});



		/* BEST BRANDS PRODUCT LIST */
		$(".BestBrands").each(function () {
			if ($(this).parent().hasClass("active")) {
				var num = $(this).data("num");

				get_BestBrandsProductList(num);
			}
		});

		/* BEST BRANDS BRAND CLICK */
		$(".BestBrands").click(function () {
			$(".BestBrands").parent().removeClass("active");
			$(this).parent().addClass("active");
					
			var num = $(this).data("num");
			var brandName = $(this).data("brandname");
			var brandImg = $(this).data("img");

			$("#BestBrandsVisual").attr("src", brandImg);
			$("#BestBrandsVisual").attr("alt", brandName);

			get_BestBrandsProductList(num);
		});
				
		/* BEST BRANDS VIEW ALL CLICK */
		$(".ViewAll").click(function () {
			$(".BestBrands").each(function () {
				if ($(this).parent().hasClass("active")) {
					var brandCode = $(this).data("brandcode");
					//location.href = "/ASP/Product/Brand.asp?SBrandCode=" + brandCode;
					APP_GoUrl("/ASP/Product/Brand.asp?SBrandCode=" + brandCode);
				}
			});
		});


		
		/* 지금뜨는상품 LIST */
		$(".NBProduct").each(function () {
			if ($(this).hasClass("active")) {
				var num = $(this).data("num");

				get_NowBestProductList(num);
			}
		});

		/* 지금뜨는상품태그 CLICK */
		$(".NBProduct").click(function () {
			$(".NBProduct").removeClass("active");
			$(this).addClass("active");
					
			var num = $(this).data("num");
			var title = $(this).data("title");
			var displayTitle1 = $(this).data("displaytitle1");
			var displayTitle2 = $(this).data("displaytitle2");
			var mobileImage1 = $(this).data("mobileimage1");

			$("#MobileImage1_05").attr("src", mobileImage1);
			$("#MobileImage1_05").attr("alt", title);
			$("#DisplayTitle1_05").html(displayTitle1 + "<br />" + displayTitle2);

			get_NowBestProductList(num);
		});

		
		/* BEST BRANDS PRODUCT LIST */
		$(".StylePeopleIcon").each(function () {
			if ($(this).parent().hasClass("active")) {
				var num = $(this).data("num");

				get_StylePeopleProductList(num);
			}
		});

		/* STYLE PEOPLE CLOCK */
		$(".StylePeopleIcon").click(function () {
			$(".StylePeopleIcon").parent().removeClass("active");
			$(this).parent().addClass("active");

			/*
			var num = $(this).data("num");
			var title = $(this).data("title");
			var mobileImage2 = $(this).data("mobileimage2");
			var displayTitle1 = $(this).data("displaytitle1");
			var displayTitle2 = $(this).data("displaytitle2");
			var displayTitle3 = $(this).data("displaytitle3");
			
			if (displayTitle3 != "") {
				displayTitle1 = "<a href=\"javascript:openExternal('" + displayTitle3 + "')\">" + displayTitle1 + "</a>";
			}
			
			$("#StylePeopleVisual").attr("src", mobileImage2);
			$("#StylePeopleVisual").attr("alt", title);
			
			$("#StylePeopleTitle1").html(displayTitle1);
			$("#StylePeopleTitle2").html(displayTitle2);
			
			get_StylePeopleProductList(num);
			*/
		});
	});
</script>

<!-- #include virtual="/INC/Footer.asp" -->
<!-- #Include virtual="/INC/SlidePopup.asp" -->
<!-- #Include virtual="/INC/TimeSale.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>


<% 'ShareGate 에서 넘어 왔을때 처리 하는 부분 %>
<%
	Dim Share_ProductCode
	Share_ProductCode = sqlFilter(Request("ProductCode"))
	
	If Trim(Share_ProductCode) <> "" Then
%>
<script type="text/javascript">
	function goProductDetail() {
		APP_GoUrl('/ASP/Product/ProductDetail.asp?ProductCode=<%=Share_ProductCode%>');
	}

	goProductDetail();
</script>
<%
	End If	
%>
<% 'ShareGate 에서 넘어 왔을때 처리 하는 부분 %>

<% 'PushGate 에서 넘어 왔을때 처리 하는 부분 %>
<%
	If Request("GoUrl") <> "" Then	
%>
<script type="text/javascript">
	function PushGoUrl() {
		APP_GoUrl('<%=Request("GoUrl")%>');
	}

	PushGoUrl();
</script>
<%
	End If
%>
<% 'PushGate 에서 넘어 왔을때 처리 하는 부분 %>