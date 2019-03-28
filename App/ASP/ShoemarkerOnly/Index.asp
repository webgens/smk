<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'Index.asp - ShoemarkerOnly
'Date		: 2019.01.07
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

DIM MainBanner		 : MainBanner		 = ""
DIM HOT_Banner		 : HOT_Banner		 = ""
DIM WHEVER_Banner	 : WHEVER_Banner	 = ""
DIM MARKERS_Banner	 : MARKERS_Banner	 = ""
DIM KIDS_Banner		 : KIDS_Banner		 = ""
DIM PCode			 : PCode			 = ""
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성


SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_ShoemarkerOnly_Select"
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		MainBanner		 = oRs("MobileMainBanner")
		HOT_Banner		 = oRs("HOT_MobileBanner")
		WHEVER_Banner	 = oRs("WHEVER_MobileBanner")
		MARKERS_Banner	 = oRs("MARKERS_MobileBanner")
		KIDS_Banner		 = oRs("KIDS_MobileBanner")
END IF
oRs.Close
%>

<!-- #include virtual="/INC/Header.asp" -->
<!-- #include virtual="/INC/Top_ShoemarkerOnly.asp" -->

    <!-- Main -->
    <main id="container" class="container">
        <div class="content">
            <section class="only-contents main-style2">
                <article class="only-main">
                    <div class="main-area">
                        <img src="<%=MainBanner%>" alt="ShoemarkerOnly">
                    </div>
					<%
					wQuery = "WHERE BCode = '15' AND DelFlag = 'N' AND StartDT <= '" & R_YEAR & R_MONTH & R_DAY & R_HOUR & R_MIN & "' AND EndDT >= '" & R_YEAR & R_MONTH & R_DAY & R_HOUR & R_MIN & "' "
					sQuery = "ORDER BY ReserveMainFlag DESC, DisplayNum ASC, Idx DESC "
					SET oCmd = Server.CreateObject("ADODB.Command")
					WITH oCmd
							.ActiveConnection	 = oConn
							.CommandType		 = adCmdStoredProc
							.CommandText		 = "USP_Admin_EShop_MainBanner_Select_For_Ing"
							.Parameters.Append .CreateParameter("@wQuery",		 adVarchar, adParaminput, 1000	, wQuery)
							.Parameters.Append .CreateParameter("@sQuery",		 adVarchar, adParaminput, 1000	, sQuery)
					END WITH
					oRs.CursorLocation = adUseClient
					oRs.Open oCmd, , adOpenStatic, adLockReadOnly
					SET oCmd = Nothing
					
					Dim SBannerCount
					SBannerCount = oRs.RecordCount

					If Not oRs.EOF Then
					%>
                    <div class="ad-event">
                        <div class="swiper-container evt-slider">
                            <ul class="swiper-wrapper">
							<%
							Do While Not oRs.EOF	
							%>
                                <li class="swiper-slide">
                                    <a href="javascript:void(0)" onclick="LinkgoUrl('<%=oRs("LinkUrl")%>')" class="listitems">
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
					End If
					oRs.Close	
					%>
                </article>
                <article class="hot">
                    <div class="tit-area">
                        <p class="section-tit">HOT</p>
                        <a href="/ASP/ShoemarkerOnly/HOT.asp" class="tit-badge">VIEW ALL</a>
                    </div>
                    <div class="story">
                        <img src="<%=HOT_Banner%>" alt="HOT">
                    </div>
                    <div class="swiper-container ranking-slider">
                        <ol class="swiper-wrapper">
						<%
						wQuery = "WHERE A.SaleState = 'Y' AND D.PCode = 'H' AND D.MainFlag = 'Y' "
						SET oCmd = Server.CreateObject("ADODB.Command")
						WITH oCmd
								.ActiveConnection	 = oConn
								.CommandType		 = adCmdStoredProc
								.CommandText		 = "USP_Mobile_EShop_ShoemarkerOnly_Product_Select_By_wQuery"
								.Parameters.Append .CreateParameter("@wQuery",		 adVarchar, adParaminput, 1000	, wQuery)
						END WITH
						oRs.CursorLocation = adUseClient
						oRs.Open oCmd, , adOpenStatic, adLockReadOnly
						SET oCmd = Nothing
					
						If Not oRs.EOF Then
							Do While Not oRs.EOF
						%>
                            <li class="swiper-slide">
                                <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>')" class="listitems">
                                    <div class="thumbnail"><img src="<%=oRs("ImageUrl")%>" alt="<%=oRs("ProductName")%>"></div>
                                    <p class="brand-name"><%=oRs("BrandName")%></p>
                                    <h1 class="product-name pname"><%=oRs("ProductName")%></h1>
                                    <p class="price"><strong><%=FormatNumber(oRs("SalePrice"), 0)%></strong>원</p>
                                </a>
                            </li>
						<%
								oRs.MoveNext
							Loop
						End If
						oRs.Close
						%>
                        </ol>
                    </div>
                </article>

                <article class="whever">
                    <div class="tit-area">
                        <p class="section-tit">WHEVER</p>
                        <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Product/Brand.asp?SBrandCode=WV')" class="tit-badge">VIEW ALL</a>
                    </div>
                    <div class="story">
                        <img src="<%=WHEVER_Banner%>" alt="WHEVER">
                        <div class="txt">
                            <p class="story-tit">BRAND STORY</p>
                            <p class="story-explain">WHENEVER+WHEREVER + WHATEVER 의 합성어로 [웨:버]로 읽으며, 언제든 어디에서든 편안하게 스타일링 할 수 있는 다양한 라이프스타일 슈즈를 선보입니다.</p>
                        </div>
                    </div>
                    <div class="swiper-container ranking-slider">
                        <ol class="swiper-wrapper">
						<%
						wQuery = "WHERE A.SaleState = 'Y' AND D.PCode = 'W' AND D.MainFlag = 'Y' "
						SET oCmd = Server.CreateObject("ADODB.Command")
						WITH oCmd
								.ActiveConnection	 = oConn
								.CommandType		 = adCmdStoredProc
								.CommandText		 = "USP_Mobile_EShop_ShoemarkerOnly_Product_Select_By_wQuery"
								.Parameters.Append .CreateParameter("@wQuery",		 adVarchar, adParaminput, 1000	, wQuery)
						END WITH
						oRs.CursorLocation = adUseClient
						oRs.Open oCmd, , adOpenStatic, adLockReadOnly
						SET oCmd = Nothing
					
						If Not oRs.EOF Then
							Do While Not oRs.EOF
						%>
                            <li class="swiper-slide">
                                <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>')" class="listitems">
                                    <div class="thumbnail"><img src="<%=oRs("ImageUrl")%>" alt="<%=oRs("ProductName")%>"></div>
                                    <p class="brand-name"><%=oRs("BrandName")%></p>
                                    <h1 class="product-name pname"><%=oRs("ProductName")%></h1>
                                    <p class="price"><strong><%=FormatNumber(oRs("SalePrice"), 0)%></strong>원</p>
                                </a>
                            </li>
						<%
								oRs.MoveNext
							Loop
						End If
						oRs.Close
						%>
                        </ol>
                    </div>
                </article>
                <article class="markers">
                    <div class="tit-area">
                        <p class="section-tit">MARK:ERS</p>
                        <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Product/Brand.asp?SBrandCode=MR')" class="tit-badge">VIEW ALL</a>
                    </div>
                    <div class="story">
                        <img src="<%=MARKERS_Banner%>" alt="MAKERS">
                        <div class="txt">
                            <p class="story-tit">BRAND STORY</p>
                            <p class="story-explain">마커스는 '무언가를 나타내는 표시'라는 의미에서 시작되어 나의 존재를 알리고 즐겁고 유쾌한 라이프스타일을 추구하는 사람들을 위한 다양한 아이템을 제안합니다. 오래 신어도 발이 편안하고, 일상에서의 쾌적함을 유지시켜 줍니다.</p>
                        </div>
                    </div>
                    <div class="swiper-container ranking-slider">
                        <ol class="swiper-wrapper">
						<%
						wQuery = "WHERE A.SaleState = 'Y' AND D.PCode = 'M' AND D.MainFlag = 'Y' "
						SET oCmd = Server.CreateObject("ADODB.Command")
						WITH oCmd
								.ActiveConnection	 = oConn
								.CommandType		 = adCmdStoredProc
								.CommandText		 = "USP_Mobile_EShop_ShoemarkerOnly_Product_Select_By_wQuery"
								.Parameters.Append .CreateParameter("@wQuery",		 adVarchar, adParaminput, 1000	, wQuery)
						END WITH
						oRs.CursorLocation = adUseClient
						oRs.Open oCmd, , adOpenStatic, adLockReadOnly
						SET oCmd = Nothing
					
						If Not oRs.EOF Then
							Do While Not oRs.EOF
						%>
                            <li class="swiper-slide">
                                <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>')" class="listitems">
                                    <div class="thumbnail"><img src="<%=oRs("ImageUrl")%>" alt="<%=oRs("ProductName")%>"></div>
                                    <p class="brand-name"><%=oRs("BrandName")%></p>
                                    <h1 class="product-name pname"><%=oRs("ProductName")%></h1>
                                    <p class="price"><strong><%=FormatNumber(oRs("SalePrice"), 0)%></strong>원</p>
                                </a>
                            </li>
						<%
								oRs.MoveNext
							Loop
						End If
						oRs.Close
						%>
                        </ol>
                    </div>
                </article>
                <article class="kids">
                    <div class="tit-area">
                        <p class="section-tit">KIDS</p>
                        <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Product/Brand.asp?SearchType=C&SBrandCode=WV&SCode1=03')" class="tit-badge">VIEW ALL</a>
                    </div>
                    <div class="story">
                        <img src="<%=KIDS_Banner%>" alt="KIDS">
                        <div class="txt">
                            <p class="story-tit">BRAND STORY</p>
                            <p class="story-explain">WHENEVER+WHEREVER + WHATEVER 의 합성어로 [웨:버]로 읽으며, 언제든 어디에서든 편안하게 스타일링 할 수 있는 다양한 라이프스타일 슈즈를 선보입니다.</p>
                        </div>
                    </div>
                    <div class="swiper-container ranking-slider">
                        <ol class="swiper-wrapper">
						<%
						wQuery = "WHERE A.SaleState = 'Y' AND D.PCode = 'K' AND D.MainFlag = 'Y' "
						SET oCmd = Server.CreateObject("ADODB.Command")
						WITH oCmd
								.ActiveConnection	 = oConn
								.CommandType		 = adCmdStoredProc
								.CommandText		 = "USP_Mobile_EShop_ShoemarkerOnly_Product_Select_By_wQuery"
								.Parameters.Append .CreateParameter("@wQuery",		 adVarchar, adParaminput, 1000	, wQuery)
						END WITH
						oRs.CursorLocation = adUseClient
						oRs.Open oCmd, , adOpenStatic, adLockReadOnly
						SET oCmd = Nothing
					
						If Not oRs.EOF Then
							Do While Not oRs.EOF
						%>
                            <li class="swiper-slide">
                                <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>')" class="listitems">
                                    <div class="thumbnail"><img src="<%=oRs("ImageUrl")%>" alt="<%=oRs("ProductName")%>"></div>
                                    <p class="brand-name"><%=oRs("BrandName")%></p>
                                    <h1 class="product-name pname"><%=oRs("ProductName")%></h1>
                                    <p class="price"><strong><%=FormatNumber(oRs("SalePrice"), 0)%></strong>원</p>
                                </a>
                            </li>
						<%
								oRs.MoveNext
							Loop
						End If
						oRs.Close
						%>
                        </ol>
                    </div>
                </article>
            </section>
		</div>
    </main>



<!-- #include virtual="/INC/Footer.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>