<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'ProductList.asp - 상품리스트
'Date		: 2019.01.05
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
PageCode1 = "PC"
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

DIM Page
DIM PageSize : PageSize = 20
DIM RecCnt
DIM PageCnt

DIM SearchType
DIM SCode1
DIM SCode2
DIM SCode3
DIM SBrandCode
DIM SSizeCD
DIM SPrice
DIM EPrice
DIM SColorCode
DIM SPickupFlag
DIM SFreeFlag
DIM SReserveFlag
Dim SSort
Dim SortText

DIM CategoryName1
DIM CategoryName
Dim Top_CategoryName

DIM ImageUrl
DIM LinkFunction
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

Page			 = sqlFilter(Request("Page"))
SearchType		 = sqlFilter(Request("SearchType"))
SCode1			 = sqlFilter(Request("SCode1"))
SCode2			 = sqlFilter(Request("SCode2"))
SCode3			 = sqlFilter(Request("SCode3"))
SBrandCode		 = sqlFilter(Request("SBrandCode"))
SSizeCD			 = sqlFilter(Request("SSizeCD"))
SPrice			 = sqlFilter(Request("SPrice"))
EPrice			 = sqlFilter(Request("EPrice"))
SColorCode		 = sqlFilter(Request("SColorCode"))
SPickupFlag		 = sqlFilter(Request("SPickupFlag"))
SFreeFlag		 = sqlFilter(Request("SFreeFlag"))
SReserveFlag	 = sqlFilter(Request("SReserveFlag"))
SSort			 = sqlFilter(Request("SSort"))

IF Page			 = "" THEN Page			 = 1
IF SearchType	 = "" THEN SearchType	 = "S"
IF SCode1		 = "" THEN SCode1		 = "01"
IF SPrice		 = "" THEN SPrice		 = 0
IF EPrice		 = "" THEN EPrice		 = 30
If SSort		 = "" THEN SSort		 = "1"

Select Case SSort
	Case "1" : SortText = "신상품순"
	Case "2" : SortText = "인기순"
	Case "3" : SortText = "할인률순"
	Case "4" : SortText = "낮은가격순"
	Case "5" : SortText = "높은가격순"
End Select

SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성

IF SCode1 <> "" THEN
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Category1_Select_By_CategoryCode1"

				.Parameters.Append .CreateParameter("@CategoryCode1", adChar, adParamInput, 2, SCode1)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		SET oCmd = Nothing

		IF NOT oRs.EOF THEN
				CategoryName1 = oRs("CategoryName1")
				CategoryName = oRs("CategoryName1")
		ELSE
				oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Call AlertMessage("없는 분류 정보 입니다.", "/")
				Response.End
		END IF
		oRs.Close
END IF
IF SCode2 <> "" THEN
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Category2_Select_By_CategoryCode2"

				.Parameters.Append .CreateParameter("@CategoryCode1", adChar, adParamInput, 2, SCode1)
				.Parameters.Append .CreateParameter("@CategoryCode2", adChar, adParamInput, 2, SCode2)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		SET oCmd = Nothing

		IF NOT oRs.EOF THEN
				CategoryName = CategoryName & " > " & oRs("CategoryName2")
		ELSE
				oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Call AlertMessage("없는 분류 정보 입니다.", "history.back();")
				Response.End
		END IF
		oRs.Close
END IF

IF SCode3 <> "" THEN
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Category3_Select_By_CategoryCode3"

				.Parameters.Append .CreateParameter("@CategoryCode1", adChar, adParamInput, 2, SCode1)
				.Parameters.Append .CreateParameter("@CategoryCode2", adChar, adParamInput, 2, SCode2)
				.Parameters.Append .CreateParameter("@CategoryCode3", adChar, adParamInput, 2, SCode3)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		SET oCmd = Nothing

		IF NOT oRs.EOF THEN
				'CategoryName = CategoryName & " > " & oRs("CategoryName3")
				Top_CategoryName = oRs("CategoryName3")
		ELSE
				oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Call AlertMessage("없는 분류 정보 입니다.", "history.back();")
				Response.End
		END IF
		oRs.Close
END IF
%>

<!-- #include virtual="/INC/Header.asp" -->
	<style type="text/css">
		.gnb + .container .content { padding-top: 125px; }
	</style>
<!-- #include virtual="/INC/Top_ProductList.asp" -->

    <main id="container" class="container">
        <div class="content">
            <section class="wrap-item-list">
                <div class="item-bg">
					<% If SCode1 = "01" Then %>
						<img src="/images/img/mo_men.jpg" alt="<%=CategoryName1%>">
					<% ElseIf SCode1 = "02" Then %>
						<img src="/images/img/mo_women.jpg" alt="<%=CategoryName1%>">
					<% ElseIf SCode1 = "03" Then %>
						<img src="/images/img/mo_kids.jpg" alt="<%=CategoryName1%>">
					<% ElseIf SCode1 = "04" Then %>
						<img src="/images/img/mo_acc.jpg" alt="<%=CategoryName1%>">
					<% End If %>
                    <p><%=CategoryName%></p>
                    <span>CATEGORY</span>
                </div>

				<% If SCode2 = "" And SCode3 = "" Then %>

                <div class="item-ranking">
                    <div id="tabs" class="tab" data-use="">
                        <ul class="tab-selector">
                            <li class="part-3"><a href="javascript:;" data-target="tabs-col1" class="sell">가장 많이 팔린</a></li>
                            <li class="part-3"><a href="javascript:;" data-target="tabs-col2" class="review">가장 후기 많은</a></li>
                            <li class="part-3"><a href="javascript:;" data-target="tabs-col3" class="review">가장 많이 찜한</a></li>
                        </ul>
                        <div id="tabs-col1" class="tab-panel">
                            <div class="swiper-container ranking-slider">
                                <ol class="swiper-wrapper">
									<%
									wQuery = "WHERE D.CategoryCode1 = '" & SCode1 & "' AND D.TypeCode = 'A' "
									sQuery = "ORDER BY D.DisplayNum ASC "
									SET oCmd = Server.CreateObject("ADODB.Command")
									WITH oCmd
											.ActiveConnection	 = oConn
											.CommandType		 = adCmdStoredProc
											.CommandText		 = "USP_Mobile_EShop_Product_Category_Top3_Select_By_wQuery"
			
											.Parameters.Append .CreateParameter("@wQuery", adVarChar, adParamInput, 1000, wQuery)
											.Parameters.Append .CreateParameter("@sQuery", adVarChar, adParamInput, 100, sQuery)
									END WITH
									oRs.CursorLocation = adUseClient
									oRs.Open oCmd, , adOpenStatic, adLockReadOnly
									SET oCmd = Nothing

									i = 1
									Do While Not oRs.EOF	
									%>
                                    <li class="swiper-slide">
                                        <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>')" class="listitems">
                                            <p class="ranking"><%=i%></p>
                                            <div class="thumbnail"><img src="<%=oRs("ImageUrl")%>" alt="상품 대표 이미지"></div>
                                            <p class="brand-name"><%=oRs("BrandName")%></p>
                                            <h1 class="product-name pname"><%=oRs("ProductName")%></h1>
                                            <p class="price"><strong><%=FormatNumber(oRs("SalePrice"), 0)%></strong>원</p>
                                        </a>
                                    </li>
									<% 
										i = i + 1
										oRs.MoveNext
									Loop
									oRs.Close
									%>
                                </ol>
                            </div>
                        </div>
                        <div id="tabs-col2" class="tab-panel">
                            <div class="swiper-container ranking-slider">
                                <ol class="swiper-wrapper">
									<%
									wQuery = "WHERE D.CategoryCode1 = '" & SCode1 & "' AND D.TypeCode = 'B' "
									sQuery = "ORDER BY D.DisplayNum ASC "
									SET oCmd = Server.CreateObject("ADODB.Command")
									WITH oCmd
											.ActiveConnection	 = oConn
											.CommandType		 = adCmdStoredProc
											.CommandText		 = "USP_Mobile_EShop_Product_Category_Top3_Select_By_wQuery"
			
											.Parameters.Append .CreateParameter("@wQuery", adVarChar, adParamInput, 1000, wQuery)
											.Parameters.Append .CreateParameter("@sQuery", adVarChar, adParamInput, 100, sQuery)
									END WITH
									oRs.CursorLocation = adUseClient
									oRs.Open oCmd, , adOpenStatic, adLockReadOnly
									SET oCmd = Nothing

									i = 1
									Do While Not oRs.EOF	
									%>
                                    <li class="swiper-slide">
                                        <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>')" class="listitems">
                                            <p class="ranking"><%=i%></p>
                                            <div class="thumbnail"><img src="<%=oRs("ImageUrl")%>" alt="상품 대표 이미지"></div>
                                            <p class="brand-name"><%=oRs("BrandName")%></p>
                                            <h1 class="product-name pname"><%=oRs("ProductName")%></h1>
                                            <p class="price"><strong><%=FormatNumber(oRs("SalePrice"), 0)%></strong>원</p>
                                        </a>
                                    </li>
									<% 
										i = i + 1
										oRs.MoveNext
									Loop
									oRs.Close
									%>
                                </ol>
                            </div>
                        </div>
                        <div id="tabs-col3" class="tab-panel">
                            <div class="swiper-container ranking-slider">
                                <ol class="swiper-wrapper">
									<%
									wQuery = "WHERE D.CategoryCode1 = '" & SCode1 & "' AND D.TypeCode = 'C' "
									sQuery = "ORDER BY D.DisplayNum ASC "
									SET oCmd = Server.CreateObject("ADODB.Command")
									WITH oCmd
											.ActiveConnection	 = oConn
											.CommandType		 = adCmdStoredProc
											.CommandText		 = "USP_Mobile_EShop_Product_Category_Top3_Select_By_wQuery"
			
											.Parameters.Append .CreateParameter("@wQuery", adVarChar, adParamInput, 1000, wQuery)
											.Parameters.Append .CreateParameter("@sQuery", adVarChar, adParamInput, 100, sQuery)
									END WITH
									oRs.CursorLocation = adUseClient
									oRs.Open oCmd, , adOpenStatic, adLockReadOnly
									SET oCmd = Nothing

									i = 1
									Do While Not oRs.EOF	
									%>
                                    <li class="swiper-slide">
                                        <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>')" class="listitems">
                                            <p class="ranking"><%=i%></p>
                                            <div class="thumbnail"><img src="<%=oRs("ImageUrl")%>" alt="<%=oRs("ProductName")%>"></div>
                                            <p class="brand-name"><%=oRs("BrandName")%></p>
                                            <h1 class="product-name pname"><%=oRs("ProductName")%></h1>
                                            <p class="price"><strong><%=FormatNumber(oRs("SalePrice"), 0)%></strong>원</p>
                                        </a>
                                    </li>
									<% 
										i = i + 1
										oRs.MoveNext
									Loop
									oRs.Close
									%>
                                </ol>
                            </div>
                        </div>
                    </div>

                </div>

				<% Else %>

				<br />

				<% End If %>

				<%
				wQuery = "WHERE BCode = '12' AND DelFlag = 'N' AND StartDT <= '" & R_YEAR & R_MONTH & R_DAY & R_HOUR & R_MIN & "' AND EndDT >= '" & R_YEAR & R_MONTH & R_DAY & R_HOUR & R_MIN & "' "
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

				If Not oRs.EOF Then
				%>
                <div class="ad-event">
                    <div class="swiper-container evt-slider">
                        <ul class="swiper-wrapper">
						<%
						Do While Not oRs.EOF
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
				End If
				oRs.Close	
				%>


                <div class="sale-item">
                    <div class="item-list">
                        <ul class="listview" id="ProductList">
							
                        </ul>
                    </div>
                </div>

                <div class="buttongroup" id="morebtn">
                    <button type="button" class="button is-expand" onclick="getNextPage();">
						<span class="icon is-right is-arrow-d2">더보기</span>
					</button>

                    <span class="pagination">
						<span class="current" id="nowpage"></span>/<span class="all" id="totalpage"></span>
                    </span>
                </div>
            </section>
    
	   </div>

    </main>

	<form name="form" method="get">
		<input type="hidden" name="SCode1" value="<%=SCode1%>" />
		<input type="hidden" name="SCode2" value="<%=SCode2%>" />
		<input type="hidden" name="SCode3" value="<%=SCode3%>" />
		<input type="hidden" name="SBrandCode" value="<%=SBrandCode%>" />
		<input type="hidden" name="SSizeCD" value="<%=SSizeCD%>" />
		<input type="hidden" name="SPrice" value="<%=SPrice%>" />
		<input type="hidden" name="EPrice" value="<%=EPrice%>" />
		<input type="hidden" name="SColorCode" value="<%=SColorCode%>" />
		<input type="hidden" name="SPickupFlag" value="<%=SPickupFlag%>" />
		<input type="hidden" name="SFreeFlag" value="<%=SFreeFlag%>" />
		<input type="hidden" name="SReserveFlag" value="<%=SReserveFlag%>" />
		<input type="hidden" name="SSort" value="<%=SSort%>" />
		<input type="hidden" name="Page" />
		<input type="hidden" name="ISTopN" />
	</form>

	<script type="text/javascript">
		function get_ProductList(page) {
			$("form[name='form'] > input[name='Page']").val(page);

			var arrHash = "";
			if (document.location.hash.replace("#", "") == "")						//해시값이 없을경우
			{
				$("form[name='form'] > input[name='ISTopN']").val("N");		// history.back() 으로 올경우 전체 페이지 다시가져오기
			}
			else																	//해시값이 있을경우
			{
				$("form[name='form'] > input[name='ISTopN']").val("Y");		// history.back() 으로 올경우 전체 페이지 다시가져오기
				arrHash = document.location.hash.replace("#", "").split("|");
			}

			$.ajax({
				url			 : '/ASP/Product/Ajax/ProductList.asp',
				data		 : $("form[name='form']").serialize(),
				async		 : false,
				type		 : 'get',
				dataType	 : 'html',
				success		 : function (data) {
								if (data == "") {
									$("form[name='form'] > input[name='Page']").val(page - 1);
									return;
								}

								arrData	 = data.split("|||||");
								Data	 = arrData[0];
								RecCnt	 = arrData[1];
								PageCnt	 = arrData[2];

								$("#nowpage").html(page);
								$("#totalpage").html(PageCnt);

								$("#morebtn").show();
								if (parseInt(page) >= parseInt(PageCnt)) {
									$("#morebtn").hide();
								}

								// 목록 로딩시키기
								if (arrHash != "") {
									$("#ProductList").html(Data);
									$(window).scrollTop(arrHash[1]);
									document.location.hash = 0;
								} else {
									if (page == 1) {
										$("#ProductList").html(Data);
									} else {
										$("#ProductList").append(Data);
									}
								}
				},
				error		 : function (data) {
								openAlertLayer("alert", "상품 리스트를 가져오는 도중 오류가 발생하였습니다.", "closePop('alertPop', '');", "");
				}
			});
		}

		function pushHash() {
			document.location.hash = $("form[name='form'] > input[name='Page']").val() + "|" + $(window).scrollTop();
		}

		function getNextPage() {
			var page = document.form.Page.value;
			get_ProductList(parseInt(page) + 1);
		}

		//하위분류 레이어
		function SubCategoryLayerOpen(scode1, scode2, scode3) {
			$.ajax({
				url: '/ASP/Product/Ajax/ProductCategoryList.asp',
				data: "SCode1=" + scode1 + "&SCode2=" + scode2 + "&SCode3=" + scode3,
				async: false,
				type: 'get',
				dataType: 'html',
				success: function (data, textStatus, jqXHR) {

					$("#subcategorylist").html(data);
					$("#SubCateogryPop").show();
				},
				error: function (data, textStatus, jqXHR) {
					//alert(jqXHR);
					//alert(data.responseText);
					alert("리스트를 가져오는 도중 오류가 발생하였습니다.");
				}
			});
		}

		//정렬 레이어
		function OrderByLayerOpen() {
			$("#SortPop").show();
		}

		//정렬 선택
		function OrderBySelect(sort) {
			document.form.SSort.value = sort;
			document.form.Page.value = 1;
			document.form.ISTopN.value = "";
			document.form.action = "/ASP/Product/ProductList.asp";
			document.form.submit();
		}

		//스마트서치
		function SmartSearchLayerOpen() {
			$("#smartsearchPop").show();
		}
	</script>

	<script type="text/javascript">
		$(document).ready(function () {
			// history.back() 시 카테고리로 다시 페이지로딩
			if (document.location.hash) {
				var arrHash = document.location.hash.replace("#", "").split("|")
				if (arrHash.length == 2) {
					document.form.Page.value = arrHash[0];
					get_ProductList(arrHash[0]);
				} else {
					get_ProductList(1);
				}
			} else {
				get_ProductList(1);
			}
		});
	</script>

<!-- #include virtual="/INC/Footer.asp" -->

    <section class="wrap-pop" id="SubCateogryPop">
        <div class="area-dim"></div>

        <div class="area-pop">
            <div class="alert">
                <div class="tit-pop">
                    <p class="tit"><%=CategoryName%></p>
                    <button type="button" class="btn-hide-pop" onclick="$('#SubCateogryPop').hide();">닫기</button>
                </div>
                <div class="container-pop">
                    <div class="contents">
                        <div class="pop-category" id="subcategorylist">
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </section>

    <section class="wrap-pop" id="SortPop">
        <div class="area-dim"></div>
        <div class="area-pop">
            <div class="alert">
                <div class="tit-pop">
                    <p class="tit">정렬</p>
                    <button type="button" class="btn-hide-pop" onclick="$('#SortPop').hide();">닫기</button>
                </div>
                <div class="container-pop">
                    <div class="contents">
                        <div class="pop-category">
                            <button type="button" onclick="OrderBySelect(1);" <% If SSort = "1" Then %>class="on"<% End If %>><span>신상품순</span></button>
                            <button type="button" onclick="OrderBySelect(2);" <% If SSort = "2" Then %>class="on"<% End If %>><span>인기순</span></button>
                            <button type="button" onclick="OrderBySelect(3);" <% If SSort = "3" Then %>class="on"<% End If %>><span>할인율순</span></button>
                            <button type="button" onclick="OrderBySelect(4);" <% If SSort = "4" Then %>class="on"<% End If %>><span>낮은가격순</span></button>
                            <button type="button" onclick="OrderBySelect(5);" <% If SSort = "5" Then %>class="on"<% End If %>><span>높은가격순</span></button>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </section>

    <section class="wrap-pop" id="smartsearchPop">
        <div class="area-dim"></div>

        <div class="area-pop">
            <div class="full">
                <div class="tit-pop">
                    <p class="tit">스마트 검색</p>
                    <button type="button" class="btn-hide-pop" onclick="$('#smartsearchPop').hide();">닫기</button>
                </div>
                <div class="container-pop">
                    <div class="contents">
						<form name="form1" method="get" action="<%=Request.ServerVariables("PATH_INFO")%>">
						<input type="hidden" name="SCode1" value="<%=SCode1%>" />
						<input type="hidden" name="SCode2" value="<%=SCode2%>" />
						<input type="hidden" name="SCode3" value="<%=SCode3%>" />
						<input type="hidden" name="SSort" value="<%=SSort%>" />
                        <!-- HTML 수정 18-11-30 --->
                        <div class="smart-search">
                            <div class="pop-accordion-selector">
                                <button type="button" data-target="smartSearch_brand"><p class="smart-tit">브랜드</p></button>
                            </div>
                            <div id="smartSearch_brand" class="smart-area pop-brand">
                                <div class="pop-accordion-panel">
<%
'# 브랜드 리스트
wQuery = "WHERE A.SaleState = 'Y' "
IF SCode1 <> "" THEN
		wQuery = wQuery & "AND A.ProductCode IN ( "
		wQuery = wQuery & "						SELECT	 DISTINCT ProductCode "
		wQuery = wQuery & "						FROM	 UVW_EShop_Product_Category "
		wQuery = wQuery & "						WHERE	 CategoryCode1 = '" & SCode1 & "' "
		IF SCode2 <> "" THEN
				wQuery = wQuery & "						 AND CategoryCode2 = '" & SCode2 & "' "
		END IF
		IF SCode3 <> "" THEN
				wQuery = wQuery & "						 AND CategoryCode3 = '" & SCode3 & "' "
		END IF
		wQuery = wQuery & ") "
END IF

SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Product_Select_For_Brand_At_ProductList"

		.Parameters.Append .CreateParameter("@WQUERY", adVarChar, adParamInput, 1000, wQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		x = 1
		Do Until oRs.EOF
%>
									<span class="check-style"><input type="checkbox" id="select-<%=x%>" name="SBrandCode" value="|<%=oRs("BrandCode")%>|" <%IF INSTR(SBrandCode, "|"&oRs("BrandCode")&"|") THEN%>checked<%END IF%>><label for="select-<%=x%>"><span><%=oRs("BrandName")%></span></label></span>                                            
<%
				oRs.MoveNext
				x = x + 1
		Loop
END IF
oRs.Close
%>
                                </div>
                            </div>
                            <p class="smart-tit">사이즈</p>
                            <div class="smart-area pop-size">
<%
'# 사이즈 리스트
wQuery = "WHERE A.SaleState = 'Y' "
IF SCode1 <> "" THEN
		wQuery = wQuery & "AND A.ProductCode IN ( "
		wQuery = wQuery & "						SELECT	 DISTINCT ProductCode "
		wQuery = wQuery & "						FROM	 UVW_EShop_Product_Category "
		wQuery = wQuery & "						WHERE	 CategoryCode1 = '" & SCode1 & "' "
		IF SCode2 <> "" THEN
				wQuery = wQuery & "						 AND CategoryCode2 = '" & SCode2 & "' "
		END IF
		IF SCode3 <> "" THEN
				wQuery = wQuery & "						 AND CategoryCode3 = '" & SCode3 & "' "
		END IF
		wQuery = wQuery & ") "
END IF

SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Product_Select_For_Option_At_ProductList"

		.Parameters.Append .CreateParameter("@WQUERY", adVarChar, adParamInput, 1000, wQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		x = 1
		Do Until oRs.EOF
%>
								<span class="check-style"><input type="checkbox" id="select<%=x%>" name="SSizeCD" value="|<%=oRs("SizeCD")%>|" <%IF INSTR(SSizeCD, "|"&oRs("SizeCD")&"|") THEN%>checked<%END IF%>><label for="select<%=x%>"><span><%=oRs("SizeCD")%></span></label></span>
<%
				oRs.MoveNext
				x = x + 1
		Loop
END IF
oRs.Close
%>
                            </div>
                            <p class="smart-tit">가격대</p>
                            <div class="smart-area range-txt">
                                <input type="text" id="amount" readonly="">
                                <input type="hidden" name="SPrice" id="SPrice" value="<%=SPrice%>" />
                                <input type="hidden" name="EPrice" id="EPrice" value="<%=EPrice%>" />
                            </div>
                            <div class="smart-area area-range">
                                <div class="range-bar ui-slider ui-corner-all ui-slider-horizontal ui-widget ui-widget-content RangeBar">
                                    <div class="ui-slider-range ui-corner-all ui-widget-header"></div>
                                    <span tabindex="0" class="ui-slider-handle ui-corner-all ui-state-default"></span>
                                    <span tabindex="0" class="ui-slider-handle ui-corner-all ui-state-default"></span>
                                </div>
                            </div>
                            <p class="smart-tit">컬러</p>
                            <div class="smart-area pop-color">
<%
'# 컬러 리스트
wQuery = "WHERE 1 = 1 "
IF SCode1 <> "" THEN
		wQuery = wQuery & "AND A.ProductCode IN ( "
		wQuery = wQuery & "						SELECT	 DISTINCT ProductCode "
		wQuery = wQuery & "						FROM	 UVW_EShop_Product_Category "
		wQuery = wQuery & "						WHERE	 CategoryCode1 = '" & SCode1 & "' "
		IF SCode2 <> "" THEN
				wQuery = wQuery & "						 AND CategoryCode2 = '" & SCode2 & "' "
		END IF
		IF SCode3 <> "" THEN
				wQuery = wQuery & "						 AND CategoryCode3 = '" & SCode3 & "' "
		END IF
		wQuery = wQuery & ") "
END IF

SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Product_Select_For_Color_At_ProductList"

		.Parameters.Append .CreateParameter("@WQUERY", adVarChar, adParamInput, 1000, wQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		x = 1
		Do Until oRs.EOF
%>

                                <input type="checkbox" id="color-rainbow<%=x%>" style="background-image:url('<%=oRs("ImageUrl")%>'); background-repeat: no-repeat;" name="SColorCode" value="|<%=oRs("ColorCode")%>|" <%IF INSTR(SColorCode, "|"&oRs("ColorCode")&"|") THEN%>checked<%END IF%>><label for="color-rainbow<%=x%>"><span class="hidden"><%=oRs("ColorCode")%></span></label>               
<%
				oRs.MoveNext
				x = x + 1
		Loop
END IF
oRs.Close
%>
                            </div>
                            <p class="smart-tit">기타</p>
                            <div class="smart-area pop-etc">
                                <div class="etc-menu">
                                    <span class="checkbox" id="SPickupFlagChecked">
                                        <input type="checkbox" id="etc-pick" name="SPickupFlag" value="Y" <%IF SPickupFlag = "Y" THEN%>checked<%END IF%>>
                                    </span>
                                    <label for="etc-pick">매장픽업</label>
                                </div>
                                <div class="etc-menu">
                                    <span class="checkbox" id="SFreeFlagChecked">
                                        <input type="checkbox" id="etc-delivery" name="SFreeFlag" value="Y" <%IF SFreeFlag = "Y" THEN%>checked<%END IF%>>
                                    </span>
                                    <label for="etc-delivery">무료배송</label>
                                </div>
                                <div class="etc-menu">
                                    <span class="checkbox" id="SReserveFlagChecked">
                                        <input type="checkbox" id="etc-reserve" name="SReserveFlag" value="Y" <%IF SReserveFlag = "Y" THEN%>checked<%END IF%>>
                                    </span>
                                    <label for="etc-reserve">예약주문 상품</label>
                                </div>
                            </div>
                        </div>
						</form>
                        <!-- HTML 수정 18-11-30 --->
                    </div>
                    <div class="btns">
                        <button type="button" class="button ty-black" onclick="init_SmartSearch();">초기화</button>
                        <button type="button" class="button ty-red" onclick="document.form1.submit();">적용</button>
                    </div>
                </div>

            </div>
        </div>
    </section>

	<script type="text/javascript">
		var sPrice = "<%=SPrice%>";
		var ePrice = "<%=EPrice%>";

		$('.RangeBar').slider({
			range: true,
			min: 0,
			max: 30,
			values: [sPrice, ePrice],
			slide: function (event, ui) {
				$('#amount').val(ui.values[0] + '만원' + ' - ' + ui.values[1] + '만원');
				$("#SPrice").val(ui.values[0]);
				$("#EPrice").val(ui.values[1]);
			}
		});
		$('#amount').val($('.RangeBar').slider('values', 0) + '만원' +
			' ~ ' + $('.RangeBar').slider('values', 1) + '만원');
	</script>

<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>
