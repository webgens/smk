<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'MyPickList.asp - 찜한 상품 / 최근 본 상품 / 브랜드
'Date		: 2018.12.17
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
PageCode1 = "05"
PageCode2 = "03"
PageCode3 = "00"
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->
<!-- #include virtual="/Common/SubCheckID.asp" -->

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


DIM PickType					'# 탭구분
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

PickType			 = sqlFilter(request("PickType"))
IF PickType = "" THEN PickType	= "1"


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>


<!-- #include virtual="/INC/Header.asp" -->
	<script type="text/javascript">
		// 탭 선택시
		function chgTab(listType) {
			$("#tabs .tab-selector li").removeClass("active");
			//$("#tabs-col1").removeClass("active");
			//$("#tabs-col2").removeClass("active");
			//$("#tabs-col3").removeClass("active");

			if (listType == "1") {
				$("#tabs .tab-selector li").eq(0).addClass("active");
				location.replace("/ASP/Mypage/MyPickList.asp?PickType=1");
				//$("#tabs-col1").addClass("active");
			}
			else if (listType == "2") {
				$("#tabs .tab-selector li").eq(1).addClass("active");
				location.replace("/ASP/Mypage/MyPickList.asp?PickType=2");
				//$("#tabs-col2").addClass("active");
			}
			else {
				$("#tabs .tab-selector li").eq(2).addClass("active");
				location.replace("/ASP/Mypage/MyPickList.asp?PickType=3");
				//$("#tabs-col3").addClass("active");
			}
		}

		/* 찜한상품 삭제 */
		function productPickDel(productCode) {
			//location.href = "/ASP/Mypage/Ajax/PwdModifyOk.asp?" + $("#chgPwdForm").serialize();
			//return;

			if (productCode == "") {
				common_msgPopOpen("", "선택하신 상품이 없습니다.", "", "msgPopup", "N");
				return;
			}

			common_msgPopOpen("", "찜한 상품을 삭제 하시겠습니까?", "productPickDel2('" + productCode + "');", "msgPopup", "C");
		}
		function productPickDel2(productCode) {
			$.ajax({
				type: "post",
				url: "/ASP/Mypage/Ajax/Product_Pick_Delete.asp",
				async: false,
				data: "ProductCode=" + productCode,
				dataType: "text",
				success: function (data) {
					var splitData = data.split("|||||");
					var result = splitData[0];
					var cont = splitData[1];


					if (result == "OK") {
						PageReload();
						return;
					}
					else if (result == "LOGIN") {
						PageReload();
						return;
					}
					else {
						common_msgPopOpen("", cont, "", "msgPopup", "N");
						return;
					}
				},
				error: function (data) {
					//alert(data.responseText)
					common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
				}
			});
		}


		/* 최근 본 상품 삭제 */
		function productLatestDel(productCode) {
			if (productCode == "") {
				common_msgPopOpen("", "선택하신 상품이 없습니다.", "", "msgPopup", "N");
				return;
			}

			common_msgPopOpen("", "최근 본 상품을 삭제 하시겠습니까?", "productLatestDel2('" + productCode + "');", "msgPopup", "C");
		}
		function productLatestDel2(productCode) {
			$.ajax({
				type: "post",
				url: "/ASP/Mypage/Ajax/Product_Latest_Delete.asp",
				async: false,
				data: "ProductCode=" + productCode,
				dataType: "text",
				success: function (data) {
					var splitData = data.split("|||||");
					var result = splitData[0];
					var cont = splitData[1];


					if (result == "OK") {
						PageReload();
						return;
					}
					else if (result == "LOGIN") {
						PageReload();
						return;
					}
					else {
						common_msgPopOpen("", cont, "", "msgPopup", "N");
						return;
					}
				},
				error: function (data) {
					//alert(data.responseText)
					common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
				}
			});
		}

		/* 브랜드 선택 삭제 */
		function brandPickDel(brandCode) {
			if (brandCode == "") {
				common_msgPopOpen("", "선택하신 브랜드가 없습니다.", "", "msgPopup", "N");
				return;
			}

			common_msgPopOpen("", "선택하신 브랜드를 삭제 하시겠습니까?", "brandPickDel2('" + brandCode + "');", "msgPopup", "C");
		}
		function brandPickDel2(brandCode) {
			$.ajax({
				type: "post",
				url: "/ASP/Mypage/Ajax/Brand_Pick_Delete.asp",
				async: false,
				data: "BrandCode=" + brandCode,
				dataType: "text",
				success: function (data) {
					var splitData = data.split("|||||");
					var result = splitData[0];
					var cont = splitData[1];


					if (result == "OK") {
						PageReload();
						return;
					}
					else if (result == "LOGIN") {
						PageReload();
						return;
					}
					else {
						common_msgPopOpen("", cont, "", "msgPopup", "N");
						return;
					}
				},
				error: function (data) {
					//alert(data.responseText)
					common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
				}
			});
		}

	</script>

<%TopSubMenuTitle = "MY슈마커"%>
<!-- #include virtual="/INC/TopSub.asp" -->


    <!-- Main -->
    <main id="container" class="container">
        <div class="sub_content">

            <div class="wrap-mypage">
				<div style="height:8px"></div>
				
                <div id="MypageSubMenu" class="ly-title accordion">
                    <div class="selector">
	                    <button type="button" class="btn-list clickEvt" data-target="MypageSubMenu">MY&hearts;</button>
					</div>
					<div class="option my-recode">
						<!-- #include virtual="/ASP/Mypage/SubMenu_MyShoeMarker.asp" -->
					</div>
                </div>


                <div class="mypage-my">
                    <div id="shoppingList">
                        <div>
                            <div id="tabs">
                                <div class="tab-mypage">
                                    <ul class="tab-selector">
                                        <li class="part-3<%IF PickType = "1" THEN%> active<%END IF%>"><a href="javascript:chgTab('1')">찜한 상품</a></li>
                                        <li class="part-3<%IF PickType = "2" THEN%> active<%END IF%>"><a href="javascript:chgTab('2')">최근 본 상품</a></li>
                                        <li class="part-3<%IF PickType = "3" THEN%> active<%END IF%>"><a href="javascript:chgTab('3')">MY브랜드</a></li>
                                    </ul>
<%
'# 찜한 상품
IF PickType = "1" THEN
%>
<%
		Set oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection = oConn
				.CommandType = adCmdStoredProc
				.CommandText = "USP_Front_EShop_Product_Pick_Select_By_MemberNum"

				.Parameters.Append .CreateParameter("@MemberNum",	adInteger,	adParamInput, ,		U_NUM)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		Set oCmd = Nothing
%>
                                    <!-- 찜한 상품 -->
                                    <div id="tabs-col1" class="tab-panel active">
                                        <div class="h-line">
                                            <h2 class="h-level4">찜한상품 전체</h2>
                                            <span class="h-num"><%=oRs.RecordCount%>건</span>
											<%IF oRs.RecordCount > 0 THEN%>
                                            <span class="h-date is-right">
                                                <button type="button" onclick="productPickDel('ALL')" class="button-ty3 ty-bd-black">
                                                    <span>전체 삭제</span>
	                                            </button>
                                            </span>
											<%END IF%>
                                        </div>
<%
		IF NOT oRs.EOF THEN
%>
                                        <ul>
<% 
				i = 1
				Do Until oRs.EOF
%>
                                            <li class="informItem">
                                                <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>')">
                                                    <span class="cont">
                                                        <span class="thumbNail">
                                                            <span class="img">
                                                                <img src="<%=oRs("ImageUrl_180")%>" alt="상품 이미지">
                                                            </span>
                                                        </span>

                                                        <span class="detail">
                                                            <span class="brand">
                                                                <span class="name"><%=oRs("BrandName")%></span>
                                                            </span>
                                                            <span class="product-name"><em><%=oRs("ProductName")%></em></span>

                                                            <span class="product-more">
                                                                <span class="price"><strong><%=FormatNumber(oRs("SalePrice"),0)%></strong>원</span>
                                                                <span class="optional-info">
                                                                    <span class="icon ico-fav"><%=oRs("WishCnt")%></span>
                                                                    <span class="icon ico-cmt"><%=oRs("ReviewCnt")%></span>
                                                                </span>
                                                            </span>
                                                        </span>
                                                    </span>
                                                </a>

                                                <div class="buttongroup">
                                                    <button type="button" onclick="APP_GoUrl('/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>')" class="button-ty2 is-expand ty-bd-gray">상품 상세정보 보기</button>
                                                </div>
                                                <div class="right-circle">
                                                    <button type="button" onclick="productPickDel('<%=oRs("ProductCode")%>')" class="closebtn">
                                                        <span class="hidden">삭제</span>
                                                    </button>
                                                </div>
                                            </li>
<%
						oRs.MoveNext
						i = i + 1
				LOOP
%>
                                        </ul>

                                        <div class="inf-type1" style="padding-bottom:20px">
                                            <p class="tit">알려드립니다.</p>
                                            <ul>
                                                <li class="bullet-ty1">찜한 상품은 최대 30개, 등록일로부터 최대 180일간 저장됩니다.</li>
                                            </ul>
                                        </div>
<%
		ELSE
%>
										<div class="area-empty">
											<span class="icon-empty"></span>
											<p class="tit-empty">찜한 상품이 없습니다.</p>
										</div>
<%
		END IF
		oRs.Close
%>
                                    </div>
                                    <!-- // 찜한 상품 -->
<%
'# 최근 본 상품
ELSEIF PickType = "2" THEN
%>
<%
		wQuery = "WHERE D.MemberNum = '"& U_NUM &"'"

		Set oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection = oConn
				.CommandType = adCmdStoredProc
				.CommandText = "USP_Front_EShop_Product_Latest_Select_By_wQuery_For_Top30"

				.Parameters.Append .CreateParameter("@wQuery",	adVarchar,	adParamInput, 1000,		wQuery)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		Set oCmd = Nothing
%>
                                    <!-- 최근 본 상품 -->
                                    <div id="tabs-col2" class="tab-panel active">
                                        <div class="h-line">
                                            <h2 class="h-level4">최근 본 상품 전체</h2>
                                            <span class="h-num"><%=oRs.RecordCount%>건</span>
											<%IF oRs.RecordCount > 0 THEN%>
                                            <span class="h-date is-right">
                                                <button type="button" onclick="productLatestDel('ALL')" class="button-ty3 ty-bd-black">
                                                    <span>전체 삭제</span>
												</button>
                                            </span>
											<%END IF%>
                                        </div>
<%
		IF NOT oRs.EOF THEN
%>
                                        <ul>
<% 
				i = 1
				Do Until oRs.EOF
%>
                                            <li class="informItem">
                                                <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>')">
                                                    <span class="cont">
                                                        <span class="thumbNail">
                                                            <span class="img">
                                                                <img src="<%=oRs("ImageUrl_180")%>" alt="상품 이미지">
                                                            </span>
                                                        </span>

                                                        <span class="detail">
                                                            <span class="brand">
                                                                <span class="name"><%=oRs("BrandName")%></span>
                                                            </span>
                                                            <span class="product-name"><em><%=oRs("ProductName")%></em></span>

                                                            <span class="product-more">
                                                                <span class="price"><strong><%=FormatNumber(oRs("SalePrice"),0)%></strong>원</span>
                                                                <span class="optional-info">
                                                                    <span class="icon ico-fav"><%=oRs("WishCnt")%></span>
                                                                    <span class="icon ico-cmt"><%=oRs("ReviewCnt")%></span>
                                                                </span>
                                                            </span>
                                                        </span>
                                                    </span>
                                                </a>

                                                <div class="buttongroup">
                                                    <button type="button" onclick="APP_GoUrl('/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>')" class="button-ty2 is-expand ty-bd-gray">상품 상세정보 보기</button>
                                                </div>
                                                <div class="right-circle">
                                                    <button type="button" onclick="productLatestDel('<%=oRs("ProductCode")%>')" class="closebtn">
                                                        <span class="hidden">삭제</span>
                                                    </button>
                                                </div>
                                            </li>
<%
						oRs.MoveNext
						i = i + 1
				LOOP
%>
                                        </ul>

                                        <div class="inf-type1" style="padding-bottom:20px">
                                            <p class="tit">알려드립니다.</p>
                                            <ul>
                                                <li class="bullet-ty1">최근 본 상품을 기준으로 최대 30개까지 저장됩니다.</li>
                                            </ul>
                                        </div>
<%
		ELSE
%>
										<div class="area-empty">
											<span class="icon-empty"></span>
											<p class="tit-empty">최근 본 상품이 없습니다.</p>
										</div>
<%
		END IF
		oRs.Close
%>
                                    </div>
                                    <!-- // 최근 본 상품 -->

<%
'# MY 브랜드
ELSEIF PickType = "3" THEN
%>
<%
		Set oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection = oConn
				.CommandType = adCmdStoredProc
				.CommandText = "USP_Front_EShop_Product_Brand_Pick_Select_By_MemberNum"

				.Parameters.Append .CreateParameter("@MemberNum",	adInteger,	adParamInput, ,		U_NUM)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		Set oCmd = Nothing
%>
                                    <!-- MY브랜드 -->
                                    <div id="tabs-col3" class="tab-panel active">
                                        <div class="h-line">
                                            <h2 class="h-level4">찜한 브랜드 전체</h2>
                                            <span class="h-num"><%=oRs.RecordCount%>건</span>
											<%IF oRs.RecordCount > 0 THEN%>
                                            <span class="h-date is-right">
                                                <button type="button" onclick="brandPickDel('ALL')" class="button-ty3 ty-bd-black">
                                                    <span>전체 삭제</span>
	                                            </button>
                                            </span>
											<%END IF%>
                                        </div>
<%
		IF NOT oRs.EOF THEN
%>
                                        <div class="like-brand">
	                                        <ul>
<% 
				i = 1
				Do Until oRs.EOF
%>
                                                <li class="brand-list">
                                                    <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Product/Brand.asp?SBrandCode=<%=oRs("BrandCode")%>')">
                                                        <p><%=oRs("BrandName")%></p>
                                                        <p class="shortcut">브랜드샵 바로가기</p>
                                                    </a>
                                                    <div class="right-circle">
                                                        <button type="button" onclick="brandPickDel('<%=oRs("BrandCode")%>')" class="closebtn">
                                                            <span class="hidden">삭제</span>
                                                        </button>
                                                    </div>
                                                </li>
<%
						oRs.MoveNext
						i = i + 1
				LOOP
%>
	                                        </ul>
                                        </div>
<%
		ELSE
%>
										<div class="area-empty">
											<span class="icon-empty"></span>
											<p class="tit-empty">선택한 브랜드가 없습니다.</p>
										</div>
<%
		END IF
		oRs.Close
%>
                                    </div>
                                    <!-- //MY브랜드 -->
<%
END IF
%>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </main>

<!-- #include virtual="/INC/FooterNoBNB.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>