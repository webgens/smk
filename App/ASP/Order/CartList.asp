<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'CartList.asp - 장바구니 리스트
'Date		: 2018.12.27
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
PageCode2 = "01"
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
DIM oRs1						'# ADODB Recordset 개체
DIM oRs2						'# ADODB Recordset 개체
DIM oCmd						'# ADODB Command 개체

DIM i
DIM j
DIM x
DIM y

DIM wQuery						'# WHERE 절
DIM sQuery						'# SORT 절

DIM CartCount

DIM SalePrice
DIM DCRate
DIM ProductImage
DIM EventProdNM

DIM TotalOrderCnt			: TotalOrderCnt			= 0
DIM TotalTagPrice			: TotalTagPrice			= 0
DIM TotalSalePrice			: TotalSalePrice		= 0
DIM TotalDeliveryPrice		: TotalDeliveryPrice	= 0
DIM ShopOrderCnt			: ShopOrderCnt			= 0
DIM ShopTagPrice			: ShopTagPrice			= 0
DIM ShopSalePrice			: ShopSalePrice			= 0
DIM ShopDeliveryPrice		: ShopDeliveryPrice		= 0

Dim WiderTracking_ProductInfo
Dim GoogleTag_ProductInfo
Dim Tracking_ProductInfo
Dim Temp_Tracking_ProductInfo
Dim FaceBookTracking_ProductInfo
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'



SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
SET oRs1		 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
SET oRs2		 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성



'# 장바구니 UserID 변경
IF U_NUM <> "" THEN
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Cart_Update_For_UserID"

				.Parameters.Append .CreateParameter("@CartID",			adVarChar,	adParamInput,  20,	 U_CARTID)
				.Parameters.Append .CreateParameter("@UserID",			adVarChar,	adParamInput,  20,	 U_NUM)

				.Execute, , adExecuteNoRecords
		END WITH
		SET oCmd = Nothing
END IF



'# 장바구니 초기화
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Cart_For_Init"

		.Parameters.Append .CreateParameter("@CartID",			adVarChar,	adParamInput,  20,	 U_CARTID)
		.Parameters.Append .CreateParameter("@UserID",			adVarChar,	adParamInput,  20,	 U_NUM)
		.Parameters.Append .CreateParameter("@CreateID",		adVarChar,	adParamInput,  20,	 U_NUM)
		.Parameters.Append .CreateParameter("@CreateIP",		adVarChar,	adParamInput,  15,	 U_IP)

		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing


'# 장바구니 건수
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Cart_Select_For_CartCount_By_CartID"

		.Parameters.Append .CreateParameter("@CartID",			adVarChar,	adParamInput,  20,	 U_CARTID)
		.Parameters.Append .CreateParameter("@UserID",			adVarChar,	adParamInput,  20,	 U_NUM)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing
																
IF NOT oRs.EOF THEN
		CartCount	= oRs("CartCount")
END IF
oRs.Close
%>

<!-- #include virtual="/INC/Header.asp" -->
	<style type="text/css">
		.listview .ico-soldout { width: 31px; line-height: 31px; border-radius: 50px; font-family: NanumSquare; color: #fff; font-size: 12px; font-weight: 700; background-color: #ff201b; }
		.cart-tit .delivery { float: right; }
		.cart-show .listitems .thumbnail { height: 112px; }
	</style>
	<script type="text/javascript">
		/* 전체 상품 선택*/
		function selectAll() {
			if ($("#check-all").prop("checked") == true) {
				$(".shoemarker-delivery li input:checkbox").each(function (i) {
					$(this).prop("checked", true);
					FormCheckbox.change($(this));
					CartItemChecked($(this).val(), "Y");
				});
			} else {
				$(".shoemarker-delivery li input:checkbox").each(function (i) {
					$(this).prop("checked", false);
					FormCheckbox.change($(this));
					CartItemChecked($(this).val(), "N");
				});
			}
		}

		/* 선택 상품 삭제*/
		function selectDel() {
			if ($(".shoemarker-delivery li input:checkbox:checked").length == 0) {
				openAlertLayer("alert", "삭제할 상품을 선택해 주십시오.", "closePop('alertPop', '');", "");
				return;
			}
			
			openAlertLayer("confirm", "선택한 상품을 삭제 하시겠습니까?", "closePop('confirmPop', '')", "closePop('confirmPop', '');deleteCart('PARTICIAL');");
		}

		/* 상품 삭제 */
		function deleteProduct(cartIdx, productcode) {
			openAlertLayer("confirm", "장바구니에서 상품을 삭제 하시겠습니까?", "closePop('confirmPop', '')", "closePop('confirmPop', '');deleteProduct2(" + cartIdx + ", "+productcode+");");
		}
		function deleteProduct2(cartIdx, productcode) {

			//에이스카운터 장바구니 선택 삭제
			AM_DEL(productcode,1);

			$.ajax({
				type		 : "get",
				url			 : "/ASP/Order/Ajax/CartProductDeleteOne.asp",
				async		 : true,
				data		 : "CartIdx=" + cartIdx,
				dataType	 : "text",
				success		 : function (data) {
								var splitData	 = data.split("|||||");
								var result		 = splitData[0];
								var cont		 = splitData[1];

								if (result == "OK") {
									PageReload();
									return;
								}
								else if (result == "LOGIN") {
									PageReload();
								}
								else {
									openAlertLayer("alert", cont, "closePop('alertPop', '');", "");
									return;
								}
								return;
				},
				error		 : function (data) {
								openAlertLayer("alert", "처리 도중 오류가 발생하였습니다.", "closePop('alertPop', '');", "");
				}
			});
		}

		/* 장바구니 삭제 */
		function deleteCart(flag) {

			//에이스카운터 장바구니 선택 삭제
			$("input[name='CartIdx']:checked").each(function () {
				var v1 = $(this).parent().parent().attr("ProductCode");
				AM_DEL(v1,1);
			});

			$.ajax({
				type		 : "get",
				url			 : "/ASP/Order/Ajax/CartProductDelete.asp",
				async		 : true,
				data		 : "Flag=" + flag,
				dataType	 : "text",
				success		 : function (data) {
								var splitData	 = data.split("|||||");
								var result		 = splitData[0];
								var cont		 = splitData[1];

								if (result == "OK") {
									PageReload();
									return;
								}
								else if (result == "LOGIN") {
									PageReload();
								}
								else {
									openAlertLayer("alert", cont, "closePop('alertPop', '');", "");
									return;
								}
								return;
				},
				error		 : function (data) {
								openAlertLayer("alert", "처리 도중 오류가 발생하였습니다.", "closePop('alertPop', '');", "");
				}
			});
		}

		/* 장바구니 업체 상품 전체 선택/해제 */
		/*
		function selectShop(shopCD) {
			var flag = "";
			if ($("#chk_shop_" + shopCD).prop("checked") == true) {
				flag = "Y";
			} else {
				flag = "N";
			}

			$(".chk_shop_" + shopCD).each(function (i) {
				if (flag == "Y") {
					$(this).prop("checked", true);
				} else {
					$(this).prop("checked", false);
				}

				CartItemChecked($(this).val(), flag);
			});
		}
		*/

		/* 장바구니 상품 선택/해제 */
		function selectProduct(shopCD, cartIdx) {
			var flag = "";
			if ($("#chk_" + cartIdx).prop("checked") == true) {
				flag = "Y";
			} else {
				flag = "N";
			}

			if (flag == "Y") {
				$("#chk_" + cartIdx).prop("checked", true);
			} else {
				$("#chk_" + cartIdx).prop("checked", false);
			}
			/*
			// 업체별 상품전체 선택이 있을 경우
			if ($(".chk_shop_" + shopCD).length == $(".chk_shop_" + shopCD + ":checked").length) {
				$("#chk_shop_" + shopCD).prop("checked", true);
			} else {
				$("#chk_shop_" + shopCD).prop("checked", false);
			}
			*/
			CartItemChecked(cartIdx, flag);
		}

		/* 장바구니 상품 선택/해제 */
		function CartItemChecked(cartIdx, flag) {
			$.ajax({
				type		 : "get",
				url			 : "/ASP/Order/Ajax/CartProductSelectUpdate.asp",
				async		 : true,
				data		 : "CartIdx=" + cartIdx + "&Flag=" + flag,
				dataType	 : "text",
				success		 : function (data) {
								var splitData	 = data.split("|||||");
								var result		 = splitData[0];
								var cont		 = splitData[1];

								if (result == "OK") {
									return;
								}
								else if (result == "LOGIN") {
									PageReload();
								}
								else {
									openAlertLayer("alert", cont, "closePop('alertPop', '');", "");
									return;
								}
								return;
				},
				error		 : function (data) {
								openAlertLayer("alert", "처리 도중 오류가 발생하였습니다.", "closePop('alertPop', '');", "");
				}
			});
		}

		/* 사이즈 변경 팝업 띄우기 */
		function openOptionChange(cartIdx) {
			$.ajax({
				type		 : "post",
				url			 : "/ASP/Order/Ajax/CartProductOptionChange.asp",
				async		 : false,
				data		 : "CartIdx=" + cartIdx,
				dataType	 : "text",
				success		 : function (data) {
								var splitData	 = data.split("|||||");
								var result		 = splitData[0];
								var cont		 = splitData[1];

								if (result == "OK") {
									$("#DimDepth1").html(cont);
									openPop('DimDepth1');
								}
								else if (result == "LOGIN") {
									PageReload();
								}
								else {
									openAlertLayer("alert", cont, "closePop('alertPop', '');", "");
									return;
								}
				},
				error		 : function (data) {
								openAlertLayer("alert", "처리 도중 오류가 발생하였습니다.", "closePop('alertPop', '');", "");
				}
			});
		}

		/* 사이즈 변경 처리 */
		function changeOption(cartIdx) {
			if ($("#CartOptionChangeLayer input[name='cSizeCD']:checked").length == 0) {
				openAlertLayer("alert", "변경할 사이즈를 선택하여 주십시오.", "closePop('alertPop', '');", "");
				return;
			}
			var cSizeCD = $("#CartOptionChangeLayer input[name='cSizeCD']:checked").val();
			
			openAlertLayer("confirm", "사이즈를 변경 하시겠습니까?", "closePop('confirmPop', '')", "closePop('confirmPop', '');changeOption2(" + cartIdx + ");");
		}
		function changeOption2(cartIdx) {
			if ($("#CartOptionChangeLayer input[name='cSizeCD']:checked").length == 0) {
				openAlertLayer("alert", "변경할 사이즈를 선택하여 주십시오.", "closePop('alertPop', '');", "");
				return;
			}
			var cSizeCD = $("#CartOptionChangeLayer input[name='cSizeCD']:checked").val();

			$.ajax({
				type		 : "post",
				url			 : "/ASP/Order/Ajax/CartProductOptionChangeOk.asp",
				async		 : false,
				data		 : "CartIdx=" + cartIdx + "&ChgSizeCD=" + cSizeCD,
				dataType	 : "text",
				success		 : function (data) {
								var splitData	 = data.split("|||||");
								var result		 = splitData[0];
								var cont		 = splitData[1];

								if (result == "OK") {
									openAlertLayer("alert", "변경 되었습니다.", "closePop('alertPop', '');PageReload();", "");
								}
								else if (result == "LOGIN") {
									PageReload();
								}
								else {
									common_msgPopOpen("장바구니", cont, "", "msgPopup", "N");
									return;
								}
				},
				error		 : function (data) {
								openAlertLayer("alert", "처리 도중 오류가 발생하였습니다.", "closePop('alertPop', '');", "");
				}
			});
		}

		/* 선택상품 주문하기 */
		function selectOrder() {
			if ($(".shoemarker-delivery li input:checkbox:checked").length == 0) {
				openAlertLayer("alert", "주문할 상품을 선택해 주십시오.", "closePop('alertPop', '');", "");
				return;
			}

			CartOrder("PARTICIAL");
		}

		/* 주문하기 */
		function CartOrder(flag) {
			$.ajax({
				type		 : "get",
				url			 : "/Asp/Order/Ajax/OrderSheet_CartAddOk.asp",
				async		 : true,
				data		 : "Flag=" + flag,
				dataType	 : "html",
				success		 : function (data) {
								var splitData	 = data.split("|||||");
								var result		 = splitData[0];
								var cont		 = splitData[1];

								if (result == "OK") {
									//APP_GoUrl("/ASP/Order/Order.asp?IsOrder=Yes&AccessType=Cart");
									location.href = "/ASP/Order/Order.asp?IsOrder=Yes&AccessType=Cart";
								}
								else if (result == "LOGIN") {
									PageReload();
								}
								else {
									openAlertLayer("alert", cont, "closePop('alertPop', '');PageReload();", "");
									return;
								}

								return;
				},
				error		 : function (data) {
								openAlertLayer("alert", "처리 도중 오류가 발생하였습니다.", "closePop('alertPop', '');", "");
				}
			});
		}
	</script>

<%TopSubMenuTitle = "장바구니"%>
<!-- #include virtual="/INC/TopCart.asp" -->


<%
'# 장바구니 담긴 상품이 있을 경우
IF CInt(CartCount) > 0 THEN
%>
    <!-- Main -->
    <main id="container" class="container">
        <div class="sub_content">
            <div class="cart cart-show">
                <div class="checkall">
                    <span class="checkbox">
                        <input type="checkbox" name="check-all" id="check-all" onclick="selectAll()" data-allchk="select" checked />
                    </span>
                    <div class="checkall-txt">
                        <label for="check-all">전체선택</label>
                        <button class="checkdelete" onclick="selectDel()" type="button">선택삭제</button>
                    </div>
                </div>
<%
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Cart_Select_For_ShopList_By_CartID"

				.Parameters.Append .CreateParameter("@CartID",			adVarChar,	adParamInput,  20,	 U_CARTID)
				.Parameters.Append .CreateParameter("@UserID",			adVarChar,	adParamInput,  20,	 U_NUM)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		SET oCmd = Nothing

		IF NOT oRs.EOF THEN
				Do Until oRs.EOF

						ShopOrderCnt		= 0
						ShopTagPrice		= 0
						ShopSalePrice		= 0
						ShopDeliveryPrice	= 0
%>
                <div class="shoemarker-delivery">
                    <p class="cart-tit">
						<%=oRs("ShopNM")%> 배송
						<%IF CDbl(oRs("ShopDeliveryPrice")) = 0 THEN%>
						<span class="delivery">무료배송</span>
						<%ELSE%>
						<span class="delivery">배송비 : <%=FormatNumber(oRs("ShopDeliveryPrice"),0)%>원</span>
						<%END IF%>
                    </p>
                    <div class="item-list">
                        <ul class="listview">
						<%
						SET oCmd = Server.CreateObject("ADODB.Command")
						WITH oCmd
								.ActiveConnection	 = oConn
								.CommandType		 = adCmdStoredProc
								.CommandText		 = "USP_Front_EShop_Cart_Select_By_CartID_N_ShopCD"

								.Parameters.Append .CreateParameter("@CartID",	adVarChar,	adParamInput,  20,	U_CARTID)
								.Parameters.Append .CreateParameter("@UserID",	adVarChar,	adParamInput,  20,	U_NUM)
								.Parameters.Append .CreateParameter("@ShopCD",	adChar,		adParamInput,  6,	oRs("ShopCD"))
						END WITH
						oRs1.CursorLocation = adUseClient
						oRs1.Open oCmd, , adOpenStatic, adLockReadOnly
						SET oCmd = Nothing

						IF NOT oRs1.EOF THEN
								y = 0
								Do Until oRs1.EOF
										IF oRs1("SalePriceType") = "2" THEN
												SalePrice	= oRs1("EmployeeSalePrice")
												DCRate		= oRs1("EmployeeDCRate")
										ELSE
												SalePrice	= oRs1("SalePrice")
												DCRate		= oRs1("DCRate")
										END IF

										IF oRs1("ProductImage_180") = "" THEN
												ProductImage = "/Images/180_noimage.png"
										ELSE
												ProductImage = oRs1("ProductImage_180")
										END IF

										'# 사은품정보
										EventProdNM		= ""
										SET oCmd = Server.CreateObject("ADODB.Command")
										WITH oCmd
												.ActiveConnection	 = oConn
												.CommandType		 = adCmdStoredProc
												.CommandText		 = "USP_Front_EShop_SubProduct_Event_Select_By_ProductCode"

												.Parameters.Append .CreateParameter("@ProductCode",		 adInteger, adParaminput,		, oRs1("ProductCode"))
										End WITH
										oRs2.CursorLocation = adUseClient
										oRs2.Open oCmd, , adOpenStatic, adLockReadOnly
										SET oCmd = Nothing

										IF NOT oRs2.EOF THEN
												Do Until oRs2.EOF
														IF EventProdNM = "" THEN
																EventProdNM		= oRs2("EventProdNM")
														ELSE
																EventProdNM		= EventProdNM & ", " & oRs2("EventProdNM")
														END IF

														IF oRs2("Qty") > 1 THEN
																EventProdNM		= EventProdNM & "(" & oRs2("Qty") & ")"
														END IF

														oRs2.MoveNext
												Loop
										END IF
										oRs2.Close

										'트랙킹 코드
										If y = 0 Then
											WiderTracking_ProductInfo = "{ i: """ & oRs1("ProductCode") & """, t: """ & oRs1("ProductName") & """ }"
											Tracking_ProductInfo = oRs1("SalePrice") & "|,|" & oRs1("ProductCode") & "|,|" & oRs1("ProductName") & "|,|" & oRs1("BrandName")
											FaceBookTracking_ProductInfo = """" & oRs1("ProductCode") & """"
										Else
											WiderTracking_ProductInfo = WiderTracking_ProductInfo & ", { i: """ & oRs1("ProductCode") & """, t: """ & oRs1("ProductName") & """ }"
											Tracking_ProductInfo = Tracking_ProductInfo & "|||" & oRs1("SalePrice") & "|,|" & oRs1("ProductCode") & "|,|" & oRs1("ProductName") & "|,|" & oRs1("BrandName")
											FaceBookTracking_ProductInfo = FaceBookTracking_ProductInfo & "," & """" & oRs1("ProductCode") & """"
										End If

										GoogleTag_ProductInfo = GoogleTag_ProductInfo & "brandIds.push('" & oRs1("ProductCode") & "');"
						%>
                            <li ProductCode="<%=oRs1("ProductCode")%>">
								<%IF oRs1("Soldout") = "Y" THEN%>
                                <span class="ico-soldout">품절</span>
								<%ELSE%>
                                <span class="checkbox">
                                    <input type="checkbox" name="CartIdx" class="chk_shop_<%=oRs("ShopCD")%>" id="chk_<%=oRs1("Idx")%>" value="<%=oRs1("Idx")%>" onclick="selectProduct('<%=oRs("ShopCD")%>','<%=oRs1("Idx")%>')" />
                                    <label class="hidden" for="chk_<%=oRs1("Idx")%>">상품 체크</label>
                                </span>
								<%END IF%>
                                <a href="javascript:APP_GoUrl('/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs1("ProductCode")%>');" class="listitems">
                                    <div class="thumbnail"><img src="<%=ProductImage%>" alt=""></div>
                                    <div class="itemtxt">
                                        <div class="item-info">
                                            <p class="brand-name"><%=oRs1("BrandName")%></p>
                                            <h1 class="product-name pname"><%=oRs1("ProductName")%></h1>
                                            <div class="item-option">
                                                <span>사이즈 : <%=oRs1("SIzeCD")%></span>
                                                <span>수량 : <%=oRs1("OrderCnt")%></span>
                                            </div>
                                        </div>
                                    </div>
                                </a>
								<button type="button" onclick="openOptionChange('<%=oRs1("Idx")%>')" class="btn-change-option">옵션변경</button>
                                <span class="closebtn">
                                    <button type="button" onclick="deleteProduct('<%=oRs1("Idx")%>', '<%=oRs1("ProductCode")%>')">
                                        <p class="hidden">삭제</p>
                                    </button>
                                </span>
								<%
								'# 1+1 상품
								IF oRs1("GroupCnt") > 1 THEN
										SET oCmd = Server.CreateObject("ADODB.Command")
										WITH oCmd
												.ActiveConnection	 = oConn
												.CommandType		 = adCmdStoredProc
												.CommandText		 = "USP_Front_EShop_Cart_Select_For_OnePlusOne"

												.Parameters.Append .CreateParameter("@CartID",		adVarChar,	adParamInput,  20,	oRs1("CartID"))
												.Parameters.Append .CreateParameter("@GroupIdx",	adInteger,	adParamInput,    ,	oRs1("GroupIdx"))
										END WITH
										oRs2.CursorLocation = adUseClient
										oRs2.Open oCmd, , adOpenStatic, adLockReadOnly
										SET oCmd = Nothing

										IF NOT oRs2.EOF THEN
												'# IF oRs2("ProductImage") = "" THEN
												'# 		ProductImage = "/Images/180_noimage.png"
												'# ELSE
												'# 		ProductImage = oRs2("ProductImage")
												'# END IF
								%>
                                <div class="change-cont">
                                    <span class="tit">1+1상품</span><span class="cont"><%=oRs2("ProductName")%> (<%=oRs2("SIzeCD")%>)</span>
                                    <button type="button" onclick="openOptionChange('<%=oRs2("Idx")%>')" class="btn-change">변경</button>
                                </div>
								<%
										END IF
										oRs2.Close
								END IF
								%>
								<%IF EventProdNM <> "" THEN%>
                                <div class="change-cont">
                                    <span class="tit">사은품</span><span class="cont"><%=EventProdNM%></span>
                                </div>
								<%END IF%>
                                <div class="price">
                                    <div class="price-info">
                                        <div class="info-wrap">
                                            <p class="price-tit">상품금액</p>
                                            <p class="price-value"><%=FormatNumber(oRs1("TagPrice"),0)%>원</p>
                                        </div>
                                        <div class="info-wrap">
                                            <p class="price-tit">할인금액</p>
                                            <p class="price-value"><%=FormatNumber(CDbl(oRs1("TagPrice")) - CDbl(SalePrice),0)%>원</p>
                                        </div>
                                    </div>
                                    <div class="order-price">
                                        <strong>주문금액</strong>
                                        <p><%=FormatNumber(SalePrice,0)%>원</p>
                                    </div>
                                </div>
                            </li>
						<%
										ShopOrderCnt	= ShopOrderCnt		+ CDbl(oRs1("OrderCnt"))
										ShopTagPrice	= ShopTagPrice		+ CDbl(oRs1("TagPrice"))
										ShopSalePrice	= ShopSalePrice		+ CDbl(SalePrice)

										y = y + 1
										oRs1.MoveNext
								Loop 
						END IF
						oRs1.Close

						'# 배송비 계산
						IF ShopSalePrice < CDbl(oRs("StandardPrice")) THEN
								ShopDeliveryPrice	= CDbl(oRs("DeliveryPrice"))
						END IF
						%>
                        </ul>
                    </div>
                </div>
<%
						'# 주문금액 계산
						TotalOrderCnt		= TotalOrderCnt			+ ShopOrderCnt
						TotalTagPrice		= TotalTagPrice			+ ShopTagPrice
						TotalSalePrice		= TotalSalePrice		+ ShopSalePrice
						TotalDeliveryPrice	= TotalDeliveryPrice	+ ShopDeliveryPrice

						oRs.MoveNext
				Loop 
		END IF
		oRs.Close
%>
                <div class="price-result">
                    <strong class="cart-tit"><span class="hidden">주문 결과</span></strong>
                    <div class="item-list">
                        <ul class="listview">
                            <li>
                                <div class="price">
                                    <div class="price-info">
                                        <div class="info-wrap">
                                            <p class="price-tit">주문 상품 수</p>
                                            <p class="price-value"><%=FormatNumber(TotalOrderCnt,0)%>개</p>
                                        </div>
                                        <div class="info-wrap">
                                            <p class="price-tit">총 주문금액</p>
                                            <p class="price-value"><%=FormatNumber(TotalTagPrice,0)%>원</p>
                                        </div>
                                        <div class="info-wrap">
                                            <p class="price-tit">총 할인금액</p>
                                            <p class="price-value"><%=FormatNumber((TotalTagPrice - TotalSalePrice) * -1,0)%>원</p>
                                        </div>
                                        <div class="info-wrap">
                                            <p class="price-tit">총 배송비</p>
                                            <p class="price-value"><%=FormatNumber(TotalDeliveryPrice,0)%>원</p>
                                        </div>
                                    </div>
                                    <div class="order-price">
                                        <strong>총 결제금액</strong>
                                        <p><%=FormatNumber(TotalSalePrice + TotalDeliveryPrice,0)%>원</p>
                                    </div>
                                </div>
                            </li>
                        </ul>
                    </div>
                </div>
                <div class="inf-type1">
                    <p class="tit">상품별 할인쿠폰은 결제단계에서 적용하실 수 있습니다.</p>
                </div>

            </div>
        </div>
    </main>
<%
ELSE
%>
    <!-- Main -->
    <main id="container" class="container cart-empty-container">
        <div class="content content-empty">
            <div class="cart cart-empty">
                <p>장바구니에 담은 상품이 없습니다.</p>
            </div>
        </div>
    </main>
<%
END IF
%>



<!-- #include virtual="/INC/FooterNoBNB.asp" -->

<%
'# 장바구니 담긴 상품이 있을 경우
IF CInt(CartCount) > 0 THEN
%>
    <!-- bnb-ty2 -->
    <article class="bnb-ty2 cart-show" style="background-color: #ff201b;">
        <button type="button" onclick="selectOrder()" class="btn-buy">구매하기</button>
    </article>
<%
END IF
%>

<script type="text/javascript">
	$(function () {
		selectAll();
	});
</script>

<!-- WIDERPLANET  SCRIPT START 2019.1.8 -->
<div id="wp_tg_cts" style="display:none;"></div>
<script type="text/javascript">
	var wptg_tagscript_vars = wptg_tagscript_vars || [];
	wptg_tagscript_vars.push(
	(function () {
		return {
			wp_hcuid: "<%=U_Num%>",  	/*고객넘버 등 Unique ID (ex. 로그인  ID, 고객넘버 등 )를 암호화하여 대입.
				 *주의 : 로그인 하지 않은 사용자는 어떠한 값도 대입하지 않습니다.*/
			ti: "24585",
			ty: "Cart",
			device: "mobile"
			, items: [
				 <%=WiderTracking_ProductInfo%>
			]
		};
	}));
</script>
<script type="text/javascript" async src="//cdn-aitg.widerplanet.com/js/wp_astg_4.0.js"></script>
<!-- // WIDERPLANET  SCRIPT END 2019.1.8 -->

<!-- Google Tag Manager Variable (eMnet) -->
<script type="text/javascript">
	var brandIds = [];
	<%=GoogleTag_ProductInfo%>
</script>
<!-- End Google Tag Manager Variable (eMnet) --> 

<%
	'0:금액, 1:코드, 2:상품명, 3:브랜드
	Temp_Tracking_ProductInfo = Split(Tracking_ProductInfo, "|||")
	For i = 0 To UBound(Temp_Tracking_ProductInfo)
		If Trim(Temp_Tracking_ProductInfo(i)) <> "" Then
			Tracking_ProductInfo = Split(Temp_Tracking_ProductInfo(i), "|,|")
%>
<!-- AceCounter Mobile eCommerce (Cart_Inout) v7.5 Start -->
<script type="text/javascript">
	var AM_Cart=(function(){
		var c={pd:'<%=Trim(Tracking_ProductInfo(1))%>',pn:'<%=Trim(Tracking_ProductInfo(2))%>',am:'<%=Trim(Tracking_ProductInfo(0))%>',qy:'1',ct:'<%=Trim(Tracking_ProductInfo(3))%>'};
		var u=(!AM_Cart)?[]:AM_Cart; u[c.pd]=c;return u;
	})();
</script>

<script type="text/javascript">
	gtag('event', 'add_to_cart', {
		"items": [
			{
				"id": "<%=Trim(Tracking_ProductInfo(1))%>",
				"name": "<%=Trim(Tracking_ProductInfo(2))%>",
				"list_name": "Cart",
				"brand": "<%=Trim(Tracking_ProductInfo(3))%>",
				"category": "",
				"variant": "",
				"list_position": 1,
				"quantity": 1,
				"price": '<%=Trim(Tracking_ProductInfo(0))%>'
			}
		]
	});
</script>
<%
		End If
	Next
%>

<script type="text/javascript">
	fbq('track', 'AddToCart', {
		content_type: 'product',
		content_ids: [<%=FaceBookTracking_ProductInfo%>],
		value: '<%=TotalTagPrice%>',
		currency: 'KRW',
	});
</script>

<script type="text/javascript" src="//wcs.naver.net/wcslog.js"></script>
<script type="text/javascript">
	var _nasa = {}; _nasa["cnv"] = wcs.cnv('3', '<%=TotalTagPrice%>');
</script>

<!-- kakao pixel script //-->
<script type="text/javascript" charset="UTF-8" src="//t1.daumcdn.net/adfit/static/kp.js"></script>
<script type="text/javascript">
	kakaoPixel('5354511058043421336').pageView();
	kakaoPixel('5354511058043421336').viewCart();
</script>
<!-- kakao pixel script //-->


<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs2 = Nothing
SET oRs1 = Nothing
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>