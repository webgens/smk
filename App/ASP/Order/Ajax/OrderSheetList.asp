<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'OrderSheetList.asp - 주문서 상품리스트
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

<%
'/****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn											'# ADODB Connection 개체
DIM oRs												'# ADODB Recordset 개체
DIM oRs1											'# ADODB Recordset 개체
DIM oRs2											'# ADODB Recordset 개체
DIM oCmd											'# ADODB Command 개체

DIM i
DIM j
DiM x
DIM y

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절

DIM MultiDelvFlag									'# 다중배송지 선택 여부

DIM OrderCount				: OrderCount			= 0
DIM PointRate				: PointRate				= 0

DIM OrderType
DIM SalePrice
DIM DCRate
DIM OrderPrice
DIM SavePoint
DIM ProductImage
DIM DiscountPrice
DIM EventProdNM

DIM TotalOrderCnt			: TotalOrderCnt			= 0
DIM TotalTagPrice			: TotalTagPrice			= 0
DIM TotalSalePrice			: TotalSalePrice		= 0
DIM TotalUseCouponPrice		: TotalUseCouponPrice	= 0
DIM TotalUseScashPrice		: TotalUseScashPrice	= 0
DIM TotalUsePointPrice		: TotalUsePointPrice	= 0
DIM TotalDeliveryPrice		: TotalDeliveryPrice	= 0
DIM TotalSavePoint			: TotalSavePoint		= 0
DIM ShopOrderCnt			: ShopOrderCnt			= 0
DIM ShopTagPrice			: ShopTagPrice			= 0
DIM ShopSalePrice			: ShopSalePrice			= 0
DIM ShopUseCouponPrice		: ShopUseCouponPrice	= 0
DIM ShopUseScashPrice		: ShopUseScashPrice		= 0
DIM ShopUsePointPrice		: ShopUsePointPrice		= 0
DIM ShopDeliveryPrice		: ShopDeliveryPrice		= 0
DIM ShopSavePoint			: ShopSavePoint			= 0

Dim Tracking_ProductInfo
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


MultiDelvFlag		= sqlFilter(Request("MultiDelvFlag"))
IF MultiDelvFlag = "" THEN MultiDelvFlag = "N"




SET oConn	= ConnectionOpen()	'//커넥션 생성
SET oRs		= Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성
SET oRs1	= Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성
SET oRs2	= Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성


'# 주문서 체크
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Front_EShop_OrderSheet_Select_For_OrderCount"

		.Parameters.Append .CreateParameter("@CartID",	 adVarChar,	 adParamInput, 20,		 U_CARTID)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing
																
IF NOT oRs.EOF THEN
		OrderCount	= CInt(oRs("OrderCount"))
END IF
oRs.Close


IF U_MFLAG = "Y" THEN
		'# 회원정보
		Set oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection = oConn
				.CommandType = adCmdStoredProc
				.CommandText = "USP_Admin_EShop_Member_Select_By_MemberNum"

				.Parameters.Append .CreateParameter("@MemberNum",	 adInteger,	 adParamInput,   ,		 U_NUM)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		Set oCmd = Nothing
																
		IF NOT oRs.EOF THEN
				PointRate			= oRs("PointRate")
		END IF
		oRs.Close
END IF

Response.Write "OK|||||"
%>					
					<input type="hidden" name="MultiDelvFlag" id="MultiDelvFlag" value="<%=MultiDelvFlag%>" />

                    <p class="cart-tit">주문상품 확인</p>
                    <div class="item-list">
                        <ul class="listview">
<%
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_OrderSheet_Select_For_ShopList_By_CartID"

		.Parameters.Append .CreateParameter("@CartID", adVarChar, adParamInput, 20, U_CARTID)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		Do Until oRs.EOF

				ShopOrderCnt		= 0
				ShopTagPrice		= 0
				ShopSalePrice		= 0
				ShopUseCouponPrice	= 0
				ShopUsePointPrice	= 0
				ShopUseScashPrice	= 0
				ShopDeliveryPrice	= 0
				ShopSavePoint		= 0
%>

<%
				SET oCmd = Server.CreateObject("ADODB.Command")
				WITH oCmd
						.ActiveConnection	 = oConn
						.CommandType		 = adCmdStoredProc
						.CommandText		 = "USP_Front_EShop_OrderSheet_Select_By_CartID_DelvType_ShopCD"

						.Parameters.Append .CreateParameter("@CartID",		adVarChar,	adParamInput, 20,	oRs("CartID"))
						.Parameters.Append .CreateParameter("@DelvType",	adChar,		adParamInput,  1,	oRs("DelvType"))
						.Parameters.Append .CreateParameter("@ShopCD",		adChar,		adParamInput,  6,	oRs("ShopCD"))
				END WITH
				oRs1.CursorLocation = adUseClient
				oRs1.Open oCmd, , adOpenStatic, adLockReadOnly
				SET oCmd = Nothing

				IF NOT oRs1.EOF THEN
						y = 0
						Do Until oRs1.EOF

								OrderType		= oRs1("OrderType")

								IF oRs1("SalePriceType") = "2" THEN
										SalePrice		= oRs1("EmployeeSalePrice")
										DCRate			= oRs1("EmployeeDCRate")
								ELSE
										SalePrice		= oRs1("SalePrice")
										DCRate			= oRs1("DCRate")
								END IF

								DiscountPrice	= oRs1("UseCouponPrice") + oRs1("UseScashPrice") + oRs1("UsePointPrice")
								OrderPrice		= CDbl(SalePrice) - CDbl(DiscountPrice)
								SavePoint		= Int(OrderPrice * CDbl(PointRate) / 100 + 0.5)


								IF oRs1("ProductImage_180") = "" THEN
										ProductImage	= "/Images/180_noimage.png"
								ELSE
										ProductImage	= oRs1("ProductImage_180")
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
									Tracking_ProductInfo = oRs1("SalePrice") & "|,|" & oRs1("ProductCode") & "|,|" & oRs1("ProductName") & "|,|" & oRs1("BrandName")
								Else
									Tracking_ProductInfo = Tracking_ProductInfo & "|||" & oRs1("SalePrice") & "|,|" & oRs1("ProductCode") & "|,|" & oRs1("ProductName") & "|,|" & oRs1("BrandName")
								End If
%>
                            <li>
                                <a class="listitems">
                                    <div class="thumbnail"><img src="<%=ProductImage%>" alt=""></div>
                                    <div class="itemtxt">
                                        <div class="item-info">
                                            <p class="brand-name"><%=oRs1("BrandName")%></p>
                                            <h1 class="product-name pname"><%=oRs1("ProductName")%></h1>
                                            <div class="item-option">
                                                <span>사이즈 : <%=oRs1("SizeCD")%></span>
                                                <span>수량 : <%=oRs1("OrderCnt")%></span>
                                            </div>
											<div class="price">
												<span><%=FormatNumber(SalePrice,0)%></span>원
												<%IF oRs1("SalePriceType") = "2" THEN%><em>임직원</em> <%END IF%><%IF CDbl(DCRate) > 0 OR oRs1("SalePriceType") = "2" THEN%><em><%=FormatNumber(DCRate,0)%>% 할인</em><%END IF%>
											</div>
                                            <!--<p class="pickup">매장픽업:천호점(직)</p>-->
                                        </div>
                                    </div>
									<%IF oRs("DelvType") = "S" THEN%>
                                    <span class="badge pickup">매장픽업</span>
									<%ELSE%>
                                    <span class="badge">택배배송</span>
									<%END IF%>
                                </a>
								<%
								'# 1+1 상품
								SET oCmd = Server.CreateObject("ADODB.Command")
								WITH oCmd
										.ActiveConnection	 = oConn
										.CommandType		 = adCmdStoredProc
										.CommandText		 = "USP_Front_EShop_OrderSheet_Select_For_OnePlusOne"

										.Parameters.Append .CreateParameter("@CartID",		adVarChar,	adParamInput, 20,	oRs1("CartID"))
										.Parameters.Append .CreateParameter("@GroupIdx",	adInteger,	adParamInput,   ,	oRs1("GroupIdx"))
								END WITH
								oRs2.CursorLocation = adUseClient
								oRs2.Open oCmd, , adOpenStatic, adLockReadOnly
								SET oCmd = Nothing

								IF NOT oRs2.EOF THEN
										IF oRs2("ProductImage_180") = "" THEN
												ProductImage	= "/Images/180_noimage.png"
										ELSE
												ProductImage	= oRs2("ProductImage_180")
										END IF
								%>
                                <div class="change-cont">
                                    <span class="tit">1+1상품</span><span class="cont"><%=oRs2("ProductName")%> (<%=oRs2("SIzeCD")%>)</span>
                                </div>
								<%
								END IF
								oRs2.Close
								%>
								<%IF EventProdNM <> "" THEN%>
                                <div class="change-cont">
                                    <span class="tit">사은품</span><span class="cont"><%=EventProdNM%></span>
                                </div>
								<%END IF%>
								<%IF U_MFLAG = "Y" AND oRs1("SalePriceType") <> "2" THEN%>
                                <div class="price-info">
                                    <p class="price-tit">쿠폰 적용</p>
                                    <p class="price-value"><%=FormatNumber(oRs1("UseCouponPrice"),0)%>원</p>
                                    <button type="button" onclick="openUseCoupon(<%=oRs1("Idx")%>)">조회/적용</button>
                                    <p class="price-tit">포인트 사용</p>
                                    <p class="price-value"><%=FormatNumber(oRs1("UsePointPrice"),0)%>원</p>
                                    <button type="button" onclick="openUsePoint(<%=oRs1("Idx")%>)">조회/적용</button>
                                    <p class="price-tit">슈즈상품권 사용</p>
                                    <p class="price-value"><%=FormatNumber(oRs1("UseScashPrice"),0)%>원</p>
                                    <button type="button" onclick="openUseScash(<%=oRs1("Idx")%>)">조회/적용</button>
                                </div>
                                <div class="order-price">
                                    <strong>주문금액</strong>
                                    <p><%=FormatNumber(OrderPrice,0)%>원</p>
                                </div>
								<%END IF%>

								<input type="hidden" name="DelvType"				value="<%=oRs1("DelvType")%>"	/>
								<input type="hidden" name="PickupShopCD"			value="<%=oRs1("PickupShopCD")%>"	/>
								<input type="hidden" name="ProductReceiveName"		value="<%=oRs1("ReceiveName")%>"	/>
								<input type="hidden" name="ProductReceiveTel"		value="<%=oRs1("ReceiveTel")%>"		/>
								<input type="hidden" name="ProductReceiveHP"		value="<%=oRs1("ReceiveHP")%>"		/>
								<input type="hidden" name="ProductReceiveZipCode"	value="<%=oRs1("ReceiveZipCode")%>" />
								<input type="hidden" name="ProductReceiveAddr1"		value="<%=oRs1("ReceiveAddr1")%>"	/>
								<input type="hidden" name="ProductReceiveAddr2"		value="<%=oRs1("ReceiveAddr2")%>"	/>

								<%
								'# 매장픽업
								IF oRs1("DelvType") = "S" THEN
										IF IsNull(oRs1("PickupShopCD")) OR oRs1("PickupShopCD") = "" THEN
								%>
								<div id="delivery_<%=oRs1("Idx")%>" class="accordion">
									<div class="selector">
										<button type="button" class="btn-select clickEvt" data-target="delivery_<%=oRs1("Idx")%>">
											<span>픽업매장 정보</span>
										</button>
									</div>
									<div class="option delivery-info">
										<div class="formfield">
											<div class="prev-delivery">
												<button type="button" onclick="openPickupStore(<%=oRs1("Idx")%>, '<%=oRs1("ProductCode")%>', '<%=oRs1("SizeCD")%>')">픽업매장 검색</button>
											</div>
										</div>
									</div>
								</div>
								<%
										ELSE
								%>
								<div id="delivery_<%=oRs1("Idx")%>" class="accordion">
									<div class="selector">
										<button type="button" class="btn-select clickEvt" data-target="delivery_<%=oRs1("Idx")%>">
											<span>픽업매장 정보</span>
										</button>
									</div>
									<div class="option delivery-info">
										<div class="formfield">
											<div class="prev-delivery">
												<p>픽업매장</p>
												<button type="button" onclick="openPickupStore(<%=oRs1("Idx")%>, '<%=oRs1("ProductCode")%>', '<%=oRs1("SizeCD")%>')">수정하기</button>
											</div>
											<div class="fieldset">
												<label class="fieldset-label">슈마커 <%=oRs1("PickupShopNM")%></label>
												<div class="fieldset-row">
													<p>주소 : <%=oRs1("PickupShopZipCode")%> | <%=oRs1("PickupShopAddr1")%> <%=oRs1("PickupShopAddr2")%></p>
													<p>전화 : <%=oRs1("PickupShopTel")%></p>
												</div>
											</div>
											<div class="fieldset">
												<label class="fieldset-label">받는 사람</label>
												<div class="fieldset-row">
													<p>이름 : <%=oRs1("ReceiveName")%></p>
													<p>전화 : <%=oRs1("ReceiveHP")%></p>
												</div>
											</div>
										</div>
									</div>
								</div>
								<%
										END IF
								'# 다중배송
								ELSEIF MultiDelvFlag = "Y" THEN
										IF IsNull(oRs1("ReceiveName")) OR oRs1("ReceiveName") = "" THEN
								%>
								<div id="delivery_<%=oRs1("Idx")%>" class="accordion">
									<div class="selector">
										<button type="button" class="btn-select clickEvt" data-target="delivery_<%=oRs1("Idx")%>">
											<span>배송지 정보</span>
										</button>
									</div>
									<div class="option delivery-info">
										<div class="formfield">
											<div class="prev-delivery">
												<button type="button" onclick="openMultiReceiverInfo(<%=oRs1("Idx")%>)">배송지 신규입력</button>
												<%IF U_NUM <> "" THEN%>
												<button type="button" onclick="openMyAddress(<%=oRs1("Idx")%>)">배송 목록에서 선택</button>
												<%END IF%>
											</div>
										</div>
									</div>
								</div>
								<%
										ELSE
								%>
								<div id="delivery_<%=oRs1("Idx")%>" class="accordion">
									<div class="selector">
										<button type="button" class="btn-select clickEvt" data-target="delivery_<%=oRs1("Idx")%>">
											<span>배송지 정보</span>
										</button>
									</div>
									<div class="option delivery-info">
										<div class="formfield">
											<div class="prev-delivery">
												<p>이 상품의 배송지</p>
												<button type="button" onclick="openMultiReceiverInfo(<%=oRs1("Idx")%>)">수정하기</button>
											</div>
											<div class="fieldset">
												<div class="fieldset-row">
													<p>받는분 : <%=oRs1("ReceiveName")%></p>
													<p>주소 : <%=oRs1("ReceiveZipCode")%> | <%=oRs1("ReceiveAddr1")%> <%=oRs1("ReceiveAddr2")%></p>
													<p>전화 : <%=oRs1("ReceiveTel")%></p>
													<p>휴대폰 : <%=oRs1("ReceiveHP")%></p>
												</div>
											</div>
										</div>
									</div>
								</div>
								<%
										END IF
								END IF
								%>

                            </li>
<%
								ShopOrderCnt		= ShopOrderCnt			+ CDbl(oRs1("OrderCnt"))
								ShopTagPrice		= ShopTagPrice			+ CDbl(oRs1("TagPrice"))
								ShopSalePrice		= ShopSalePrice			+ CDbl(SalePrice)
								ShopUseCouponPrice	= ShopUseCouponPrice	+ CDbl(oRs1("UseCouponPrice"))
								ShopUsePointPrice	= ShopUsePointPrice		+ CDbl(oRs1("UsePointPrice"))
								ShopUseScashPrice	= ShopUseScashPrice		+ CDbl(oRs1("UseScashPrice"))
								ShopSavePoint		= ShopSavePoint			+ CDbl(SavePoint)

								y = y + 1
								oRs1.MoveNext
						Loop 
				END IF
				oRs1.Close

				'# 일반택배배송일 경우 배송비 계산
				IF oRs("DelvType") = "P" AND ShopSalePrice < CDbl(oRs("StandardPrice")) THEN
						ShopDeliveryPrice	= CDbl(oRs("DeliveryPrice"))
				END IF
%>
<%
				'# 주문금액 계산
				TotalOrderCnt			= TotalOrderCnt			+ ShopOrderCnt
				TotalTagPrice			= TotalTagPrice			+ ShopTagPrice
				TotalSalePrice			= TotalSalePrice		+ ShopSalePrice
				TotalUseCouponPrice		= TotalUseCouponPrice	+ ShopUseCouponPrice
				TotalUsePointPrice		= TotalUsePointPrice	+ ShopUsePointPrice
				TotalUseScashPrice		= TotalUseScashPrice	+ ShopUseScashPrice
				TotalSavePoint			= TotalSavePoint		+ ShopSavePoint
				TotalDeliveryPrice		= TotalDeliveryPrice	+ ShopDeliveryPrice

				oRs.MoveNext
		Loop 
END IF
oRs.Close
%>
                        </ul>
                    </div>

			<%IF OrderType = "R" THEN%>
                    <div class="inf-type1">
                        <p class="tit">알려드립니다.</p>
                        <ul>
                            <li class="bullet-ty1">위 상품의 발송예정일은 <span class="strong">1개월 이내</span>입니다.<br>(배송일이 임박하면 문자/이메일로 안내해드립니다.)</li>
                            <li class="bullet-ty1">예약상품이 2건 이상일 경우 마지막 상품이 입고되는 날짜에 맞춰 일괄배송됩니다. 먼저 받고 싶다면 개별로 주문해주세요. </li>
                        </ul>
                    </div>
			<%END IF%>

					<script type="text/javascript">
						$(function () {
							var deliverySelectOption = function () {
								var selector,
									module;

								selector = {
									button: '.clickEvt',
									toggler: '.selector',
									panel: '.option'
								};

								module = {
									init: function () {
										$(selector.button).on('click', function () {
											module.accordion(this);
										});
										$(window).trigger('scroll');
									},
									accordion: function (el) {
										var target = $(el).data('target');

										$(selector.panel).slideUp(400);
										$(selector.toggler).removeClass('is-focus');

										if ($(selector.panel, '#' + target).css('display') === 'none') {
											$(selector.panel, '#' + target).slideDown();
											$(selector.toggler, '#' + target).addClass('is-focus');
										}
									}
								};
								module.init();
							}();
						});
					</script>

<%
Response.Write "|||||"
%>
                    <p class="cart-tit">최종 결제금액</p>
                    <div class="price">
                        <div class="price-info">
                            <div class="info-wrap">
                                <p class="price-tit">주문 상품 수</p>
                                <p class="price-value"><%=FormatNumber(TotalOrderCnt,0)%>개</p>
                            </div>
                            <div class="info-wrap">
                                <p class="price-tit">총 상품금액</p>
                                <p class="price-value"><%=FormatNumber(TotalSalePrice,0)%>원</p>
                            </div>
                            <div class="info-wrap">
                                <p class="price-tit">쿠폰 적용</p>
                                <p class="price-value"><%=FormatNumber(TotalUseCouponPrice * -1,0)%>원</p>
                            </div>
                            <div class="info-wrap">
                                <p class="price-tit">포인트 사용</p>
                                <p class="price-value"><%=FormatNumber(TotalUsePointPrice * -1,0)%>원</p>
                            </div>
                            <div class="info-wrap">
                                <p class="price-tit">슈즈상품권 사용</p>
                                <p class="price-value"><%=FormatNumber(TotalUseScashPrice * -1,0)%>원</p>
                            </div>
                            <div class="info-wrap">
                                <p class="price-tit">총 배송비</p>
                                <p class="price-value"><%=FormatNumber(TotalDeliveryPrice,0)%>원</p>
                            </div>
							<%IF U_MFLAG = "Y" THEN%>
                            <div class="info-wrap">
                                <p class="price-tit">적립포인트</p>
                                <p class="price-value"><%=FormatNumber(TotalSavePoint,0)%>원</p>
                            </div>
							<%END IF%>
                        </div>
                        <div class="order-price">
                            <strong>총 결제금액</strong>
                            <p><%=FormatNumber(TotalSalePrice - TotalUseCouponPrice - TotalUseScashPrice - TotalUsePointPrice + TotalDeliveryPrice, 0)%>원</p>
                        </div>
                    </div>|||||<%=Tracking_ProductInfo%>
<%
Set oRs2 = Nothing
Set oRs1 = Nothing
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>