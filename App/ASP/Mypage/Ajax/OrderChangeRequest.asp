<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'OrderChangeRequest.asp - 주문교환 요청 폼 페이지
'Date		: 2019.01.02
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
IF U_NUM = "" AND N_NAME = "" THEN
		Response.Write "LOGIN|||||로그인 정보가 없습니다. 다시 로그인하여 주십시오."
		Response.End
END IF

'/****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn											'# ADODB Connection 개체
DIM oRs												'# ADODB Recordset 개체
DIM oRs1											'# ADODB Recordset 개체
DIM oCmd											'# ADODB Command 개체

DIM i
DIM j
DiM x
DIM y

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절

DIM OrderCode
DIM OPIdx

DIM OrderDate
DIM ProductCode
DIM ProductName
DIM SizeCD
DIM BrandName
DIM OrderCnt
DIM OrderPrice
DIM OrderState
DIM CancelState1
DIM CancelState2
DIM ProductImage
DIM ReceiveName
DIM ReceiveHp
DIM ReceiveZipCode
DIM ReceiveAddr1
DIM ReceiveAddr2
DIM DelvFee

Dim ProductStockCnt
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


OrderCode			= sqlFilter(Request("OrderCode"))
OPIdx				= sqlFilter(Request("OPIdx"))




IF OrderCode = "" OR OPIdx = "" THEN
		Response.Write "FAIL|||||선택한 주문정보가 없습니다."
		Response.End
END IF



SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성


'-----------------------------------------------------------------------------------------------------------'
'# 주문정보 Start
'-----------------------------------------------------------------------------------------------------------'
wQuery = ""
wQuery = wQuery & "WHERE A.IsShowFlag = 'Y' AND A.ProductType IN ('P','O') "
wQuery = wQuery & "AND A.OrderCode = '" & OrderCode & "' "
wQuery = wQuery & "AND A.Idx = " & OPIdx & " "
IF U_NUM <> "" THEN
		wQuery = wQuery & "AND B.UserID = '" & U_NUM & "' "
ELSE
		wQuery = wQuery & "AND B.OrderName = '" & N_NAME & "' AND B.OrderHp = '" & N_HP & "' AND B.OrderEmail = '" & N_EMAIL & "' "
END IF

sQuery = ""

SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Order_Product_Select_For_Order_Detail"

		.Parameters.Append .CreateParameter("@WQuery",		 adVarchar, adParaminput, 1000	, wQuery)
		.Parameters.Append .CreateParameter("@SQuery",		 adVarchar, adParaminput, 100	, sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		IF oRs("OrderState") <> "5" OR oRs("CancelState2") <> "0" THEN
				Response.Write "FAIL|||||교환신청이 불가능한 상태의 주문 입니다."
				Response.End
		END IF

		OrderDate		= oRs("OrderDate")
		ProductCode		= oRs("ProductCode")
		ProductName		= oRs("ProductName")
		SizeCD			= oRs("SizeCD")
		BrandName		= oRs("BrandName")
		OrderCnt		= oRs("OrderCnt")
		OrderPrice		= oRs("OrderPrice")
		OrderState		= oRs("OrderState")
		CancelState1	= oRs("CancelState1")
		CancelState2	= oRs("CancelState2")
		
		ReceiveName		= oRs("ReceiveName")
		ReceiveHp		= oRs("ReceiveHp")
		ReceiveZipCode	= oRs("ReceiveZipCode")
		ReceiveAddr1	= oRs("ReceiveAddr1")
		ReceiveAddr2	= oRs("ReceiveAddr2")

		DelvFee			= oRs("VendorDeliveryPrice") * 2		'# 업체 배송비 * 왕복

		IF oRs("ProductImage_180") = "" THEN
				ProductImage	= "/Images/180_noimage.png"
		ELSE
				ProductImage	= oRs("ProductImage_180")
		END IF
ELSE
		Response.Write "FAIL|||||주문정보가 없습니다."
		Response.End
END IF
oRs.Close
'-----------------------------------------------------------------------------------------------------------'
'# 주문정보 End
'-----------------------------------------------------------------------------------------------------------'

SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Product_Select_For_Available_Check"

		.Parameters.Append .CreateParameter("@ProductCode", adInteger, adParamInput, , ProductCode)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF oRs.BOF OR oRs.EOF THEN
		ProductStockCnt = 0
ELSE
		ProductStockCnt = oRs("StockCnt")
END IF
oRs.Close

IF ProductStockCnt <= 0 Then
	Response.Write "FAIL|||||해당 제품의 재고가 없습니다. 반품신청을 해 주시기 바랍니다."
	Response.End
END IF


Response.Write "OK|||||"
%>					
        <div class="area-dim"></div>

        <div class="area-pop" id="OrderChange">
            <div class="full">
                <div class="tit-pop">
                    <div class="tit">교환 신청하기</div>
                    <button type="button" onclick="closePop('DimDepth1')" class="btn-hide-pop">닫기</button>
                </div>

                <div class="container-pop mypage-ty2">
                    <!-- 팝업 스타일 변경으로 'mypage-ty2'클래스 명 추가 -->
                    <div class="contents">
                        <div class="wrap-order">
                            <div class="h-line">
                                <h3 class="h-level4">신청할 주문</h3>
                            </div>

							<form name="OrderChangeReturnForm" method="post" action="OrderChangeReturnRequestOk.asp">
								<input type="hidden" name="OrderCode"			value="<%=OrderCode%>" />
								<input type="hidden" name="OPIdx"				value="<%=OPIdx%>" />
								<input type="hidden" name="CancelType"			value="X" />
								<input type="hidden" name="ProductCode"			value="<%=ProductCode%>" />
								<input type="hidden" name="SizeCD"				value="<%=SizeCD%>" />
								<input type="hidden" name="DeliveryCouponIdx"	value="" />

                            <div class="informView">
                                <div class="informItem">
                                    <a href="#">
										<span class="head-tit">
											<span class="tit">주문번호 : <%=OrderCode%></span>
											<span class="date"><%=GetDateYmd(OrderDate)%></span>
										</span>
										<span class="cont">
											<span class="thumbNail">
												<span class="img">
													<img src="<%=ProductImage%>" alt="상품 이미지">
												</span>
												<span class="about">
													<span class="process"><%=GetOrderState(OrderState, CancelState1, CancelState2)%></span>
												</span>
											</span>
											
											<span class="detail">
												<span class="brand">
													<span class="name"><%=BrandName%></span>
												</span>
												<span class="product-name"><em><%=ProductName%></em></span>
												
												<span class="inform">
													<span class="list">
														<span class="tit">사이즈</span>
														<span class="opt"><%=SizeCD%></span>
													</span>
													<span class="list">
														<span class="tit">수량</span>
														<span class="opt"><%=OrderCnt%></span>
													</span>
													<span class="list">
														<span class="tit">결제금액</span>
														<span class="opt price"><em><%=FormatNumber(OrderPrice,0)%></em>원</span>
													</span>
												</span>
											</span>
										</span>
									</a>
                                </div>
                            </div>

                            <div class="h-line">
                                <h3 class="h-level4">교환 사이즈 선택</h3>
                            </div>

                            <div class="reason">
                                <span class="select">
									<select name="ChgSizeCD" title="교환 사이즈 선택">
										<option value="">사이즈 선택</option>
										<%
										SET oCmd = Server.CreateObject("ADODB.Command")
										WITH oCmd
												.ActiveConnection	 = oConn
												.CommandType		 = adCmdStoredProc
												'# .CommandText		 = "USP_Admin_EShop_Stock_Select_For_SizeList_By_ProductCode"
												.CommandText		 = "USP_Front_EShop_Product_SizeCD_Select_With_EShop_Stock"
		
												.Parameters.Append .CreateParameter("@ProductCode",			 adInteger,	 adParamInput,  ,	 ProductCode)
										END WITH
										oRs.CursorLocation = adUseClient
										oRs.Open oCmd, , adOpenStatic, adLockReadOnly
										SET oCmd = Nothing

										IF NOT oRs.EOF THEN
												Do Until oRs.EOF
														IF oRs("StockCnt") > 0 THEN
										%>
										<option value="<%=oRs("SizeCD")%>"><%=oRs("SizeCD")%></option>
										<%
														ELSE
										%>
										<option value="<%=oRs("SizeCD")%>" disabled><%=oRs("SizeCD")%>&nbsp;[품절]</option>
										<%
														END IF

														oRs.MoveNext
												Loop
										END IF
										oRs.Close
										%>
									</select>
									<span class="value">사이즈 선택</span>
                                </span>
                            </div>

                            <div class="h-line">
                                <h3 class="h-level4">사유 선택</h3>
                            </div>

                            <div class="reason">
                                <span class="select">
									<select name="ReasonType" title="교환 사유 선택" onchange="chgReasonType()">
										<option value="">사유선택</option>
										<%
										'# 교환사유
										SET oCmd = Server.CreateObject("ADODB.Command")
										WITH oCmd
												.ActiveConnection	 = oConn
												.CommandType		 = adCmdStoredProc
												.CommandText		 = "USP_Front_EShop_Order_Product_Cancel_ReasonType_Select"

												.Parameters.Append .CreateParameter("@CancelType",		 adChar, adParaminput, 1,	"X")
										END WITH
										oRs.CursorLocation = adUseClient
										oRs.Open oCmd, , adOpenStatic, adLockReadOnly
										SET oCmd = Nothing

										IF NOT oRs.EOF THEN
												Do Until oRs.EOF
										%>
											<option value="<%=oRs("ReasonType")%>"><%=oRs("ReasonName")%></option>
										<%
														oRs.MoveNext
												Loop
										END IF
										oRs.Close
										%>
									</select>
									<span class="value">사유선택</span>
                                </span>

                                <span class="input">
									<input type="text" name="Memo" placeholder="사유를 직접 입력하세요." />
								</span>
                            </div>

							<input type="hidden" name="ReturnName"		value="<%=ReceiveName%>"	/>
							<input type="hidden" name="ReturnHp"		value="<%=ReceiveHp%>"		/>
							<input type="hidden" name="ReturnZipCode"	value="<%=ReceiveZipCode%>" />
							<input type="hidden" name="ReturnAddr1"		value="<%=ReceiveAddr1%>"	/>
							<input type="hidden" name="ReturnAddr2"		value="<%=ReceiveAddr2%>"	/>
							<input type="hidden" name="ReceiveName"		value="<%=ReceiveName%>"	/>
							<input type="hidden" name="ReceiveHp"		value="<%=ReceiveHp%>"		/>
							<input type="hidden" name="ReceiveZipCode"	value="<%=ReceiveZipCode%>" />
							<input type="hidden" name="ReceiveAddr1"	value="<%=ReceiveAddr1%>"	/>
							<input type="hidden" name="ReceiveAddr2"	value="<%=ReceiveAddr2%>"	/>

                            <div class="h-line">
                                <h3 class="h-level4">상품 수거지</h3>
                                <a href="javascript:openZipCodeSearch('Return');" class="all-view is-right">변경</a>
                            </div>

                            <div class="addr-list">
                                <div class="list">
                                    <div class="tit ReturnName">
                                        <span><%=ReceiveName%></span><span><%=ReceiveHp%></span>
                                    </div>
                                    <div class="address ReturnAddr">[<%=ReceiveZipCode%>] <%=ReceiveAddr1%> <%=ReceiveAddr2%></div>
                                </div>
                            </div>

                            <div class="h-line">
                                <h3 class="h-level4">상품 수령지</h3>
                                <a href="javascript:openZipCodeSearch('Receive');" class="all-view is-right">변경</a>
                            </div>

                            <div class="addr-list">
                                <div class="list">
                                    <div class="tit ReceiveName">
                                        <span><%=ReceiveName%></span><span><%=ReceiveHp%></span>
                                    </div>
                                    <div class="address ReceiveAddr">[<%=ReceiveZipCode%>] <%=ReceiveAddr1%> <%=ReceiveAddr2%></div>
                                </div>
                            </div>

                            <div class="h-line">
                                <h3 class="h-level4">교환 배송비 결제방법</h3>
                            </div>

                            <div class="delivery-price">
                                <div class="list">
                                    <div class="tit">
                                        <span>교환 택배 배송비 <em class="ty-red"><%=FormatNumber(DelvFee,0)%>원</em>을</span>
                                    </div>
                                </div>
                            </div>

                            <div class="area-radio">
                                <span class="rad-ty1">
									<input type="radio" name="DelvFeeType" id="DelvFeeType_6" value="6" checked="checked" />
									<label for="DelvFeeType_6">신용카드 결제</label>
								</span>
								<%IF U_MFLAG = "Y" THEN%>
                                <span class="rad-ty1">
									<input type="radio" name="DelvFeeType" id="DelvFeeType_7" value="7" />
									<label for="DelvFeeType_7">교환배송비 쿠폰</label>
								</span>
								<%END IF%>
                                <span class="rad-ty1">
									<input type="radio" name="DelvFeeType" id="DelvFeeType_2" value="2" />
									<label for="DelvFeeType_2">동봉</label>
								</span>
                                <span class="rad-ty1">
									<input type="radio" name="DelvFeeType" id="DelvFeeType_3" value="3" />
									<label for="DelvFeeType_3">계좌이체</label>
								</span>
                                <span class="rad-ty1" style="display:none">
									<input type="radio" name="DelvFeeType" id="DelvFeeType_1" value="1" />
									<label for="DelvFeeType_1">슈마커부담</label>
								</span>
                            </div>

                            <div class="h-line">
                                <h3 class="h-level4">교환 절차 안내</h3>
                            </div>

                            <div class="process-info">
                                <ol class="exchange">
                                    <li>반품<br>접수</li>
                                    <li>제품<br>회수</li>
                                    <li>검수<br>승인</li>
                                    <li>교환<br>발송</li>
                                    <li>교환<br>완료</li>
                                </ol>
                            </div>

                            <div class="inf-type1">
                                <p class="tit">알려드립니다.</p>
                                <ul>
                                    <li class="bullet-ty1">고객변심으로 인한 교환/반품은 상품을 수령하신 날로부터 7일 이내에 신청 가능합니다.</li>
                                    <li class="bullet-ty1">받으신 상품의 내용이 표시된 광고와 다른 경우는 상품수령일로 부터 3개월 이내에 신청 가능합니다.</li>
                                    <li class="bullet-ty1">배송완료 후 7일이 경과하였거나 물품하자/오배송으로 인한 교환/반품은 1:1 상담으로 문의해주시기 바랍니다.</li>
                                    <li class="bullet-ty1">교환 및 환불 배송비는 상품하자나 불량으로 귀책사유가 당사에 있을 경우 회수 배송비는 당사가 부담하지만 단순히 고객변심으로 교환 및 환불 하는 경우는 고객님께서 부담하셔야 합니다.</li>
                                </ul>
                            </div>

							</form>

                        </div>
                    </div>

                    <div class="btns">
                        <button type="button" onclick="orderChangeReturnCheck('X')" class="button ty-red">교환 신청 하기</button>
                    </div>
                </div>
            </div>
        </div>

		<script type="text/javascript">
			$(function () {
				$("#OrderChange select").on("change", function () {
					$(this).parent().find('.value').text($('option:selected', $(this)).text());
				});
			});
		</script>

<%
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>