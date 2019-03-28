<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'OrderReturnRequest.asp - 주문반품 요청 폼 페이지
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
DIM PayType
DIM EscrowFlag
DIM TotalSettlePrice

DIM ProductCode
DIM ProductName
DIM SizeCD
DIM BrandName
DIM OPIdx_Group
DIM GroupCnt
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
DIM DelvType
DIM DelvFee
DIM AddDeliveryPrice		: AddDeliveryPrice	= 0		'# 추가 배송비

DIM AddProductCode
DIM AddProductName
DIM AddSizeCD
DIM AddOrderCnt
DIM AddBrandName
DIM AddOrderState
DIM AddCancelState1
DIM AddCancelState2
DIM AddProductImage

DIM arrHP1

DIM RefundBankCode
DIM RefundBankName
DIM RefundAccountNum
DIM RefundAccountName
DIM RefundPhone1
DIM RefundPhone2
DIM RefundPhone3
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


OrderCode			= sqlFilter(Request("OrderCode"))
OPIdx				= sqlFilter(Request("OPIdx"))




IF OrderCode = "" OR OPIdx = "" THEN
		Response.Write "FAIL|||||선택한 주문정보가 없습니다."
		Response.End
END IF


arrHP1		= ARRAY("010", "011", "016", "017", "018", "019")


SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성


'# 주문정보 체크
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Order_Select_For_OrderInfo"

		.Parameters.Append .CreateParameter("@OrderCode",	adVarchar, adParaminput,	20,		OrderCode)
		.Parameters.Append .CreateParameter("@UserID",		adVarchar, adParaminput,	20,		U_NUM)
		.Parameters.Append .CreateParameter("@OrderName",	adVarChar, adParamInput,	50,		N_NAME)
		.Parameters.Append .CreateParameter("@OrderHp",		adVarChar, adParamInput,	20,		N_HP)
		.Parameters.Append .CreateParameter("@OrderEmail",	adVarChar, adParamInput,	50,		N_EMAIL)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		OrderDate			= oRs("OrderDate")
		PayType				= oRs("PayType")
		EscrowFlag			= oRs("EscrowFlag")
		TotalSettlePrice	= oRs("OrderPrice") + oRs("DeliveryPrice")

ELSE
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||반품할 주문내역이 없습니다."
		Response.End
END IF
oRs.Close


'-----------------------------------------------------------------------------------------------------------'
'# 주문정보 Start
'-----------------------------------------------------------------------------------------------------------'
wQuery = ""
wQuery = wQuery & "WHERE A.IsShowFlag = 'Y' AND A.ProductType = 'P' "
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
				Response.Write "FAIL|||||반품신청이 불가능한 상태의 주문 입니다."
				Response.End
		END IF

		ProductCode		= oRs("ProductCode")
		ProductName		= oRs("ProductName")
		SizeCD			= oRs("SizeCD")
		BrandName		= oRs("BrandName")
		OPIdx_Group		= oRs("OPIdx_Group")
		GroupCnt		= oRs("GroupCnt")
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

		DelvType		= oRs("DelvType")

		DelvFee			= oRs("VendorDeliveryPrice")		'# 반품 배송비

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

'-----------------------------------------------------------------------------------------------------------'
'# 1+1 주문정보 Start
'-----------------------------------------------------------------------------------------------------------'
IF CInt(GroupCnt) > 1 THEN
		wQuery = ""
		wQuery = wQuery & "WHERE A.IsShowFlag = 'Y' AND A.ProductType = 'O' "
		wQuery = wQuery & "AND A.OrderCode = '" & OrderCode & "' "
		wQuery = wQuery & "AND A.OPIdx_Group = " & OPIdx_Group & " "
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
						Response.Write "FAIL|||||1+1 상품이 반품신청이 불가능한 상태 입니다."
						Response.End
				END IF

				AddProductCode		= oRs("ProductCode")
				AddProductName		= oRs("ProductName")
				AddSizeCD			= oRs("SizeCD")
				AddOrderCnt			= oRs("OrderCnt")
				AddBrandName		= oRs("BrandName")
				AddOrderState		= oRs("OrderState")
				AddCancelState1		= oRs("CancelState1")
				AddCancelState2		= oRs("CancelState2")
		
				IF oRs("ProductImage_180") = "" THEN
						AddProductImage	= "/Images/180_noimage.png"
				ELSE
						AddProductImage	= oRs("ProductImage_180")
				END IF
		ELSE
				Response.Write "FAIL|||||1+1 상품의 주문정보가 없습니다."
				Response.End
		END IF
		oRs.Close
END IF
'-----------------------------------------------------------------------------------------------------------'
'# 1+1 주문정보 End
'-----------------------------------------------------------------------------------------------------------'

'-----------------------------------------------------------------------------------------------------------'
'# 주문 추가배송비 계산 Start
'-----------------------------------------------------------------------------------------------------------'
'# 일반택배 주문일 경우만 추가배송비 계산
IF DelvType = "P" THEN
		wQuery = ""
		wQuery = wQuery & "WHERE A.IsShowFlag = 'Y' AND A.ProductType = 'P' "
		wQuery = wQuery & "AND A.OrderCode = '" & OrderCode & "' "
		wQuery = wQuery & "AND A.Idx = " & OPIdx & " "
		IF U_NUM <> "" THEN
				wQuery = wQuery & "AND B.UserID = '" & U_NUM & "' "
		ELSE
				wQuery = wQuery & "AND B.OrderName = '" & N_NAME & "' AND B.OrderHp = '" & N_HP & "' AND B.OrderEmail = '" & N_EMAIL & "' "
		END IF


		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Order_Product_Select_For_Vendor_TotalPrice"

				.Parameters.Append .CreateParameter("@WQUERY",		adVarchar, adParaminput,	1000,	wQuery)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		SET oCmd = Nothing

		IF NOT oRs.EOF THEN
				'# 택배주문수량이 여러개 이거나 배송비 결제내역이 없으면 추가배송비 산정
				IF oRs("OrderCnt_P") > 1 OR oRs("DeliveryPrice") = 0 THEN
						AddDeliveryPrice	= DelvFee
				END IF
		END IF
		oRs.Close
ELSE
		'# 매장픽업 주문일 경우는 추가배송비 없음
		AddDeliveryPrice	= 0
END IF
'-----------------------------------------------------------------------------------------------------------'
'# 주문 추가배송비 계산 체크 End
'-----------------------------------------------------------------------------------------------------------'



Response.Write "OK|||||"
%>					
        <div class="area-dim"></div>

        <div class="area-pop" id="OrderReturn">
            <div class="full">
                <div class="tit-pop">
                    <div class="tit">반품 신청하기</div>
                    <button type="button" onclick="closePop('DimDepth1')" class="btn-hide-pop">닫기</button>
                </div>

                <div class="container-pop mypage-ty2">
                    <div class="contents">
                        <div class="wrap-order">
                            <div class="h-line">
                                <h3 class="h-level4">신청할 주문</h3>
                            </div>

							<form name="OrderChangeReturnForm" method="post" action="OrderChangeReturnRequestOk.asp">
								<input type="hidden" name="OrderCode"			value="<%=OrderCode%>" />
								<input type="hidden" name="OPIdx"				value="<%=OPIdx%>" />
								<input type="hidden" name="CancelType"			value="R" />
								<input type="hidden" name="ProductCode"			value="<%=ProductCode%>" />
								<input type="hidden" name="SizeCD"				value="<%=SizeCD%>" />
								<input type="hidden" name="DeliveryCouponIdx"	value="" />
								<input type="hidden" name="PayType"				value="<%=PayType%>" />
								<input type="hidden" name="EscrowFlag"			value="<%=EscrowFlag%>" />
								<input type="hidden" name="TotalSettlePrice"	value="<%=TotalSettlePrice%>" />

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
													<span class="item-code">FS3SCA5331X-BLK FS3SCA5331X-BLK</span>
												</span>
												<span class="product-name"><em><%=ProductName%></em></span>
												
												<span class="inform">
													<span class="list">
														<span class="tit">옵션</span>
														<span class="opt"><%=SizeCD%></span>
													</span>
													<span class="list">
														<span class="tit">수량</span>
														<span class="opt"><%=OrderCnt %></span>
													</span>
													<span class="list">
														<span class="tit">결제금액</span>
														<span class="opt price"><em><%=FormatNumber(OrderPrice,0)%></em>원</span>
													</span>
													<span class="list">
														<span class="tit">결제수단</span>
														<span class="opt"><%=GetPayType(PayType)%></span>
													</span>
												</span>
											</span>
										</span>
									</a>
                                </div>
								<%IF CInt(GroupCnt) > 1 AND AddProductCode <> "" THEN%>
                                <div class="informItem">
                                    <a href="#">
										<span class="cont">
											<span class="thumbNail">
												<span class="img">
													<img src="<%=AddProductImage%>" alt="상품 이미지">
												</span>
												<span class="about">
													<span class="process"><%=GetOrderState(AddOrderState, AddCancelState1, AddCancelState2)%></span>
												</span>
											</span>
											
											<span class="detail">
												<span class="brand">
													<span class="name"><%=AddBrandName%></span>
													<span class="oneplusone">[1+1]</span>
												</span>
												<span class="product-name"><em><%=AddProductName%></em></span>
												
												<span class="inform">
													<span class="list">
														<span class="tit">옵션</span>
														<span class="opt"><%=AddSizeCD%></span>
													</span>
													<span class="list">
														<span class="tit">수량</span>
														<span class="opt"><%=AddOrderCnt %></span>
													</span>
												</span>
											</span>
										</span>
									</a>
                                </div>
								<%END IF%>
                            </div>

                            <div class="h-line">
                                <h3 class="h-level4">사유 선택</h3>
                            </div>

                            <div class="reason">
                                <span class="select">
									<select name="ReasonType" title="취소 사유 선택">
										<option value="">사유선택</option>
									<%
									'# 반품사유
									SET oCmd = Server.CreateObject("ADODB.Command")
									WITH oCmd
											.ActiveConnection	 = oConn
											.CommandType		 = adCmdStoredProc
											.CommandText		 = "USP_Front_EShop_Order_Product_Cancel_ReasonType_Select"

											.Parameters.Append .CreateParameter("@CancelType",		 adChar, adParaminput, 1,	"R")
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
									<input type="text" name="Memo" placeholder="사유를 직접 입력하세요.">
								</span>
                            </div>

							<input type="hidden" name="ReturnName"		value="<%=ReceiveName%>"	/>
							<input type="hidden" name="ReturnHp"		value="<%=ReceiveHp%>"		/>
							<input type="hidden" name="ReturnZipCode"	value="<%=ReceiveZipCode%>" />
							<input type="hidden" name="ReturnAddr1"		value="<%=ReceiveAddr1%>"	/>
							<input type="hidden" name="ReturnAddr2"		value="<%=ReceiveAddr2%>"	/>

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
                                <h3 class="h-level4">반품 배송비 결제방법</h3>
                            </div>

                            <div class="delivery-price">
                                <div class="list">
                                    <div class="tit">
                                        <span>반품 택배 배송비 <em class="ty-red"><%=FormatNumber(DelvFee + AddDeliveryPrice,0)%>원</em>을</span>
                                    </div>
                                </div>
                            </div>

                            <div class="area-radio">
								<%IF EscrowFlag <> "Y" AND OrderPrice >= (DelvFee + AddDeliveryPrice) THEN%>
                                <span class="rad-ty1">
									<input type="radio" name="DelvFeeType" id="DelvFeeType_5" value="5" checked="checked" />
									<label for="DelvFeeType_5">환불금액 차감</label>
								</span>
								<%END IF%>
                                <span class="rad-ty1">
									<input type="radio" name="DelvFeeType" id="DelvFeeType_6" value="6" <%IF EscrowFlag = "Y" OR OrderPrice < (DelvFee + AddDeliveryPrice) THEN%>checked="checked"<%END IF%> />
									<label for="DelvFeeType_6">신용카드 결제</label>
								</span>
								<%IF U_MFLAG = "Y" THEN%>
                                <span class="rad-ty1">
									<input type="radio" name="DelvFeeType" id="DelvFeeType_7" value="7" />
									<label for="DelvFeeType_7">반품쿠폰 사용</label>
								</span>
								<%END IF%>
								<%IF OrderPrice < (DelvFee + AddDeliveryPrice) THEN%>
                                <span class="rad-ty1">
									<input type="radio" name="DelvFeeType" id="DelvFeeType_2" value="2" />
									<label for="DelvFeeType_2">동봉</label>
								</span>
								<%END IF%>
                                <span class="rad-ty1">
									<input type="radio" name="DelvFeeType" id="DelvFeeType_3" value="3" />
									<label for="DelvFeeType_3">계좌이체</label>
								</span>
                            </div>

                            <div class="h-line">
                                <h3 class="h-level4">환불 금액 확인</h3>
                            </div>

                            <ul class="detailView on-right" id="RefundInfo">
                            </ul>

							<div class="refundaccount-info" style="display:none;">
                            <div class="h-line">
                                <h3 class="h-level4">환불 계좌 정보</h3>
                            </div>
                            <ul class="detailView">
                                <li class="detailList">
									<fieldset>
										<div class="fieldset">
											<label for="RefundBankCode" class="fieldset-label">환불 은행</label>
											<div class="fieldset-row">
												<span class="select">
													<select name="RefundBankCode" id="RefundBankCode">
														<option value="">은행선택</option>
													<%
													'# 환불은행코드
													SET oCmd = Server.CreateObject("ADODB.Command")
													WITH oCmd
															.ActiveConnection	 = oConn
															.CommandType		 = adCmdStoredProc
															.CommandText		 = "USP_Admin_EShop_RefundBank_Select_For_Use"
													END WITH
													oRs.CursorLocation = adUseClient
													oRs.Open oCmd, , adOpenStatic, adLockReadOnly
													SET oCmd = Nothing

													IF NOT oRs.EOF THEN
															Do Until oRs.EOF
																	IF oRs("BankCode") = RefundBankCode THEN
																			RefundBankName = oRs("BankName")
													%>
														<option value="<%=oRs("BankCode")%>" selected="selected"><%=oRs("BankName")%></option>
													<%
																	ELSE
													%>
														<option value="<%=oRs("BankCode")%>"><%=oRs("BankName")%></option>
													<%
																	END IF

																	oRs.MoveNext
															Loop
													END IF
													oRs.Close
													%>
													</select>
													<span class="value"><%=RefundBankName%></span>
												</span>
											</div>
										</div>
										<div class="fieldset">
											<label for="RefundAccountNum" class="fieldset-label">환불 계좌</label>
											<div class="fieldset-row">
												<span class="input is-expand">
													<input type="text" name="RefundAccountNum" id="RefundAccountNum" />
												</span>
											</div>
										</div>
										<div class="fieldset">
											<label for="RefundAccountName" class="fieldset-label">예금주명</label>
											<div class="fieldset-row">
												<span class="input is-expand">
													<input type="text" name="RefundAccountName" id="RefundAccountName" />
												</span>
											</div>
										</div>
										<div class="fieldset ty-col2 pt0">
											<label for="RefundPhone23" class="fieldset-label">휴대폰 번호</label>
											<div class="fieldset-row">
												<span class="select">
													<select name="RefundPhone1" id="RefundPhone1">
														<option value="">선택</option>
														<%FOR i = 0 TO UBOUND(arrHP1)%>
														<option value="<%=arrHP1(i)%>"<%IF arrHP1(i) = RefundPhone1 THEN%> selected="selected"<%END IF%>><%=arrHP1(i)%></option>
														<%NEXT%>
													</select>
													<span class="value"><%=RefundPhone1%></span>
												</span>
												<span class="input">
													<input type="text" name="RefundPhone23" id="RefundPhone23" value="<%=RefundPhone2 & RefundPhone3%>" title="휴대폰번호의 앞 번호와 뒷 번호 입력" />
												</span>
											</div>
										</div>
									</fieldset>
								</li>
							</ul>
							</div>

                            <div class="h-line">
                                <h3 class="h-level4">반품 절차 안내</h3>
                            </div>

                            <div class="process-info">
                                <ol class="return">
                                    <li>반품<br>접수</li>
                                    <li>제품<br>회수</li>
                                    <li>검수<br>승인</li>
                                    <li>환불<br>완료</li>
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
                        <button type="button" onclick="orderChangeReturnCheck('R')" class="button ty-red">반품 신청 하기</button>
                    </div>
                </div>
            </div>
        </div>

		<script type="text/javascript">
			$(function () {
				$("#OrderReturn select").on("change", function () {
					$(this).parent().find('.value').text($('option:selected', $(this)).text());
				});
			})
		</script>

<%
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>