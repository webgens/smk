<%@ Language=VBScript codepage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'OrderCRXList.asp - 주문 취소/교환/반품 리스트
'Date		: 2019.01.03
'Update		: 
'*****************************************************************************************'
	
'# 페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include Virtual = "/ADO/ADODBCommon.asp" -->
<!-- #include Virtual = "/Common/Common.asp" -->
<!-- #include Virtual = "/Common/SetInfo.asp" -->

<%
IF U_NUM = "" AND N_NAME = "" THEN
		Response.Write "LOGIN|||||로그인 정보가 없습니다. 다시 로그인하여 주십시오."
		Response.End
END IF

'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn							'# ADODB Connection 개체
DIM oRs								'# ADODB Recordset 개체
DIM oRs1							'# ADODB Recordset 개체
DIM oCmd							'# ADODB Command 개체

DIM wQuery							'# WHERE 절
DIM sQuery							'# SORT 절

DIM x
DIM y

DIM SDate
DIM EDate
DIM SCancelType

DIM OrderStateNM
DIM ProductImage
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'
	
	
SDate			 = sqlFilter(Request("SDate"))
EDate			 = sqlFilter(Request("EDate"))
SCancelType		 = sqlFilter(Request("SCancelType"))


SET oConn				 = ConnectionOpen()							'# 커넥션 생성
SET oRs					 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성




Response.Write "OK|||||"
%>
                                        <div class="h-line">
                                            <h2 class="h-level4">신청/처리 내역</h2>
                                            <span class="h-date is-right"><%=SDate%> ~ <%=EDate%></span>
                                        </div>

                                        <div class="ly-accord-sub">
<%
'# 취소 신청내역
wQuery = "WHERE B.IsShowFlag = 'Y' AND B.ProductType = 'P' AND (B.OrderState = 'C' OR B.CancelRequestFlag = 'C') "
IF SDate <> "" THEN
		wQuery = wQuery & "AND A.OrderDate >= '" & Replace(SDate, "-", "") & "' "
END IF
IF EDate <> "" THEN
		wQuery = wQuery & "AND A.OrderDate <= '" & Replace(EDate, "-", "") & "' "
END IF
IF U_NUM <> "" THEN
		wQuery = wQuery & "AND A.UserID = '" & U_NUM & "' "
ELSEIF N_NAME <> "" THEN
		wQuery = wQuery & "AND (A.UserID = '' OR A.UserID IS NULL) AND A.OrderName = '" & N_NAME & "' AND A.OrderHp = '" & N_HP & "' AND A.OrderEmail = '" & N_EMAIL & "' "
END IF

sQuery = "ORDER BY A.OrderCode DESC, B.Idx "



SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Mobile_EShop_Order_Product_Select"

		.Parameters.Append .CreateParameter("@wQuery",		 adVarchar, adParaminput, 1000	, wQuery)
		.Parameters.Append .CreateParameter("@sQuery",		 adVarchar, adParaminput, 100	, sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing
%>
                                            <!-- 주문취소 신청 내역 -->
                                            <div id="processIssues_cancel" class="accord-sub-mypage">
                                                <div class="ly-title_sub">
                                                    <button type="button" class="btn-list clickEvt_sub" data-target="processIssues_cancel">주문취소 신청 (<%=FormatNumber(oRs.RecordCount, 0)%>)</button>
                                                </div>
                                                <div class="ly-content_sub">
<%
IF NOT oRs.EOF THEN	
%>
                                                    <ul class="informView">
<%
		Do Until oRs.EOF
				OrderStateNM	= GetOrderState(oRs("OrderState"), oRs("CancelState1"), oRs("CancelState2"))

				IF oRs("ProductImage") = "" THEN
						ProductImage	= "/Images/180_noimage.png"
				ELSE
						ProductImage	= oRs("ProductImage")
				END IF
%>
                                                        <li class="informItem">
                                                            <a href="/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>">
															<span class="cont">
																<span class="thumbNail">
																	<span class="img">
																		<img src="<%=ProductImage%>" alt="상품 이미지">
																	</span>
																	<span class="about">
																		<span class="process"><%=OrderStateNM%></span>
																	</span>
																</span>
							
																<span class="detail">
																	<span class="brand">
																		<span class="name"><%=oRs("BrandName")%> </span>
																		<%IF oRs("ProductType") = "O" THEN%><span class="oneplusone">[1+1]</span><%END IF%>
																	</span>
																	<span class="product-name"><em><%=oRs("ProductName")%></em></span>
																	
																	<span class="inform">
																		<span class="list">
																			<span class="tit">옵션</span>
																			<span class="opt"><%=oRs("SizeCD") %></span>
																		</span>
																		<span class="list">
																			<span class="tit">수량</span>
																			<span class="opt"><%=oRs("OrderCnt") %></span>
																		</span>
																		<span class="list">
																			<span class="tit">결제금액</span>
																			<span class="opt price"><em><%=FormatNumber(oRs("OrderPrice"),0)%></em>원</span>
																		</span>
																		<span class="list">
																			<span class="tit">배송/결제</span>
																			<span class="opt"><%IF oRs("DelvType") = "S" THEN%>매장픽업<%ELSE%>택배<%END IF%> / <%=GetPayType(oRs("PayType"))%></span>
																		</span>
																	</span>
																</span>
															</span>
															</a>

                                                            <div class="buttongroup">
                                                                <button type="button" onclick="getOrderDetail('<%=oRs("Idx")%>')" class="button-ty2 is-expand ty-bd-gray">신청내용 확인</button>
                                                            </div>
                                                        </li>
<%
				oRs.MoveNext
		Loop
%>
                                                    </ul>
<%
ELSE
%>
													<div class="no-history">
														<p>신청 내역이 없습니다.</p>
													</div>
<%
END IF
oRs.Close
%>
                                                </div>
                                            </div>
                                            <!-- // 주문취소 신청 내역 -->

<%
'# 반품 신청내역
wQuery = "WHERE B.IsShowFlag = 'Y' AND B.ProductType = 'P' AND B.CancelRequestFlag = 'R' "
IF SDate <> "" THEN
		wQuery = wQuery & "AND A.OrderDate >= '" & Replace(SDate, "-", "") & "' "
END IF
IF EDate <> "" THEN
		wQuery = wQuery & "AND A.OrderDate <= '" & Replace(EDate, "-", "") & "' "
END IF
IF U_NUM <> "" THEN
		wQuery = wQuery & "AND A.UserID = '" & U_NUM & "' "
ELSEIF N_NAME <> "" THEN
		wQuery = wQuery & "AND (A.UserID = '' OR A.UserID IS NULL) AND A.OrderName = '" & N_NAME & "' AND A.OrderHp = '" & N_HP & "' AND A.OrderEmail = '" & N_EMAIL & "' "
END IF

sQuery = "ORDER BY A.OrderCode DESC, B.Idx "



SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Mobile_EShop_Order_Product_Select"

		.Parameters.Append .CreateParameter("@wQuery",		 adVarchar, adParaminput, 1000	, wQuery)
		.Parameters.Append .CreateParameter("@sQuery",		 adVarchar, adParaminput, 100	, sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing
%>
                                            <!-- 반품 신청 내역 -->
                                            <div id="processIssues_return" class="accord-sub-mypage">
                                                <div class="ly-title_sub">
                                                    <button type="button" class="btn-list clickEvt_sub" data-target="processIssues_return">반품 신청 (<%=FormatNumber(oRs.RecordCount, 0)%>)</button>
                                                </div>
                                                <div class="ly-content_sub">
<%
IF NOT oRs.EOF THEN	
%>
                                                    <ul class="informView">
<%
		Do Until oRs.EOF
				OrderStateNM	= GetOrderState(oRs("OrderState"), oRs("CancelState1"), oRs("CancelState2"))

				IF oRs("ProductImage") = "" THEN
						ProductImage	= "/Images/180_noimage.png"
				ELSE
						ProductImage	= oRs("ProductImage")
				END IF
%>
                                                        <li class="informItem">
                                                            <a href="/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>">
															<span class="cont">
																<span class="thumbNail">
																	<span class="img">
																		<img src="<%=ProductImage%>" alt="상품 이미지">
																	</span>
																	<span class="about">
																		<span class="process"><%=OrderStateNM%></span>
																	</span>
																</span>
							
																<span class="detail">
																	<span class="brand">
																		<span class="name"><%=oRs("BrandName")%> </span>
																		<%IF oRs("ProductType") = "O" THEN%><span class="oneplusone">[1+1]</span><%END IF%>
																	</span>
																	<span class="product-name"><em><%=oRs("ProductName")%></em></span>
																	
																	<span class="inform">
																		<span class="list">
																			<span class="tit">옵션</span>
																			<span class="opt"><%=oRs("SizeCD") %></span>
																		</span>
																		<span class="list">
																			<span class="tit">수량</span>
																			<span class="opt"><%=oRs("OrderCnt") %></span>
																		</span>
																		<span class="list">
																			<span class="tit">결제금액</span>
																			<span class="opt price"><em><%=FormatNumber(oRs("OrderPrice"),0)%></em>원</span>
																		</span>
																		<span class="list">
																			<span class="tit">배송/결제</span>
																			<span class="opt"><%IF oRs("DelvType") = "S" THEN%>매장픽업<%ELSE%>택배<%END IF%> / <%=GetPayType(oRs("PayType"))%></span>
																		</span>
																	</span>
																</span>
															</span>
															</a>

                                                            <div class="buttongroup">
                                                                <button type="button" onclick="getOrderDetail('<%=oRs("Idx")%>')" class="button-ty2 is-expand ty-bd-gray">신청내용 확인</button>
                                                            </div>
                                                        </li>
<%
				oRs.MoveNext
		Loop
%>
                                                    </ul>
<%
ELSE
%>
													<div class="no-history">
														<p>신청 내역이 없습니다.</p>
													</div>
<%
END IF
oRs.Close
%>
                                                </div>
                                            </div>
                                            <!-- // 반품 신청 내역 -->


<%
'# 교환 신청내역
wQuery = "WHERE B.IsShowFlag = 'Y' AND B.ProductType = 'P' AND B.CancelRequestFlag = 'X' "
IF SDate <> "" THEN
		wQuery = wQuery & "AND A.OrderDate >= '" & Replace(SDate, "-", "") & "' "
END IF
IF EDate <> "" THEN
		wQuery = wQuery & "AND A.OrderDate <= '" & Replace(EDate, "-", "") & "' "
END IF
IF U_NUM <> "" THEN
		wQuery = wQuery & "AND A.UserID = '" & U_NUM & "' "
ELSEIF N_NAME <> "" THEN
		wQuery = wQuery & "AND (A.UserID = '' OR A.UserID IS NULL) AND A.OrderName = '" & N_NAME & "' AND A.OrderHp = '" & N_HP & "' AND A.OrderEmail = '" & N_EMAIL & "' "
END IF

sQuery = "ORDER BY A.OrderCode DESC, B.Idx "



SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Mobile_EShop_Order_Product_Select"

		.Parameters.Append .CreateParameter("@wQuery",		 adVarchar, adParaminput, 1000	, wQuery)
		.Parameters.Append .CreateParameter("@sQuery",		 adVarchar, adParaminput, 100	, sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing
%>
                                            <!-- 교환 신청 내역 -->
                                            <div id="processIssues_exchange" class="accord-sub-mypage">
                                                <div class="ly-title_sub">
                                                    <button type="button" class="btn-list clickEvt_sub" data-target="processIssues_exchange">교환 신청 (<%=FormatNumber(oRs.RecordCount, 0)%>)</button>
                                                </div>
                                                <div class="ly-content_sub">
<%
IF NOT oRs.EOF THEN	
%>
                                                    <ul class="informView">
<%
		Do Until oRs.EOF
				OrderStateNM	= GetOrderState(oRs("OrderState"), oRs("CancelState1"), oRs("CancelState2"))

				IF oRs("ProductImage") = "" THEN
						ProductImage	= "/Images/180_noimage.png"
				ELSE
						ProductImage	= oRs("ProductImage")
				END IF
%>
                                                        <li class="informItem">
                                                            <a href="/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>">
															<span class="cont">
																<span class="thumbNail">
																	<span class="img">
																		<img src="<%=ProductImage%>" alt="상품 이미지">
																	</span>
																	<span class="about">
																		<span class="process"><%=OrderStateNM%></span>
																	</span>
																</span>
							
																<span class="detail">
																	<span class="brand">
																		<span class="name"><%=oRs("BrandName")%> </span>
																		<%IF oRs("ProductType") = "O" THEN%><span class="oneplusone">[1+1]</span><%END IF%>
																	</span>
																	<span class="product-name"><em><%=oRs("ProductName")%></em></span>
																	
																	<span class="inform">
																		<span class="list">
																			<span class="tit">옵션</span>
																			<span class="opt"><%=oRs("SizeCD") %></span>
																		</span>
																		<span class="list">
																			<span class="tit">수량</span>
																			<span class="opt"><%=oRs("OrderCnt") %></span>
																		</span>
																		<span class="list">
																			<span class="tit">결제금액</span>
																			<span class="opt price"><em><%=FormatNumber(oRs("OrderPrice"),0)%></em>원</span>
																		</span>
																		<span class="list">
																			<span class="tit">배송/결제</span>
																			<span class="opt"><%IF oRs("DelvType") = "S" THEN%>매장픽업<%ELSE%>택배<%END IF%> / <%=GetPayType(oRs("PayType"))%></span>
																		</span>
																	</span>
																</span>
															</span>
															</a>

                                                            <div class="buttongroup">
                                                                <button type="button" onclick="getOrderDetail('<%=oRs("Idx")%>')" class="button-ty2 is-expand ty-bd-gray">신청내용 확인</button>
                                                            </div>
                                                        </li>
<%
				oRs.MoveNext
		Loop
%>
                                                    </ul>
<%
ELSE
%>
													<div class="no-history">
														<p>신청 내역이 없습니다.</p>
													</div>
<%
END IF
oRs.Close
%>
                                                </div>
                                            </div>
                                            <!-- // 교환 신청 내역 -->

										</div>

										<script type="text/javascript">
											$(function () {
												// 주문취소/반품/교환 내 아코디언 안에 아코디언
												var orderCrxAccodion = function () {
													var selector,
														module;

													selector = {
														parent_sub: '.accord-sub-mypage',
														button_sub: '.clickEvt_sub',
														toggler_sub: '.ly-title_sub',
														panel_sub: '.ly-content_sub'
													};

													module = {
														init: function () {
															$(selector.button_sub).on('click', function () {
																module.accordion_sub(this);
															});
															$(window).trigger('scroll');
															$(selector.parent_sub).eq(0).find($(selector.panel_sub)).show();
															$(selector.parent_sub).eq(0).find($(selector.toggler_sub)).addClass('is-on');
														},
														accordion_sub: function (el) {
															var target = $(el).data('target');

															$(selector.panel_sub).slideUp(300);
															$(selector.toggler_sub).removeClass('is-on');

															if ($(selector.panel_sub, '#' + target).css('display') === 'none') {
																$(selector.panel_sub, '#' + target).slideDown(300);
																$(selector.toggler_sub, '#' + target).addClass('is-on');
															}
														}
													};
													module.init();
												}();
											})
										</script>
<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>