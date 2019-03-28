<%@ Language=VBScript codepage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'OrderList.asp - 주문 리스트
'Date		: 2018.12.31
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
DIM SOrderState

DIM OrderStateNM
DIM ProductImage
DIM DeliveryUrl
DIM	ReviewFlag


DIM OrderTotalCnt	: OrderTotalCnt	= 0		'# 주문총건수
DIM OrderState_1	: OrderState_1	= 0		'# 주문접수
DIM OrderState_3	: OrderState_3	= 0		'# 결제완료
DIM OrderState_4	: OrderState_4	= 0		'# 상품준비중
DIM OrderState_5	: OrderState_5	= 0		'# 배송중
DIM OrderState_6	: OrderState_6	= 0		'# 배송완료
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'
	
	
SDate			 = sqlFilter(Request("SDate"))
EDate			 = sqlFilter(Request("EDate"))
SOrderState		 = sqlFilter(Request("SOrderState"))


SET oConn				 = ConnectionOpen()							'# 커넥션 생성
SET oRs					 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
SET oRs1				 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성




Response.Write "OK|||||"


wQuery = "WHERE B.IsShowFlag = 'Y' AND B.ProductType = 'P' AND B.OrderState IN ('1', '3', '4', '5', '6', '7') "
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


SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Order_Product_Select_For_OrderState_Count"

		.Parameters.Append .CreateParameter("@wQuery",		 adVarchar, adParaminput, 1000	, wQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN	
		OrderState_1	= oRs("OrderState_1")
		OrderState_3	= oRs("OrderState_3")
		OrderState_4	= oRs("OrderState_4")
		OrderState_5	= oRs("OrderState_5")
		OrderState_6	= oRs("OrderState_6")
		OrderTotalCnt	= OrderState_1 + OrderState_3 + OrderState_4 + OrderState_5 + OrderState_6
END IF
oRs.Close



IF SOrderState = "5" THEN
		wQuery = wQuery & "AND B.OrderState = '5' AND B.CancelState2 = '0' "
ELSEIF SOrderState = "6" THEN
		wQuery = wQuery & "AND B.OrderState IN ('6', '7') "
ELSEIF SOrderState <> "" THEN
		wQuery = wQuery & "AND B.OrderState = '" & SOrderState & "' "
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
                                        <div class="h-line">
                                            <h2 class="h-level4">주문조회 결과</h2>
                                            <span class="h-date is-right"><%=SDate%> ~ <%=EDate%></span>
                                        </div>
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

				DeliveryUrl		= ""

				'# 배송조회 URL
				IF oRs("DelvType") = "P" AND IsNull(oRs("DelvComp")) = false AND oRs("DelvComp") <> "" AND IsNull(oRs("DelvNumber")) = false AND oRs("DelvNumber") <> "" THEN
						SET oCmd = Server.CreateObject("ADODB.Command")
						WITH oCmd
								.ActiveConnection	 = oConn
								.CommandType		 = adCmdStoredProc
								.CommandText		 = "USP_Admin_EShop_Delivery_Company_Select_By_DelvCompCode"

								.Parameters.Append .CreateParameter("@DelvCompCode",	 adChar, adParaminput,	8,	oRs("DelvComp"))
						END WITH
						oRs1.CursorLocation = adUseClient
						oRs1.Open oCmd, , adOpenStatic, adLockReadOnly
						SET oCmd = Nothing

						IF NOT oRs1.EOF THEN
								IF IsNull(oRs1("DelvTraceUrl")) = false AND oRs1("DelvTraceUrl") <> "" THEN
										DeliveryUrl		= oRs1("DelvTraceUrl") & Replace(oRs("DelvNumber"), "-", "")
								END IF
						END IF
						oRs1.Close
				END IF

				ReviewFlag = "N"
				IF InStr(OrderStateNM, "구매확정") > 0 THEN
						'-----------------------------------------------------------------------------------------------------------'
						'# 상품후기 등록여부 체크 Start
						'-----------------------------------------------------------------------------------------------------------'
						SET oCmd = Server.CreateObject("ADODB.Command")
						WITH oCmd
								.ActiveConnection	 = oConn
								.CommandType		 = adCmdStoredProc
								.CommandText		 = "USP_Front_EShop_Product_Review_Select_By_Order_Product_Idx"

								.Parameters.Append .CreateParameter("@Order_Product_Idx",	 adInteger, adParaminput, 	, oRs("Idx"))
						END WITH
						oRs1.CursorLocation = adUseClient
						oRs1.Open oCmd, , adOpenStatic, adLockReadOnly
						SET oCmd = Nothing

						IF NOT oRs1.EOF THEN
								ReviewFlag = "Y"
						END IF
						oRs1.Close
						'-----------------------------------------------------------------------------------------------------------'
						'# 상품후기 등록여부 체크 End
						'-----------------------------------------------------------------------------------------------------------'
				END IF
%>
                                            <li class="informItem">
                                                <a href="javascript:getOrderDetail(<%=oRs("Idx")%>);">
													<span class="head-tit">
														<span class="tit">주문번호 : <%=oRs("OrderCode")%></span>
														<span class="date"><%=GetDateYMD(oRs("OrderDate"))%></span>
													</span>
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
															</span>
															<span class="product-name"><em><%=oRs("ProductName")%></em></span>
															
															<span class="inform">
																<span class="list">
																	<span class="tit">옵션</span>
																	<span class="opt"><%=oRs("SizeCD")%></span>
																</span>
																<span class="list">
																	<span class="tit">수량</span>
																	<span class="opt"><%=oRs("OrderCnt")%></span>
																</span>
																<span class="list">
																	<span class="tit">결제금액</span>
																	<span class="opt price"><em><%=FormatNumber(oRs("OrderPrice"),0)%></em>원</span>
																</span>
																<span class="list">
																	<span class="tit">구분</span>
																	<span class="opt"><%=oRs("VendorNM")%>배송</span>
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
													<%IF InStr(OrderStateNM, "입금대기") > 0 THEN%>
                                                    <button type="button" onclick="openNonDepositOrderCancel('<%=oRs("OrderCode")%>')" class="button-ty2 is-expand ty-bd-gray">주문 취소하기</button>
													<%ELSEIF InStr(OrderStateNM, "결제완료") > 0 THEN%>
                                                    <button type="button" onclick="openOrderCancel('C','<%=oRs("OrderCode")%>','<%=oRs("Idx")%>')" class="button-ty2 is-expand ty-bd-gray">주문 취소하기</button>
													<%ELSEIF InStr(OrderStateNM, "상품준비중") > 0 THEN%>
                                                    <button type="button" onclick="openOrderCancel('R','<%=oRs("OrderCode")%>','<%=oRs("Idx")%>')" class="button-ty2 is-expand ty-bd-gray">주문 취소신청하기</button>
													<%ELSEIF InStr(OrderStateNM, "배송중") > 0 OR InStr(OrderStateNM, "배송완료") > 0 THEN%>
													<div class="bd">
														<button type="button" onclick="window.open('<%=DeliveryUrl%>','delv_pop','resizable=yes, fullscreen=no, menubar=no, status=no, toolbar=no, scrollbars=yes')" class="merger part-4">배송조회</button>
														<button type="button" onclick="openOrderConfirm('<%=oRs("OrderCode")%>','<%=oRs("Idx")%>')" class="merger part-4">구매확정</button>
														<button type="button" onclick="openOrderChangeReturn('R','<%=oRs("OrderCode")%>','<%=oRs("Idx")%>')" class="merger part-4">반품신청</button>
														<button type="button" onclick="openOrderChangeReturn('X','<%=oRs("OrderCode")%>','<%=oRs("Idx")%>')" class="merger part-4">교환신청</button>
													</div>
													<%ELSEIF InStr(OrderStateNM, "구매확정") > 0 AND ReviewFlag = "N" THEN%>
                                                    <button type="button" onclick="openReviewWrite('<%=oRs("OrderCode")%>','<%=oRs("Idx")%>')" class="button-ty2 ty-bd-gray part-2">후기 작성하기</button>
                                                    <button type="button" onclick="openAfterService('<%=oRs("OrderCode")%>','<%=oRs("Idx")%>')" class="button-ty2 ty-bd-gray part-2">A/S 신청</button>
													<%ELSEIF InStr(OrderStateNM, "구매확정") > 0 AND ReviewFlag = "Y" THEN%>
                                                    <button type="button" onclick="openAfterService('<%=oRs("OrderCode")%>','<%=oRs("Idx")%>')" class="button-ty2 is-expand ty-bd-gray">A/S 신청</button>
													<%END IF%>
                                                </div>
                                            </li>
<%
				oRs.MoveNext
		Loop
%>
                                        </ul>

                                        <div class="inf-type1">
                                            <p class="tit">알려드립니다.</p>
                                            <ul>
                                                <li class="bullet-ty1">상품정보를 클릭하면 주문상세내역을 확인하실 수 있습니다.</li>
                                                <li class="bullet-ty1">배송상태가 ‘배송중’인 경우 ‘배송추적’을 클릭하시면 배송상황을 자세히 확인하실 수 있습니다.</li>
                                                <li class="bullet-ty1">주문하신 상품의 배송업체가 다를 경우 상품별로 따로 배송될 수도 있습니다.</li>
                                                <li class="bullet-ty1">‘구매확정’ 후 ‘상품후기 쓰기’가 가능하며 참여 시 포인트가 지급됩니다.</li>
                                                <li class="bullet-ty1">배송후 10일 이후에 회원등급에 따른 구매포인트가 자동으로 적립됩니다. (즉시사용가능)</li>
                                            </ul>
                                        </div>
<%
ELSE
%>
                                        <div class="no-history">
                                            <p>주문 내역이 없습니다.</p>
                                        </div>
<%
END IF
oRs.Close
%>
<%
SET oRs1 = Nothing
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>