<%@ Language=VBScript codepage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'ReceiptList.asp - 영수증 발급 리스트
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
<!-- #include virtual = "/Common/OpenXpay/lgdacom/md5.asp" -->

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

DIM OrderCnt
DIM ProductImage
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'
	
	
SDate			 = sqlFilter(Request("SDate"))
EDate			 = sqlFilter(Request("EDate"))


SET oConn				 = ConnectionOpen()							'# 커넥션 생성
SET oRs					 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
SET oRs1				 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성




Response.Write "OK|||||"


wQuery = "WHERE A.IsShowFlag = 'Y' AND A.SettleFlag = 'Y' "
wQuery = wQuery & "AND (A.PayType = 'C' OR A.ReceiptFlag = 'Y') "
wQuery = wQuery & "AND A.OrderCode IN (SELECT DISTINCT OrderCode FROM EShop_Order_Product WITH (NOLOCK) WHERE IsShowFlag = 'Y' AND ProductType = 'P' AND OrderState IN ('3', '4', '5', '6', '7')) "
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

sQuery = "ORDER BY A.OrderCode DESC "




SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Mobile_EShop_Order_Select"

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
				'# 주문상품 리스트
				wQuery = ""
				wQuery = wQuery & "WHERE A.IsShowFlag = 'Y' AND A.ProductType = 'P' AND A.OPIdx_Prev = 0 AND A.OrderState IN ('3', '4', '5', '6', '7') "
				wQuery = wQuery & "AND A.OrderCode = '" & oRs("OrderCode") & "' "

				sQuery = "ORDER BY A.OPIdx_Org"


				SET oCmd = Server.CreateObject("ADODB.Command")
				WITH oCmd
						.ActiveConnection	 = oConn
						.CommandType		 = adCmdStoredProc
						.CommandText		 = "USP_Front_EShop_Order_Product_Select_For_Order_Detail"

						.Parameters.Append .CreateParameter("@WQuery",		 adVarchar, adParaminput, 1000	, wQuery)
						.Parameters.Append .CreateParameter("@SQuery",		 adVarchar, adParaminput, 100	, sQuery)
				END WITH
				oRs1.CursorLocation = adUseClient
				oRs1.Open oCmd, , adOpenStatic, adLockReadOnly
				SET oCmd = Nothing

				OrderCnt	= oRs1.RecordCount

				IF NOT oRs1.EOF THEN	
						IF oRs1("ProductImage") = "" THEN
								ProductImage	= "/Images/180_noimage.png"
						ELSE
								ProductImage	= oRs1("ProductImage")
						END IF
%>
                                            <li class="informItem">
                                                <a href="javascript:getOrderDetailList('<%=oRs("OrderCode")%>');">
													<span class="head-tit">
														<span class="tit">주문번호 : <%=oRs("OrderCode")%></span>
														<span class="date"><%=GetDateYMD(oRs("OrderDate"))%></span>
													</span>
													<span class="cont">
														<span class="thumbNail">
															<span class="img">
																<img src="<%=ProductImage%>" alt="상품 이미지">
															</span>
														</span>
							
														<span class="detail">
															<span class="brand">
																<span class="name"><%=oRs1("BrandName")%> </span>
															</span>
															<span class="product-name"><em><%=oRs1("ProductName")%> </em><%IF CInt(OrderCnt) > 1 THEN%><span class="mum-all">외 <%=FormatNumber(OrderCnt - 1,0)%>건</span><%END IF%></span>
															
															<span class="inform">
																<span class="list">
																	<span class="tit">결제금액</span>
																	<span class="opt price"><em><%=FormatNumber(oRs("OrderPrice") + oRs("DeliveryPrice"), 0)%></em>원<%IF CDbl(oRs("DeliveryPrice")) > 0 THEN%> <span class="deliverFee">(배송비 <%=FormatNumber(oRs("DeliveryPrice"),0)%>원 포함)</span><%END IF%></span>
																</span>
																<span class="list">
																	<span class="tit">배송비</span>
																	<span class="opt"><%IF CDbl(oRs("DeliveryPrice")) = 0 THEN%>무료배송<%ELSE%><%=FormatNumber(oRs("DeliveryPrice"),0)%>원<%END IF%></span>
																</span>
																<span class="list">
																	<span class="tit">결제수단</span>
																	<span class="opt"><%=GetPayType(oRs("PayType"))%></span>
																</span>
															</span>
														</span>
													</span>
												</a>

                                                <div class="buttongroup">
													<%
													IF oRs("PayType") = "C" THEN	'# 카드 영수증
														DIM AuthData
														AuthData = MD5(oRs("LGD_MID") & oRs("LGD_TID") & LGD_MERTKEY)
													%>
                                                    <button type="button" onclick="APP_PopupGoUrl('/ASP/MyPage/ajax/ReceiptView.asp?rType=C&LGD_MID=<%=oRs("LGD_MID")%>&LGD_TID=<%=oRs("LGD_TID")%>&AuthData=<%=AuthData%>', '0', '')" class="button-ty2 is-expand ty-bd-gray"><span class="icon ico-receipt">영수증 확인</span></button>
													<%ELSEIF oRs("PayType") = "B" And oRs("ReceiptFlag") = "Y" THEN	'# 계좌이체 현금영수증%>
                                                    <button type="button" onclick="APP_PopupGoUrl('/ASP/MyPage/ajax/ReceiptView.asp?rType=B&LGD_MID=<%=oRs("LGD_MID")%>&OrderCode=<%=oRs("OrderCode")%>', '0', '')" class="button-ty2 is-expand ty-bd-gray"><span class="icon ico-receipt">현금영수증조회</span></button>
													<%ELSEIF oRs("PayType") = "V" And oRs("ReceiptFlag") = "Y" THEN	'# 가상계좌 현금영수증%>
                                                    <button type="button" onclick="APP_PopupGoUrl('/ASP/MyPage/ajax/ReceiptView.asp?rType=V&LGD_MID=<%=oRs("LGD_MID")%>&OrderCode=<%=oRs("OrderCode")%>&LGD_CASSEQNO=<%=oRs("LGD_CASSEQNO")%>', '0', '')" class="button-ty2 is-expand ty-bd-gray"><span class="icon ico-receipt">현금영수증조회</span></button>
													<%END IF %>	
                                                </div>
                                            </li>
<%
				END IF
				oRs1.Close

				oRs.MoveNext
		Loop
%>
                                        </ul>
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