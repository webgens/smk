<%@ Language=VBScript codepage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'OrderDetailList.asp - 주문 상세 리스트
'Date		: 2019.01.02
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

DIM OrderCode

DIM OrderStateNM
DIM ProductImage
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'
	
	
OrderCode		 = sqlFilter(Request("OrderCode"))


SET oConn				 = ConnectionOpen()							'# 커넥션 생성
SET oRs					 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
SET oRs1				 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성




Response.Write "OK|||||"
%>
        <div class="area-dim"></div>

        <div class="area-pop">
            <div class="full">
                <div class="tit-pop">
                    <div class="tit">주문조회 상세</div>
                    <button type="button" onclick="closePop('DimDepth1')" class="btn-hide-pop">닫기</button>
                </div>

                <div class="container-pop mypage-ty1">
                    <!-- 팝업 스타일 변경으로 'mypage-ty1'클래스 명 추가 -->
                    <div class="contents">
                        <div class="wrap-order">
<%
wQuery = "WHERE A.IsShowFlag = 'Y' "
wQuery = wQuery & "AND A.OrderCode = '" & OrderCode & "' "
IF U_NUM <> "" THEN
		wQuery = wQuery & "AND A.UserID = '" & U_NUM & "' "
ELSEIF N_NAME <> "" THEN
		wQuery = wQuery & "AND (A.UserID = '' OR A.UserID IS NULL) AND A.OrderName = '" & N_NAME & "' AND A.OrderHp = '" & N_HP & "' AND A.OrderEmail = '" & N_EMAIL & "' "
END IF

sQuery = ""




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

IF NOT oRs.EOF THEN	
%>
                            <div class="order-number">
                                <p class="number">주문번호 : <%=oRs("OrderCode")%></p>
                                <div class="date">(<%=GetDateYMD(oRs("OrderDate"))%>)</div>
                            </div>

                            <div class="h-line">
                                <h3 class="h-level4">주문 상품</h3>
                            </div>

                            <ul class="informView">
<%



		wQuery = ""
		wQuery = wQuery & "WHERE A.IsShowFlag = 'Y' AND A.ProductType = 'P' AND A.OrderState IN ('1', '3', '4', '5', '6', '7') "
		wQuery = wQuery & "AND A.OrderCode = '" & oRs("OrderCode") & "' "

		sQuery = "ORDER BY A.OPIdx_Group, A.OPIdx_Org, A.Idx"




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


		IF NOT oRs1.EOF THEN	
				Do Until oRs1.EOF
						OrderStateNM	= GetOrderState(oRs1("OrderState"), oRs1("CancelState1"), oRs1("CancelState2"))
						IF oRs1("ProductImage_180") = "" THEN
								ProductImage	= "/Images/180_noimage.png"
						ELSE
								ProductImage	= oRs1("ProductImage_180")
						END IF
%>
                                <li class="informItem">
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
											<a href="#">
												<span class="brand">
													<span class="name"><%=oRs1("BrandName")%></span>
												</span>
												<span class="product-name"><em><%=oRs1("ProductName")%></em></span>
											</a>

											<span class="inform">
												<span class="list">
													<span class="tit">옵션</span>
													<span class="opt"><%=oRs1("SizeCD")%></span>
												</span>
												<span class="list">
													<span class="tit">수량</span>
													<span class="opt"><%=oRs1("OrderCnt")%></span>
												</span>
												<span class="list">
													<span class="tit">결제금액</span>
													<span class="opt price"><em><%=FormatNumber(oRs1("OrderPrice"),0)%></em>원</span>
												</span>
												<span class="list">
													<span class="tit">구분</span>
													<span class="opt"><%=oRs1("VendorName")%> 배송</span>
												</span>
												<span class="list">
													<span class="tit">배송</span>
													<span class="opt"><%IF oRs1("DelvType") = "S" THEN%>매장픽업<%ELSE%>택배<%END IF%></span>
												</span>
											</span>
										</span>
                                    </span>
                                </li>
<%
						oRs1.MoveNext
				Loop
		END IF
		oRs1.Close
%>
                            </ul>

                            <div class="h-line">
                                <h3 class="h-level4">결제정보</h3>
                            </div>

                            <ul class="detailView on-right">
                                <li class="detailList">
                                    <div class="tit">결제수단</div>
                                    <div class="cont">
                                        <span class="general"><em class="strong"><%=GetPayType(oRs("PayType"))%></em></span>
                                    </div>
                                </li>
                                <li class="detailList">
                                    <div class="tit">주문상품금액</div>
                                    <div class="cont">
                                        <span class="general"><em class="strong"><%=FormatNumber(oRs("SalePrice"),0)%></em>원</span>
                                    </div>
                                </li>
                                <li class="detailList">
                                    <div class="tit">쿠폰할인</div>
                                    <div class="cont">
                                        <span class="general"><%=FormatNumber(oRs("UseCouponPrice"),0)%>원</span>
                                    </div>
                                </li>
                                <li class="detailList">
                                    <div class="tit">포인트차감</div>
                                    <div class="cont">
                                        <span class="general"><%=FormatNumber(oRs("UsePointPrice"),0)%>원</span>
                                    </div>
                                </li>
                                <li class="detailList">
                                    <div class="tit">슈즈상품권차감</div>
                                    <div class="cont">
                                        <span class="general"><%=FormatNumber(oRs("UseScashPrice"),0)%>원</span>
                                    </div>
                                </li>
                                <li class="detailList">
                                    <div class="tit">배송비</div>
                                    <div class="cont">
                                        <span class="general"><em class="strong"><%=FormatNumber(oRs("DeliveryPrice"),0)%></em>원</span>
                                    </div>
                                </li>
                                <li class="detailList">
                                    <div class="tit"><em class="strong">총 결제 금액</em></div>
                                    <div class="cont">
                                        <span class="general ty-red"><em class="strong"><%=FormatNumber(oRs("OrderPrice") + oRs("DeliveryPrice"),0)%></em>원</span>
                                    </div>
                                </li>
                            </ul>
<%
ELSE
%>
							<div class="area-empty" style="margin-bottom: 30px">
								<span class="icon-empty"></span>
								<p class="tit-empty">주문 내역이 존재하지 않습니다.</p>
							</div>
<%
END IF
oRs.Close
%>
                        </div>
                    </div>

                    <div class="btns">
                        <button type="button" onclick="closePop('DimDepth1')" class="button ty-red">목록으로</button>
                    </div>
                </div>
            </div>
        </div>
<%
SET oRs1 = Nothing
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>