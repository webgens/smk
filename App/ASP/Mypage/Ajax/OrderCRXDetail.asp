<%@ Language=VBScript codepage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'OrderCRXDetail.asp - 주문취소/반품/교환 상세내역
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

DIM Page
DIM SDate
DIM EDate
DIM SCancelType

DIM SOPIdx

DIM Num				: Num = 0
DIM OrderStateNM
DIM ProductImage
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'
	
	
Page			 = sqlFilter(Request("Page"))
SDate			 = sqlFilter(Request("SDate"))
EDate			 = sqlFilter(Request("EDate"))
SCancelType		 = sqlFilter(Request("SCancelType"))
SOPIdx			 = sqlFilter(Request("SOPIdx"))

IF Page			 = "" THEN Page	 = 1


SET oConn				 = ConnectionOpen()							'# 커넥션 생성
SET oRs					 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
SET oRs1				 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성




Response.Write "OK|||||"


wQuery = ""
wQuery = wQuery & "WHERE A.IsShowFlag = 'Y' AND A.ProductType IN ('P','O') AND (A.OrderState = 'C' OR A.CancelRequestFlag IN ('C', 'R', 'X')) "
wQuery = wQuery & "AND A.Idx = " & SOPIdx & " "

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
		OrderStateNM	= GetOrderState(oRs("OrderState"), oRs("CancelState1"), oRs("CancelState2"))
		IF oRs("ProductImage_180") = "" THEN
				ProductImage	= "/Images/180_noimage.png"
		ELSE
				ProductImage	= oRs("ProductImage_180")
		END IF

		DIM CancelTypeNM
		SELECT CASE oRs("CancelRequestFlag")
				CASE "C" : CancelTypeNM = "취소"
				CASE "R" : CancelTypeNM = "반품"
				CASE "X" : CancelTypeNM = "교환"
				CASE ELSE IF oRs("OrderState") = "C" THEN CancelTypeNM = "취소"
		END SELECT
%>
        <div class="area-dim"></div>

        <div class="area-pop">
            <div class="full">
                <div class="tit-pop">
                    <div class="tit">주문 <%=CancelTypeNM%><%IF CStr(oRs("OPCIdx")) <> "0" THEN%>신청<%END IF%> 내역</div>
                    <button type="button" onclick="closePop('DimDepth1')" class="btn-hide-pop">닫기</button>
                </div>

                <div class="container-pop mypage-ty2">
                    <div class="contents">
                        <div class="wrap-order">

					<%IF CStr(oRs("OPCIdx")) = "0" THEN%>
                            <div class="h-line">
                                <h3 class="h-level4"><%=CancelTypeNM%> 상품</h3>
                            </div>
					<%ELSE%>
                            <div class="h-line">
                                <h3 class="h-level4">신청한 주문</h3>
                            </div>
					<%END IF%>


                            <ul class="informView">
                                <li class="informItem" style="margin-bottom:0;">
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
											<a href="#">
												<span class="brand">
													<span class="name"><%=oRs("BrandName")%></span>
												</span>
												<span class="product-name"><em><%=oRs("ProductName")%></em></span>
											</a>

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
													<span class="tit">배송방법</span>
													<span class="opt"><%IF oRs("DelvType") = "S" THEN%>매장픽업<%ELSE%>택배<%END IF%></span>
												</span>
											</span>
										</span>
                                    </span>
                                </li>
					<%
					'# 1+1상품
					IF (oRs("OrderState") = "C" OR oRs("CancelRequestFlag") = "C" OR oRs("CancelRequestFlag") = "R") AND oRs("GroupCnt") > 1 THEN
							wQuery = ""
							wQuery = wQuery & "WHERE A.IsShowFlag = 'Y' AND A.ProductType = 'O' "
							wQuery = wQuery & "AND A.OPIdx_Group = " & oRs("OPIdx_Group") & " "

							sQuery = "ORDER BY A.Idx DESC"


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
									IF oRs1("ProductImage_180") = "" THEN
											ProductImage	= "/Images/180_noimage.png"
									ELSE
											ProductImage	= oRs1("ProductImage_180")
									END IF
					%>
                                <li class="informItem" style="border-top:none;">
                                    <span class="cont">
										<span class="thumbNail">
											<span class="img">
												<img src="<%=ProductImage%>" alt="상품 이미지">
											</span>
											<span class="about">
												<span class="process"><%=GetOrderState(oRs1("OrderState"), oRs1("CancelState1"), oRs1("CancelState2"))%></span>
											</span>
				                        </span>

										<span class="detail">
											<a href="#">
												<span class="brand">
													<span class="name"><%=oRs1("BrandName")%></span>
													<span class="oneplusone"><strong>[1+1]</strong></span>
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
													<span class="tit">배송</span>
													<span class="opt"><%IF oRs("DelvType") = "S" THEN%>매장픽업<%ELSE%>택배<%END IF%></span>
												</span>
											</span>
										</span>
                                    </span>
                                </li>
					<%
							END IF
							oRs1.Close
					END IF
					%>
                            </ul>

<%
		'-----------------------------------------------------------------------------------------------------------'
		'# 취소/교환/반품 신청이 없을 경우 Start
		'-----------------------------------------------------------------------------------------------------------'
		IF CStr(oRs("OPCIdx")) = "0" THEN
%>
                            <div class="h-line">
                                <h3 class="h-level4">주문 상세</h3>
                            </div>

                            <ul class="detailView">
                                <li class="detailList">
                                    <div class="tit">받는분</div>
                                    <div class="cont"><span class="general"><%=oRs("ReceiveName")%></span></div>
                                </li>
                                <li class="detailList">
                                    <div class="tit">전화번호</div>
                                    <div class="cont"><span class="general"><%=oRs("ReceiveTel")%></span></div>
                                </li>
                                <li class="detailList">
                                    <div class="tit">휴대전화</div>
                                    <div class="cont"><span class="general"><%=oRs("ReceiveHp")%></span></div>
                                </li>
                                <li class="detailList">
                                    <div class="tit">배송지</div>
                                    <div class="cont">
										<span class="general">
										<%IF oRs("DelvType") = "S" THEN%><span class="general"><%=oRs("ShopNM")%></span><%END IF %>
										<span class="general">[<%=oRs("ReceiveZipCode")%>] <%=oRs("ReceiveAddr1")%> <%=oRs("ReceiveAddr2")%></span>
                                    </div>
                                </li>
                                <li class="detailList">
                                    <div class="tit">배송메모</div>
                                    <div class="cont"><span class="general"><%=oRs("Memo")%></span></div>
                                </li>
                            </ul>

                            <div class="h-line">
                                <h3 class="h-level4">결제정보</h3>
                            </div>

                            <ul class="detailView on-right">
                                <li class="detailList">
                                    <div class="tit">결제수단</div>
                                    <div class="cont">
                                        <span class="general"><em class="strong"><%=GetPayType(oRs("PayType"))%></em></span>
										<span class="general">
											<%IF oRs("PayType") = "C" THEN%>
												<%=oRs("LGD_FINANCENAME")%>카드 | 
												<%IF oRs("LGD_CARDINSTALLMONTH") = "00" THEN%>
													일시불
												<%ELSE%>
													<%=FormatNumber(oRs("LGD_CARDINSTALLMONTH"),0)%>개월 할부
												<%END IF%>
											<%ELSEIF oRs("PayType") = "B" THEN%>
												<%=oRs("LGD_FINANCENAME")%>은행
											<%ELSEIF oRs("PayType") = "V" THEN%>
												<%=oRs("LGD_FINANCENAME")%>은행 | <%=oRs("LGD_ACCOUNTNUM")%> | <%=MALL_LGD_ACCOUNTOWNER%>
											<%ELSEIF oRs("PayType") = "M" THEN%>
												<%=oRs("LGD_FINANCENAME")%> | <%=oRs("LGD_TELNO")%>
											<%ELSEIF oRs("PayType") = "N" THEN%>
											<%END IF%>
                                        </span>
										<span class="general"><%=GetDateYMD(oRs("OrderDate")) & " " & GetTimeHMS(oRs("OrderTime"))%></span>
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
                                    <div class="tit"><em class="strong">총 결제 금액</em></div>
                                    <div class="cont">
                                        <span class="general ty-red"><em class="strong"><%=FormatNumber(oRs("OrderPrice"),0)%></em>원</span>
                                    </div>
                                </li>
                            </ul>

                            <div class="h-line">
                                <h3 class="h-level4">적립내역</h3>
                            </div>

                            <ul class="detailView">
                                <li class="detailList">
                                    <div class="tit"><%IF oRs("OrderState") = "7" THEN%>적립<%ELSE%>적립 예정<%END IF%> 포인트</div>
                                    <div class="cont on-right"><em><%=FormatNumber(oRs("ProductPoint"),0)%></em>P</div>
                                </li>
                            </ul>

<%
		'-----------------------------------------------------------------------------------------------------------'
		'# 취소/교환/반품 신청이 없을 경우 End
		'-----------------------------------------------------------------------------------------------------------'


		'-----------------------------------------------------------------------------------------------------------'
		'# 취소/교환/반품 신청이 있을 경우 Start
		'-----------------------------------------------------------------------------------------------------------'
		ELSE
				DIM SizeChangeFlag
				DIM ChgSizeCD
				DIM ReasonName
				DIM CancelMemo
				DIM DelvFee
				DIM DelvFeeType
				DIM DepositBankName
				DIM DepositNumber
				DIM DepositName
				DIM ReturnName
				DIM ReturnHp
				DIM ReturnZipCode
				DIM ReturnAddr1
				DIM ReturnAddr2
				DIM ReceiveName
				DIM ReceiveHp
				DIM ReceiveZipCode
				DIM ReceiveAddr1
				DIM ReceiveAddr2

				SET oCmd = Server.CreateObject("ADODB.Command")
				WITH oCmd
						.ActiveConnection	 = oConn
						.CommandType		 = adCmdStoredProc
						.CommandText		 = "USP_Front_EShop_Order_Product_Cancel_Select_By_Key"

						.Parameters.Append .CreateParameter("@OPIdx",		 adInteger,	adParaminput, 	 , oRs("Idx"))
						.Parameters.Append .CreateParameter("@CancelType",	 adChar,	adParaminput, 	1, oRs("CancelRequestFlag"))
				END WITH
				oRs1.CursorLocation = adUseClient
				oRs1.Open oCmd, , adOpenStatic, adLockReadOnly
				SET oCmd = Nothing


				IF NOT oRs1.EOF THEN
						SizeChangeFlag		= oRs1("SizeChangeFlag")
						ChgSizeCD			= oRs1("SizeCD")
						ReasonName			= oRs1("ReasonName")
						CancelMemo			= oRs1("Memo")

						DelvFee				= oRs1("DelvFee")
						DelvFeeType			= oRs1("DelvFeeType")

						DepositBankName		= oRs1("DepositBankName")
						DepositNumber		= oRs1("DepositNumber")
						DepositName			= oRs1("DepositName")

						ReturnName			= oRs1("ReturnName")
						ReturnHp			= oRs1("ReturnHp")
						ReturnZipCode		= oRs1("ReturnZipCode")
						ReturnAddr1			= oRs1("ReturnAddr1")
						ReturnAddr2			= oRs1("ReturnAddr2")

						ReceiveName			= oRs1("ReceiveName")
						ReceiveHp			= oRs1("ReceiveHp")
						ReceiveZipCode		= oRs1("ReceiveZipCode")
						ReceiveAddr1		= oRs1("ReceiveAddr1")
						ReceiveAddr2		= oRs1("ReceiveAddr2")
				END IF
				oRs1.Close
%>
				<%IF oRs("CancelRequestFlag") = "X" THEN%>
                            <div class="h-line">
                                <h3 class="h-level4">교환 사이즈</h3>
                            </div>

                            <div class="reason">
                                <span class="input">
									<input type="text" placeholder="[<%=ChgSizeCD%>] <%IF SizeChangeFlag = "Y" THEN Response.Write "사이즈교환" ELSE Response.Write "동일사이즈" END IF%>" disabled />
								</span>
                            </div>
				<%END IF%>

                            <div class="h-line">
                                <h3 class="h-level4"><%=CancelTypeNM%> 사유</h3>
                            </div>

                            <div class="reason">
                                <span class="input">
									<input type="text" placeholder="[<%=ReasonName%>] <%=CancelMemo%>" disabled>
								</span>
                            </div>

				<%IF oRs("CancelRequestFlag") = "R" OR oRs("CancelRequestFlag") = "X" THEN%>
                            <div class="h-line">
                                <h3 class="h-level4"><%=CancelTypeNM%> 배송비 결제방법</h3>
                            </div>

                            <div class="delivery-price">
                                <div class="list">
                                    <div class="tit">
                                        <span><%=CancelTypeNM%> 택배 배송비 <em class="ty-red"><%=FormatNumber(DelvFee,0)%>원</em>을 <em class="ty-red"><%=GetDelvFeeType(DelvFeeType)%></em></span>
                                    </div>
                                </div>
                            </div>
				<%END IF%>

				<%IF oRs("CancelRequestFlag") = "R" OR oRs("CancelRequestFlag") = "X" THEN%>
                            <div class="h-line">
                                <h3 class="h-level4">상품 수거지</h3>
                            </div>

                            <div class="addr-list">
                                <div class="list">
                                    <div class="tit">
                                        <span><%=ReturnName%></span><span><%=ReturnHp%></span>
                                    </div>
                                    <div class="address">[<%=ReturnZipCode%>] <%=ReturnAddr1%> <%=ReturnAddr2%></div>
                                </div>
                            </div>
				<%END IF%>

				<%IF oRs("CancelRequestFlag") = "X" THEN%>
                            <div class="h-line">
                                <h3 class="h-level4">상품 수령지</h3>
                            </div>

                            <div class="addr-list">
                                <div class="list">
                                    <div class="tit">
                                        <span><%=ReceiveName%></span><span><%=ReceiveHp%></span>
                                    </div>
                                    <div class="address">[<%=ReceiveZipCode%>] <%=ReceiveAddr1%> <%=ReceiveAddr2%></div>
                                </div>
                            </div>
				<%END IF%>

				<%IF IsNull(DepositNumber) = false AND DepositNumber <> "" THEN%>
                            <div class="h-line">
                                <h3 class="h-level4">환불 계좌 정보</h3>
                            </div>

                            <div class="addr-list">
                                <div class="list">
                                    <div class="tit">
                                        <span><%=DepositBankName%></span><span><%=DepositNumber%></span><span><%=DepositName%></span>
                                    </div>
                                </div>
                            </div>
				<%END IF%>


				<%IF oRs("CancelRequestFlag") = "X" THEN%>
                            <div class="h-line">
                                <h3 class="h-level4">교환 절차 안내</h3>
                            </div>

                            <div class="process-info">
                                <ol class="exchange">
                                    <li<%IF oRs("CancelState2") = "2" THEN%> class="current"<%END IF%>>반품<br>접수</li>
                                    <li<%IF oRs("CancelState2") = "3" THEN%> class="current"<%END IF%>>제품<br>회수</li>
                                    <li<%IF oRs("CancelState2") = "4" THEN%> class="current"<%END IF%>>검수<br>승인</li>
                                    <li<%IF oRs("CancelState2") = "7" THEN%> class="current"<%END IF%>>교환<br>발송</li>
                                    <li<%IF oRs("OrderState") = "7" THEN%> class="current"<%END IF%>>교환<br>완료</li>
                                </ol>
                            </div>
				<%ELSEIF oRs("CancelRequestFlag") = "R" THEN%>
                            <div class="h-line">
                                <h3 class="h-level4">반품 절차 안내</h3>
                            </div>

                            <div class="process-info">
                                <ol class="return">
                                    <li<%IF oRs("CancelState2") = "2" THEN%> class="current"<%END IF%>>반품<br>접수</li>
                                    <li<%IF oRs("CancelState2") = "3" THEN%> class="current"<%END IF%>>제품<br>회수</li>
                                    <li<%IF oRs("CancelState2") = "4" THEN%> class="current"<%END IF%>>검수<br>승인</li>
                                    <li<%IF oRs("CancelState2") = "7" THEN%> class="current"<%END IF%>>환불<br>완료</li>
                                </ol>
                            </div>
				<%END IF%>
<%
		END IF
%>
<%
ELSE
%>
							<div class="area-empty">
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