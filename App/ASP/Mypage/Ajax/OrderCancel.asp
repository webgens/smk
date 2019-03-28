<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'OrderCancel.asp - 주문취소 폼 페이지
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

DIM PayType
DIM EscrowFlag
DIM ProductImage
DIM TotalOrderCnt
DIM CancelableCount		: CancelableCount	= 0
DIM CancelableFlag

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




IF OrderCode = "" THEN
		Response.Write "FAIL|||||선택한 주문번호가 없습니다."
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
		'# 결제완료 상태가 아니면 취소불가
		IF oRs("SettleFlag") <> "Y" THEN
				oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Response.Write "FAIL|||||결제완료 되지않은 주문은 주문취소할 수 없습니다."
				Response.End
		END IF

		PayType			= oRs("PayType")
		EscrowFlag		= oRs("EscrowFlag")
		TotalOrderCnt	= oRs("OrderCnt")

ELSE
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
		Response.Write "FAIL|||||취소할 주문내역이 없습니다."
		Response.End
END IF
oRs.Close


Response.Write "OK|||||"
%>					
        <div class="area-dim"></div>

        <div class="area-pop" id="OrderCancel">
            <div class="full">
                <div class="tit-pop">
                    <div class="tit">주문 취소하기</div>
                    <button type="button" onclick="closePop('DimDepth1')" class="btn-hide-pop">닫기</button>
                </div>

                <div class="container-pop mypage-ty1">
                    <div class="contents">
                        <div class="wrap-order">
                            <div class="order-number">
                                <p class="number">주문번호 : <%=OrderCode%></p>
                            </div>

							<form name="OrderCancelForm" method="post">
								<input type="hidden" name="OrderCode"		value="<%=OrderCode%>" />
								<input type="hidden" name="PayType"			value="<%=PayType%>" />
								<input type="hidden" name="EscrowFlag"		value="<%=EscrowFlag%>" />
								<input type="hidden" name="TotalOrderCnt"	value="<%=TotalOrderCnt%>" />

                            <div class="h-line">
                                <h3 class="h-level4">취소할 상품</h3>
                            </div>

                            <ul class="informView">
<%
wQuery = ""
wQuery = wQuery & "WHERE A.IsShowFlag = 'Y' AND A.ProductType IN ('P','O') "
wQuery = wQuery & "AND A.OrderCode = '" & OrderCode & "' "
IF U_NUM <> "" THEN
		wQuery = wQuery & "AND B.UserID = '" & U_NUM & "' "
ELSE
		wQuery = wQuery & "AND B.OrderName = '" & N_NAME & "' AND B.OrderHp = '" & N_HP & "' AND B.OrderEmail = '" & N_EMAIL & "' "
END IF

sQuery = "ORDER BY A.OPIdx_Group, A.OPIdx_Org"

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
		Do Until oRs.EOF
				'# 결제완료 상태가 아니면 취소불가
				IF oRs("OrderState") = "3" AND oRs("CancelState1") = "0" AND oRs("CancelState2") = "0" THEN
						CancelableFlag	= "Y"
						IF oRs("ProductType") = "P" THEN
								CancelableCount	= CancelableCount + 1
						END IF
				ELSE
						CancelableFlag = "N"
				END IF

				IF oRs("ProductImage_180") = "" THEN
						ProductImage	= "/Images/180_noimage.png"
				ELSE
						ProductImage	= oRs("ProductImage_180")
				END IF
%>
                                <li class="informItem">
                                    <span class="cont">
										<%IF oRs("ProductType") = "P" AND CancelableFlag = "Y" THEN%>
										<span class="checkbox<%IF EscrowFlag = "Y" OR CStr(oRs("Idx")) = CStr(OPIdx) THEN%> is-checked<%END IF%>" style="z-index:1">
											<input type="checkbox" name="OPIdx" id="chk_<%=oRs("Idx")%>" value="<%=oRs("Idx")%>" onclick="getRefundPrice('C')" <%IF EscrowFlag = "Y" OR CStr(oRs("Idx")) = CStr(OPIdx) THEN%>checked="checked"<%END IF%> />
											<label for="chk_<%=oRs("Idx")%>"></label>
										</span>
										<%END IF%>

										<span class="thumbNail">
											<span class="img">
												<img src="<%=ProductImage%>" alt="상품 이미지">
											</span>
											<%IF CancelableFlag <> "Y" THEN%>
											<span class="about">
												<span class="process"><%=GetOrderState(oRs("OrderState"), oRs("CancelState1"), oRs("CancelState2"))%></span>
											</span>
											<%END IF%>
				                        </span>

										<span class="detail">
											<span class="brand">
												<span class="name"><%=oRs("BrandName")%></span>
												<%IF oRs("ProductType") = "O" THEN%><span class="oneplusone"><strong>[1+1]</strong></span><%END IF%>
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
												<%IF oRs("ProductType") = "P" THEN%>
												<span class="list">
													<span class="tit">결제금액</span>
													<span class="opt price"><em><%=FormatNumber(oRs("OrderPrice"),0)%></em>원</span>
												</span>
												<%END IF%>
												<span class="list">
													<span class="tit">배송</span>
													<span class="opt"><%IF oRs("DelvType") = "S" THEN%>매장픽업<%ELSE%>택배<%END IF%></span>
												</span>
											</span>
										</span>
                                    </span>
                                </li>
<%
				oRs.MoveNext
		Loop
END IF
oRs.Close
%>
                            </ul>

                            <div class="h-line">
                                <h3 class="h-level4">환불 금액 확인</h3>
                            </div>

                            <ul class="detailView on-right" id="RefundInfo">
                            </ul>

						<%IF PayType = "V" THEN%>
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
						<%END IF%>


						<%IF CancelableCount = 0 THEN%>
							<div class="inf-type1">
								<p class="tit">주문 취소 불가</p>
								<ul>
									<li class="bullet-ty1">주문상품 전체가 취소할 수 없는 상태 입니다.</li>
								</ul>
							</div>
						<%ELSEIF EscrowFlag = "Y" AND CancelableCount < TotalOrderCnt THEN%>
							<div class="inf-type1">
								<p class="tit">주문 취소 불가</p>
								<ul>
									<li class="bullet-ty1">에스크로 적용 주문은 전체상품 주문취소만 가능합니다.</li>
								</ul>
							</div>
						<%END IF%>

                        </div>
                    </div>

                    <div class="btns">
						<%IF (EscrowFlag <> "Y" AND CancelableCount > 0) OR (EscrowFlag = "Y" AND CancelableCount = TotalOrderCnt) THEN%>
                        <button type="button" onclick="orderCancel('C')" class="button ty-red">주문 취소하기</button>
						<%ELSE%>
                        <button type="button" onclick="closePop('DimDepth1')" class="button ty-red">목록으로</button>
						<%END IF%>
                    </div>
                </div>
            </div>
        </div>

		<script type="text/javascript">
			$(function () {
				$("#OrderCancel input[name='OPIdx']").on("click", function () {
					if ($(this).is(':checked')) {
						$(this).parent().addClass('is-checked');
					} else {
						$(this).parent().removeClass('is-checked');
					}
				});

				$("#OrderCancel select").on("change", function () {
					$(this).parent().find('.value').text($('option:selected', $(this)).text());
				});
			});
		</script>

<%
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>