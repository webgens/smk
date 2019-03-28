<%@ Language=VBScript codepage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'NonDepositOrderCancel.asp - 입금전 주문취소 팝업
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

DIM ProductImage
DIM CancelableFlag		: CancelableFlag	= "Y"
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'
	
	
OrderCode		 = sqlFilter(Request("OrderCode"))


IF OrderCode = "" THEN
		Response.Write "FAIL|||||선택한 주문번호가 없습니다."
		Response.End
END IF


SET oConn				 = ConnectionOpen()							'# 커넥션 생성
SET oRs					 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성




Response.Write "OK|||||"
%>
        <div class="area-dim"></div>

        <div class="area-pop">
            <div class="full">
                <div class="tit-pop">
                    <div class="tit">입금전 주문취소</div>
                    <button type="button" onclick="closePop('DimDepth1')" class="btn-hide-pop">닫기</button>
                </div>

                <div class="container-pop mypage-ty1">
                    <div class="contents">
                        <div class="wrap-order">
                            <div class="order-number">
                                <p class="number">주문번호 : <%=OrderCode%></p>
                            </div>

                            <div class="h-line">
                                <h3 class="h-level4">주문 상품</h3>
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
				'# 미입금 상태가 아니면 취소불가
				IF oRs("OrderState") <> "1" THEN
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
										<span class="thumbNail">
											<span class="img">
												<img src="<%=ProductImage%>" alt="상품 이미지">
											</span>
				                        </span>

										<span class="detail">
											<a href="#">
												<span class="brand">
													<span class="name"><%=oRs("BrandName")%></span>
													<%IF oRs("ProductType") = "O" THEN%><span class="oneplusone"><strong>[1+1]</strong></span><%END IF%>
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
												<%IF oRs("ProductType") = "P" THEN%>
												<span class="list">
													<span class="tit">상품금액</span>
													<span class="opt price"><em><%=FormatNumber(oRs("OrderPrice"),0)%></em>원</span>
												</span>
												<%END IF%>
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

							<%IF CancelableFlag = "N" THEN%>
							<div class="inf-type1">
								<p class="tit">주문 취소 불가</p>
								<ul>
									<li class="bullet-ty1">주문취소할 수 없는 상태의 상품이 있습니다.</li>
								</ul>
							</div>
							<%END IF%>

                        </div>
                    </div>


                    <div class="btns">
						<%IF CancelableFlag = "Y" THEN%>
                        <button type="button" onclick="nonDepositOrderCancel('<%=OrderCode%>')" class="button ty-red">주문 취소하기</button>
						<%ELSE%>
                        <button type="button" onclick="closePop('DimDepth1')" class="button ty-red">목록으로</button>
						<%END IF%>
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