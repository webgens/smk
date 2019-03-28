<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'MYAddrList.asp - 마이페이지 > MY슈마커 > 상품평
'Date		: 2018.12.24
'Update		: 
'*****************************************************************************************'

'# 페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'-----------------------------------------------------------------------------------------'

'# 페이지 코드----------------------------------------------------------------------------'
DIM PageCode1, PageCode2, PageCode3, PageCode4
PageCode1 = "05"
PageCode2 = "03"
PageCode3 = "04"
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->
<!-- #include virtual="/Common/CheckID_Ajax.asp" -->

<%
'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn						'# ADODB Connection 개체
DIM oRs							'# ADODB Recordset 개체
DIM oCmd						'# ADODB Command 개체

DIM i
DIM j
DIM x
DIM y

DIM wQuery						'# WHERE 절
DIM sQuery						'# SORT 절


DIM Page
DIM PageSize : PageSize = 1000
DIM RecCnt
DIM PageCnt
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'



Page			 = Request("page")
If Page = "" Then Page = 1


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성

wQuery = "WHERE B.IsShowFlag = 'Y' AND B.ProductType IN ('P','O') AND B.OrderState = '7' "
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
		.CommandText		 = "USP_Mobile_EShop_Order_MyReview_Select"

		.Parameters.Append .CreateParameter("@PAGE",		 adInteger, adParaminput,		, Page)
		.Parameters.Append .CreateParameter("@PAGE_SIZE",	 adInteger, adParaminput,		, PageSize)
		.Parameters.Append .CreateParameter("@wQuery",		 adVarchar, adParaminput, 1000	, wQuery)
		.Parameters.Append .CreateParameter("@sQuery",		 adVarchar, adParaminput, 100	, sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

RecCnt	 = oRs(0)
PageCnt	 = FIX((RecCnt+(PageSize-1))/PageSize)

SET oRs = oRs.NextrecordSet

Response.Write "OK|||||"
%>
                            <p class="ad-area">포토 후기 작성 시 <span class="bold">3,000원</span> 지급</p>
                            <div class="h-line">
                                <h2 class="h-level4">나의 상품 후기</h2>
                            </div>
							<form name="MyReviewListForm" method="post">
							<input type="hidden" name="reviewGubun" />
							<input type="hidden" name="ordercode" />
							<input type="hidden" name="idx" />
                            <ul class="informView">
<%
IF NOT oRs.EOF THEN
	i = 1
	Do While Not oRs.EOF
		IF oRs("ReviewDT") <> "" THEN
%>
                                <li class="informItem">
                                    <a href="/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode") %>">
                                        <span class="head-tit">
                                            <span class="tit">구매일 : <%=GetDateYMD(oRs("OrderDate"))%></span>
                                            <span class="date">후기작성 : <%=oRs("ReviewDT") %></span>
                                        </span>
                                        <span class="cont">
                                            <span class="thumbNail">
                                                <span class="img">
                                                    <img src="<%=oRs("productImage")%>" alt="상품 이미지">
                                                </span>
                                                <span class="about">
                                                    <span class="process">배송완료</span>
                                                    <span class="date"><%=oRs("DelvDT")%></span>
                                                </span>
                                            </span>

                                            <span class="detail">
                                                <span class="brand">
                                                    <span class="name"><%=oRs("BrandName") %></span>
                                                    <span class="item-code"><%=oRs("ProdCD") %></span>
                                                </span>
                                                <span class="product-name"><em><%=oRs("ProductName") %></em></span>

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
                                                        <span class="opt price"><em><%=oRs("OrderPrice") %></em>원 <%IF oRs("DeliveryPrice") > 0 THEN %><span class="deliverFee">(배송비 <%=oRs("DeliveryPrice")%>원 포함)</span><%END IF%></span>
                                                    </span>
                                                    <span class="list">
                                                        <span class="tit">구분</span>
                                                        <span class="opt">슈마커배송??입점몰??</span>
                                                    </span>
                                                    <span class="list">
                                                        <span class="tit">배송/결제</span>
                                                        <span class="opt"><%IF oRs("DelvType") = "P" THEN%>택배<%ELSE%>매장픽업<%END IF%> / <%IF oRs("PayType") = "C" THEN%>신용카드<%ELSEIF oRs("PayType") = "B" THEN%>실시간 계좌이체<%ELSEIF oRs("PayType") = "V" THEN%>가상계좌<%ELSEIF oRs("PayType") = "V" THEN%>모바일<%END IF%></span>
                                                    </span>
                                                </span>
                                            </span>
                                        </span>
                                    </a>
                                    <div class="buttongroup">
                                        <button type="button" class="button-ty2 is-expand ty-bd-gray" onclick="insert_MyReview('view','<%=oRs("ordercode")%>','<%=oRs("idx")%>');">
                                            <p class="star-score">
                                                <span class="point val<%=int((round(oRs("ReViewGrade"),1)*2)+0.5) %>0"></span>
                                                <span class="score"><%=oRs("ReViewGrade")%></span>
                                            </p>후기 보기
                                        </button>
                                    </div>
                                </li>
<%
		ELSE
%>
                                <li class="informItem">
                                    <a href="/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode") %>">
                                        <span class="head-tit">
                                            <span class="tit">구매일 : <%=GetDateYMD(oRs("OrderDate"))%></span>
                                            <span class="date">작성 전</span>
                                        </span>
                                        <span class="cont">
                                            <span class="thumbNail">
                                                <span class="img">
                                                    <img src="<%=oRs("productImage")%>" alt="상품 이미지">
                                                </span>
                                                <span class="about">
                                                    <span class="process">배송완료</span>
                                                    <span class="date"><%=oRs("DelvDT") %></span>
                                                </span>
                                            </span>

                                            <span class="detail">
                                                <span class="brand">
                                                    <span class="name"><%=oRs("BrandName") %></span>
                                                    <span class="item-code"><%=oRs("ProdCD") %></span>
                                                </span>
                                                <span class="product-name"><em><%=oRs("ProductName") %></em></span>

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
                                                        <span class="opt price"><em><%=oRs("OrderPrice") %></em>원 <%IF oRs("DeliveryPrice") > 0 THEN %><span class="deliverFee">(배송비 <%=oRs("DeliveryPrice")%>원 포함)</span><%END IF%></span>
                                                    </span>
                                                    <span class="list">
                                                        <span class="tit">구분</span>
                                                        <span class="opt">슈마커배송??입점몰??</span>
                                                    </span>
                                                    <span class="list">
                                                        <span class="tit">배송/결제</span>
                                                        <span class="opt"><%IF oRs("DelvType") = "P" THEN%>택배<%ELSE%>매장픽업<%END IF%> / <%IF oRs("PayType") = "C" THEN%>신용카드<%ELSEIF oRs("PayType") = "B" THEN%>실시간 계좌이체<%ELSEIF oRs("PayType") = "V" THEN%>가상계좌<%ELSEIF oRs("PayType") = "V" THEN%>모바일<%END IF%></span>
                                                    </span>
                                                </span>
                                            </span>
                                        </span>
                                    </a>
                                    <div class="buttongroup">
                                        <button type="button" class="button-ty2 is-expand ty-bd-gray icon ico-inquire" onclick="insert_MyReview('add','<%=oRs("ordercode")%>','<%=oRs("idx")%>');">상품평 작성하기</button>
                                        <button type="button" class="button-ty2 is-expand ty-bd-gray icon ico-inquire" onclick="javascript:location.href='/ASP/Mypage/Ajax/ReviewWrite.asp?ordercode=<%=oRs("ordercode")%>&idx=<%=oRs("idx")%>&reviewType=add'">상품평 작성하기</button>
                                    </div>
                                </li>
<%
		END IF
		i = i + 1
		oRs.MoveNext
	Loop
ELSE
%>
                                <li class="informItem">
                                    <span class="head-tit">
                                        <span class="tit">상품후기 정보가 없습니다.</span>
                                    </span>
                                </li>
<%
END IF
%>
                            </ul>
							</form>

                            <div class="inf-type1">
                                <p class="tit">알려드립니다.</p>
                                <ul>
                                    <li class="bullet-ty1">포토후기 작성 시 2,000원 할인쿠폰 증정<br>(구매확정일로부터 30일 이내 작성에 한함.)</li>
                                    <li class="bullet-ty1">포토후기의 경우 직접 촬영한 사진이 아닐 경우 당첨과 쿠폰이 취소됩니다.</li>
                                    <li class="bullet-ty1">상품후기와 관련없는 내용일 경우 관리자에 의해 통보 없이 미등록, 삭제 될 수 있습니다.</li>
                                </ul>
                            </div>


<%
oRs.Close
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>