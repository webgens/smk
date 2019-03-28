<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'MyReview.asp - 상품후기
'Date		: 2019.01.06
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
PageCode3 = "00"
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->
<!-- #include virtual="/Common/SubCheckID.asp" -->

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
DIM PageSize : PageSize = 10
DIM RecCnt
DIM PageCnt
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

Page			 = Request("page")
If Page = "" Then Page = 1


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>


<!-- #include virtual="/INC/Header.asp" -->
	<script type="text/javascript" src ="/ASP/Mypage/JS/Order.js?ver=<%=U_DATE & U_TIME%>"></script>

<%TopSubMenuTitle = "MY슈마커"%>
<!-- #include virtual="/INC/TopSub.asp" -->


    <!-- Main -->
    <main id="container" class="container">
        <div class="sub_content">

            <div class="wrap-mypage">
				<div style="height:8px"></div>
				
                <div id="MypageSubMenu" class="ly-title accordion">
                    <div class="selector">
	                    <button type="button" class="btn-list clickEvt" data-target="MypageSubMenu">상품후기</button>
					</div>
					<div class="option my-recode">
						<!-- #include virtual="/ASP/Mypage/SubMenu_MyShoeMarker.asp" -->
					</div>
                </div>

                <div class="mypage-my-brand">
                    <div id="shoppingList">
                        <div>
                            <p class="ad-area">포토 후기 작성 시 <span class="bold"><%=FormatNumber(MALL_REVIEW_POINT_P,0)%>원</span> 포인트 증정</p>
                            <div class="h-line">
                                <h2 class="h-level4">나의 상품 후기</h2>
                            </div>
<%
wQuery = "WHERE B.IsShowFlag = 'Y' AND B.ProductType = 'P' AND B.OrderState = '7' "
wQuery = wQuery & "AND A.UserID = '" & U_NUM & "' "

sQuery = "ORDER BY A.OrderCode DESC, B.Idx "

SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Order_MyReview_Select"

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
%>
<%
IF NOT oRs.EOF THEN
%>
                            <ul class="informView">
<%
		Do Until oRs.EOF
%>
                                <li class="informItem">
                                    <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>')">
                                        <span class="head-tit">
                                            <span class="tit">구매일 : <%=GetDateYMD(oRs("OrderDate"))%></span>
											<%IF oRs("ReviewDT") <> "" THEN%>
                                            <span class="date">후기작성 : <%=oRs("ReviewDT")%></span>
											<%END IF%>
                                        </span>
                                        <span class="cont">
                                            <span class="thumbNail">
                                                <span class="img">
                                                    <img src="<%=oRs("productImage_180")%>" alt="상품 이미지">
                                                </span>
                                            </span>

                                            <span class="detail">
                                                <span class="brand">
                                                    <span class="name"><%=oRs("BrandName")%></span>
                                                </span>
                                                <span class="product-name"><em><%=oRs("ProductName") %></em></span>

                                                <span class="inform">
                                                    <span class="list">
                                                        <span class="tit">옵션</span>
                                                        <span class="opt"><%=oRs("SizeCD") %></span>
                                                    </span>
                                                </span>
                                            </span>
                                        </span>
                                    </a>

                                    <div class="buttongroup">
										<%IF oRs("ReviewDT") <> "" THEN%>
                                        <button type="button" onclick="openReview('<%=oRs("ReviewIdx")%>')" class="button-ty2 is-expand ty-bd-gray">
                                            <p class="star-score">
                                                <span class="point val<%=GetStarGrade(oRs("ReviewGrade"))%>"></span>
                                                <span class="score"><%=FormatNumber(oRs("ReviewGrade"),1) %></span>
                                            </p>후기 보기
                                        </button>
										<%ELSE%>
                                        <button type="button" onclick="openReviewWrite('<%=oRs("OrderCode")%>','<%=oRs("Idx")%>');" class="button-ty2 is-expand ty-bd-gray icon ico-inquire">상품후기 작성하기</button>
										<%END IF%>
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
							<div class="area-empty">
								<span class="icon-empty"></span>
								<p class="tit-empty">작성한 상품후기가 없습니다.</p>
							</div>
<%
END IF
oRs.close
%>
                            <div class="inf-type1" style="padding-bottom:20px">
                                <p class="tit">알려드립니다.</p>
                                <ul>
                                    <li class="bullet-ty1">포토후기 작성 시 <%=FormatNumber(MALL_REVIEW_POINT_P,0)%>원 할인쿠폰 증정<br>(구매확정일로부터 30일 이내 작성에 한함.)</li>
                                    <li class="bullet-ty1">포토후기의 경우 직접 촬영한 사진이 아닐 경우 당첨과 쿠폰이 취소됩니다.</li>
                                    <li class="bullet-ty1">상품후기와 관련없는 내용일 경우 관리자에 의해 통보 없이 미등록, 삭제 될 수 있습니다.</li>
                                </ul>
                            </div>
                        </div>	<!-- -->
                    </div>	<!--shoppingList-->
                </div>	<!--my-re-entry-->
            </div>	<!--wrap-mypage-->
        </div> <!--content-->
    </main>

<!-- #include virtual="/INC/FooterNoBNB.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>