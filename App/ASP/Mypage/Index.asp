<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'index.asp - 마이페이지
'Date		: 2018.12.17
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
PageCode2 = "00"
PageCode3 = "00"
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->
<!-- #include virtual="/Common/MyPageCheckID.asp" -->

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

DIM ProductImage
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성


DIM MyPage_Employee		: MyPage_Employee		= ""
DIM MyPage_MemberGroup	: MyPage_MemberGroup	= ""
DIM MyPage_MemberGrade	: MyPage_MemberGrade	= ""
DIM MyPage_MemberCoupon : MyPage_MemberCoupon	= 0
DIM MyPage_MemberPoint	: MyPage_MemberPoint	= 0
DIM MyPage_MemberSCash	: MyPage_MemberSCash	= 0



'# 회원 정보
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_Member_Select_By_MemberNum"
		
		.Parameters.Append .CreateParameter("@MemberNum", adInteger, adParamInput, , U_NUM)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		IF oRs("EmployeeFlag") = "Y" THEN
				SELECT CASE oRs("EmployeeType") 
					CASE "S"	: MyPage_Employee	= "슈마커 임직원 입니다"
					CASE "J"	: MyPage_Employee	= "JD 임직원 입니다"
					CASE ELSE	: MyPage_Employee	= ""
				END SELECT
		END IF
END IF
oRs.Close


'# 회원 그룹명
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_Member_Group_Select_By_GroupCode"
		
		.Parameters.Append .CreateParameter("@GroupCode", adInteger, adParamInput, , U_GROUP)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		MyPage_MemberGrade	= oRs("GroupCode")-999
		MyPage_MemberGroup = oRs("GroupName")
END IF
oRs.Close


	
'# 회원 보유 쿠폰(사용가능한 쿠폰만)
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Coupon_Member_Select_For_Useable_Count"
		
		.Parameters.Append .CreateParameter("@MemberNum", adInteger, adParamInput, , U_NUM)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		MyPage_MemberCoupon = oRs("CouponCount")
END IF
oRs.Close



'# 회원 보유 포인트
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_Member_Point_Select_For_Sum_By_MemberNum"
		
		.Parameters.Append .CreateParameter("@MemberNum", adInteger, adParamInput, , U_NUM)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		MyPage_MemberPoint = oRs("Point")
END IF
oRs.Close



'# 회원 보유 슈즈상품권
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_Member_SCash_Select_For_Sum_By_MemberNum"
		
		.Parameters.Append .CreateParameter("@MemberNum", adInteger, adParamInput, , U_NUM)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		MyPage_MemberSCash = oRs("SCash")
END IF
oRs.Close
%>


<!-- #include virtual="/INC/Header.asp" -->
    <script type="text/javascript" src="/JS/dev/mypage.js?ver=<%=U_DATE%><%=U_TIME%>"></script>
    <script type="text/javascript" src="/JS/dev/login.js?ver=<%=U_DATE%><%=U_TIME%>"></script>
<!-- #include virtual="/INC/TopMypage.asp" -->


    <!-- Main -->
    <main id="container" class="container">
        <div class="content">

            <div class="wrap-mypage">
                <div class="mypage-main">
                    <%IF U_MFLAG = "Y" THEN%>
                    <section class="my-information">
                        <div class="inform grade_m_<%=MyPage_MemberGrade%>">
                            <h2 class="grade"><%=U_NAME%>님의 등급은 <em class="level"><%=MyPage_MemberGroup%></em>입니다.</h2>
                            <div class="etc">
                                <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/MyMemberShip.asp');">나의 멤버십</a>
								<a href="javascript:sm_Logout()">로그아웃</a>
							</div>
                        </div>

                        <div class="have-benefit">
                            <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/CouponList.asp')" class="coupon">나의 쿠폰 <em class="num"><%=FormatNumber(MyPage_MemberCoupon, 0)%></em></a>
                            <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/PointList.asp')" class="point">포인트 <em class="num"><%=FormatNumber(MyPage_MemberPoint, 0)%></em></a>
                            <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/ShoesGiftList.asp')" class="gift-card">슈즈 상품권 <em class="num"><%=FormatNumber(MyPage_MemberSCash, 0)%></em></a>
                        </div>
                    </section>
					<%ELSE%>
                    <section class="my-information">
                        <div class="inform">
                            <h2 class="grade"><%=U_NAME%>님은 <em class="level">간편로그인</em>으로 로그인 하셨습니다.</h2>
                            <div class="etc">
                                <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Member/JoinChgMem.asp')">정회원 전환</a>
								<a href="javascript:sm_Logout()">로그아웃</a>
							</div>
                        </div>
                    </section>
					<%END IF%>

<%
DIM OrderState_1	: OrderState_1	= 0		'# 주문접수
DIM OrderState_3	: OrderState_3	= 0		'# 결제완료
DIM OrderState_4	: OrderState_4	= 0		'# 상품준비중
DIM OrderState_5	: OrderState_5	= 0		'# 배송중
DIM OrderState_6	: OrderState_6	= 0		'# 배송완료

wQuery = "WHERE B.IsShowFlag = 'Y' AND B.ProductType = 'P' AND B.OrderState IN ('1', '3', '4', '5', '6', '7', 'C') "
wQuery = wQuery & "AND A.OrderDate >= '" & Replace(DateAdd("m", -1, Date), "-", "") & "' "
wQuery = wQuery & "AND A.OrderDate <= '" & Replace(Date, "-", "") & "' "
wQuery = wQuery & "AND A.UserID = '" & U_NUM & "' "


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
END IF
oRs.Close
%>

                    <section class="my-order-list">
                        <div class="order-list">
                            <div class="h-line">
                                <h3 class="h-level5">최근 1개월간 주문 내역</h3>
                                <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/OrderList.asp')" class="all-view is-right">전체 내역 보기</a>
                            </div>

                            <ol>
                                <li<%IF OrderState_1 > 0 THEN%> class="current"<%END IF%>>
                                    <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/OrderList.asp?SOrderState=1')"><span class="num"><%=FormatNumber(OrderState_1,0)%></span>주문접수</a>
                                </li>
                                <li<%IF OrderState_3 > 0 THEN%> class="current"<%END IF%>>
                                    <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/OrderList.asp?SOrderState=3')"><span class="num"><%=FormatNumber(OrderState_3,0)%></span>결제완료</a>
                                </li>
                                <li<%IF OrderState_4 > 0 THEN%> class="current"<%END IF%>>
                                    <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/OrderList.asp?SOrderState=4')"><span class="num"><%=FormatNumber(OrderState_4,0)%></span>상품준비</a>
                                </li>
                                <li<%IF OrderState_5 > 0 THEN%> class="current"<%END IF%>>
                                    <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/OrderList.asp?SOrderState=5')"><span class="num"><%=FormatNumber(OrderState_5,0)%></span>배송중</a>
                                </li>
                                <li<%IF OrderState_6 > 0 THEN%> class="current"<%END IF%>>
                                    <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/OrderList.asp?SOrderState=6')"><span class="num"><%=FormatNumber(OrderState_6,0)%></span>배송완료</a>
                                </li>
                            </ol>
                        </div>


						<%IF U_MFLAG = "Y" THEN%>
                        <ul class="quick-menu">
                            <li class="menu-1">
                                <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/OrderList.asp')">주문/배송</a>
                            </li>
                            <li class="menu-2">
                                <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/OrderCRXList.asp')">교환/반품</a>
                            </li>
                            <li class="menu-3">
                                <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/OrderASList.asp')">A/S신청</a>
                            </li>
                            <li class="menu-4">
                                <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/MyReview.asp')">상품리뷰</a>
                            </li>
                            <li class="menu-5">
                                <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/Qna.asp?QnaType=2')">1:1 문의내역</a>
                            </li>
                            <li class="menu-6">
                                <a href="javascript:common_PopOpen('DimDepth1','MyInfoModify');">개인정보 수정</a>
                            </li>
                        </ul>
						<%END IF%>

                    </section>
					
					
					
					<div class="h-level5">마이페이지 전체 메뉴</div>
                    <section class="all-mypage-menu">
						<%IF U_MFLAG = "Y" THEN%>
                        <div id="depth_1" class="accord-mypage">
                            <div class="ly-title">
                                <button type="button" class="btn-list clickEvt" data-target="depth_1">쇼핑내역</button>
                            </div>
                            <div class="ly-content">
                                <ul class="sub-menu">
                                    <li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/OrderList.asp')">주문/배송조회/영수증발급</a></li>
                                    <li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/OrderCRXList.asp')">주문취소/반품/교환</a></li>
                                    <li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/OrderASList.asp')">A/S신청</a></li>
                                </ul>
                            </div>
                        </div>

                        <div id="depth_2" class="accord-mypage">
                            <div class="ly-title">
                                <button type="button" class="btn-list clickEvt" data-target="depth_2">MY SHOEMARKER</button>
                            </div>
                            <div class="ly-content">
                                <ul class="sub-menu">
                                    <li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/MyPickList.asp')">MY&hearts;</a></li>
                                    <li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/MyReentry.asp')">재입고알림</a></li>
                                    <li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/MyReview.asp')">상품후기</a></li>
                                    <li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/Qna.asp')">상품문의</a></li>
                                </ul>
                            </div>
                        </div>
                        <div id="depth_3" class="accord-mypage">
                            <div class="ly-title">
                                <button type="button" class="btn-list clickEvt" data-target="depth_3">쇼핑혜택</button>
                            </div>
                            <div class="ly-content">
                                <ul class="sub-menu">
                                    <li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/CouponList.asp')">쿠폰북</a></li>
                                    <li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/ShoesGiftList.asp')">슈즈 상품권</a></li>
                                    <li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/PointList.asp')">포인트</a></li>
                                    <!--<li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/ShoeMarkerPay.asp')">슈마커 PAY</a></li>-->
                                </ul>
                            </div>
                        </div>
                        <div id="depth_4" class="accord-mypage">
                            <div class="ly-title">
                                <button type="button" class="btn-list clickEvt" data-target="depth_4">회원정보 관리</button>
                            </div>
                            <div class="ly-content">
                                <ul class="sub-menu">
                                    <li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/MyMemberShip.asp')">나의 멤버십</a></li>
                                    <li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/AddressList.asp')">배송지 관리</a></li>
                                    <li><a href="javascript:common_PopOpen('DimDepth1','MyInfoModify');">나의 정보 수정</a></li>
                                    <li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/SnsList.asp')">SNS 계정설정</a></li>
                                </ul>
                            </div>
                        </div>
						<%ELSE%>
                        <div id="depth_1" class="accord-mypage">
                            <div class="ly-title">
                                <button type="button" class="btn-list clickEvt" data-target="depth_1">쇼핑내역</button>
                            </div>
                            <div class="ly-content">
                                <ul class="sub-menu">
                                    <li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/OrderList.asp')">주문/배송조회/영수증발급</a></li>
                                    <li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/OrderCRXList.asp')">주문취소/반품/교환</a></li>
                                    <li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/OrderASList.asp')">A/S 신청내역</a></li>
                                </ul>
                            </div>
                        </div>

                        <div id="depth_2" class="accord-mypage">
                            <div class="ly-title">
                                <button type="button" class="btn-list clickEvt" data-target="depth_2">MY SHOEMARKER</button>
                            </div>
                            <div class="ly-content">
                                <ul class="sub-menu">
                                    <li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/MyPickList.asp')">MY&hearts;</a></li>
                                    <li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/MyReview.asp')">상품후기</a></li>
                                    <li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/Qna.asp')">상품문의</a></li>
                                </ul>
                            </div>
                        </div>
                        <div id="depth_3" class="accord-mypage">
                            <div class="ly-title">
                                <button type="button" class="btn-list clickEvt" data-target="depth_3">쇼핑혜택</button>
                            </div>
                            <div class="ly-content">
                                <ul class="sub-menu">
                                    <li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/MemberShip.asp')">등급별 혜택 안내</a></li>
                                    <!--<li><a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Mypage/ShoeMarkerPay.asp')">슈마커 PAY</a></li>-->
                                </ul>
                            </div>
                        </div>
						<%END IF%>

                    </section>
                </div>
            </div>
        </div>
    </main>


<!-- #include virtual="/INC/Footer.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>
