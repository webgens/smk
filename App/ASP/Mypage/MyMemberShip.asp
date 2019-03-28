<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'MyMemberShip.asp - 마이페이지 > 회원정보 > 나의 멤버십
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
PageCode2 = "05"
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


DIM GroupName(10)
DIM StartAmount(10)
DIM EndAmount(10)
DIM MAmount(10)
DIM PointRate(10)

DIM TotalOrderPrice	: TotalOrderPrice	= 0
DIM GoalAmount		: GoalAmount		= 0


DIM MyPage_MemberGroup	: MyPage_MemberGroup	= ""
DIM MyPage_MemberGrade	: MyPage_MemberGrade	= ""

Dim SDate
Dim EDate
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

SDate = R_YEAR & "-" & R_MONTH & "-01"
EDate = DateAdd("d", -1, DateAdd("m", 1, SDate))


FOR i = 0 TO 10
		PointRate(i) = 0
NEXT


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성


'# 포인트 적립율
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Member_Group_Select"
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		Do Until oRs.EOF
				GroupName(oRs("GroupCode") - 1000)	 = oRs("GroupName")
				StartAmount(oRs("GroupCode") - 1000) = oRs("StartAmount")
				EndAmount(oRs("GroupCode") - 1000)	 = oRs("EndAmount")
				PointRate(oRs("GroupCode") - 1000)	 = oRs("PointRate")

				IF oRs("StartAmount") = 0 THEN
						MAmount(oRs("GroupCode") - 1000) = oRs("EndAmount") / 10000 & "만원 미만"
				ELSE
						MAmount(oRs("GroupCode") - 1000) = oRs("StartAmount") / 10000 & "만원 이상"
				END IF
				oRs.MoveNext
		Loop
END IF
oRs.Close

'# 년간 구매금액
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Order_Product_Select_For_Year_TotalOrderPrice"

		.Parameters.Append .CreateParameter("@MemberNum",		 adVarchar, adParaminput, 20	, U_NUM)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN	
		TotalOrderPrice		= oRs("TotalOrderPrice")
		IF ISNULL(TotalOrderPrice) THEN TotalOrderPrice = 0
		GoalAmount			= oRs("GoalAmount")
		IF ISNULL(GoalAmount) THEN GoalAmount = 0
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
%>


<!-- #include virtual="/INC/Header.asp" -->
	<style type="text/css">
		#OrderMenu .selector { margin-bottom: 0; }
		#OrderMenu .selector.is-focus .btn-list:after { background: url("/images/ico/ico_arrow_u2.png")no-repeat; background-size: 100% auto; }
	</style>
    <script type="text/javascript" src="/JS/dev/mypage.js?ver=<%=U_DATE%><%=U_TIME%>"></script>
	<script type="text/javascript" src="/JS/dev/login.js?ver=<%=U_DATE%><%=U_TIME%>"></script>

<%TopSubMenuTitle = "회원정보"%>
<!-- #include virtual="/INC/TopSub.asp" -->


    <!-- Main -->
    <main id="container" class="container">
        <div class="sub_content">

            <div class="wrap-mypage">
				<div style="height:8px"></div>


				
                        <div id="OrderMenu" class="ly-title accordion">
                            <div class="selector">
	                            <button type="button" class="btn-list clickEvt" data-target="OrderMenu">나의 멤버십</button>
							</div>
							<div class="option my-recode">
								<ul>
									<li><a href="/ASP/Mypage/MyMemberShip.asp">나의 멤버십</a></li>
									<li><a href="/ASP/Mypage/AddressList.asp">배송지관리</a></li>
									<li><a href="javascript:common_PopOpen('DimDepth1','MyInfoModify');">나의 정보 수정</a></li>
									<li><a href="/ASP/Mypage/SnsList.asp">SNS 계정설정</a></li>
								</ul>
							</div>
                        </div>



                <div class="mypage-membership">
                    <section id="contentList_1" class="accord-mypage">

                        <div class="ly-content1">
							<%IF U_MFLAG = "Y" THEN%>
                            <!-- 나의 멤버십 등급 -->
                            <div class="h-line">
                                <h2 class="h-level4">나의 멤버십 등급</h2>
                            </div>
                            <div class="membership">
                                <p class="grade grade_m_<%=MyPage_MemberGrade%>"><%=U_NAME%> 님의 등급은 <span class="bold"><%=MyPage_MemberGroup%></span>입니다</p>
                                <p class="accure">최근 1년간 총 누적금액 <span class="bold"><%=FormatNumber(TotalOrderPrice,0)%></span>원</p>
                                <p class="remain">다음 등급까지 <span class="bold"><%=FormatNumber(GoalAmount,0)%></span>원 남았습니다.</p>
                            </div>
                            <!-- 받은 등급 혜택 -->
                            <div class="h-line">
                                <h2 class="h-level4">받은 등급 혜택</h2>
                            </div>
                            <div class="grade-benefit">
                                <div class="cnt">
                                    <p class="date"></p>
                                    <p class="tit"><%=MyPage_MemberGroup%>의 구매급액별 포인트 적립율</p>
                                    <p></p>
                                    <p><span><%=PointRate(CDbl(U_GROUP) - 1000)%></span>%</p>
                                </div>
<%
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Coupon_Member_Select_By_Idx_For_Grade"

		.Parameters.Append .CreateParameter("@MemberNum",	 adInteger, adParaminput,	  , U_NUM)
		.Parameters.Append .CreateParameter("@SDate",		 adVarChar, adParaminput, 	50, SDate & " 00:00:00")
		.Parameters.Append .CreateParameter("@EDate",		 adVarChar, adParaminput, 	50, EDate & " 23:59:59")
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing
	
If Not oRs.EOF Then
	Do While Not oRs.EOF
%>
                                <div class="cnt">
                                    <p class="date"><%=Left(oRs("ReceiveDT"), 10)%></p>
                                    <p class="tit"><%=oRs("CouponName")%> 쿠폰 지급</p>
                                    <p><%=oRs("CouponName")%></p>
                                    <p><span><%=oRs("Discount")%></span><% If oRs("MoneyType") = "P" Then %>%<% Else %>원<% End If %></p>
                                </div>
<%
		oRs.MoveNext
	Loop
End If

Set oRs = oRs.NextRecordset

If Not oRs.EOF Then
	Do While Not oRs.EOF
%>
                                <div class="cnt">
                                    <p class="date"><%=Left(oRs("ReceiveDT"), 10)%></p>
                                    <p class="tit"><%=oRs("CouponName")%></p>
                                    <p>사이즈 무료교환 쿠폰</p>
                                    <p><span>1</span>장</p>
                                </div>
<%
		oRs.MoveNext
	Loop
End If
oRs.Close		
%>
                            </div>
							<%ELSE%>
                            <!-- 나의 멤버십 등급 -->
                            <div class="h-line">
                                <h2 class="h-level4">나의 멤버십 등급</h2>
                            </div>
                            <div class="membership">
                                <p class="grade"><%=U_NAME%> 님은 간편로그인 회원입니다.</p>
                                <p class="accure">최근 1년간 총 누적금액 <span class="bold"><%=FormatNumber(TotalOrderPrice,0)%></span>원</p>
                                <p class="remain">정회원 전환시 각종 쿠폰/혜택을 더 받아보실 수 있습니다.</p>
                            </div>
							<%END IF%>
                            <!-- 등급 상향 기준 안내 -->
                            <div class="h-line">
                                <h2 class="h-level4">등급별 상향 기준안내</h2>
                                <span class="h-date">고객 등급은 매월1일에 조정됩니다.</span>
                            </div>
                            <div class="grade-standard">
                                <div class="grade">
                                    <div class="tit-wrap">
                                        <p class="tit"><%=UCASE(GroupName(0))%></p>
                                        <p class="ratio">기본 적립율 <span class="bold"><%=PointRate(0)%>%</span></p>
                                    </div>
                                    <div class="explain">
                                        <p>대상 : <%=MAmount(0)%> 구매고객</p>
                                        <p>지급 : 10% 가입축하</p>
                                    </div>
                                </div>
                                <div class="grade">
                                    <div class="tit-wrap">
                                        <p class="tit"><%=UCASE(GroupName(1))%></p>
                                        <p class="ratio">기본 적립율 <span class="bold"><%=PointRate(1)%>%</span></p>
                                    </div>
                                    <div class="explain">
                                        <p>대상 : <%=MAmount(1)%> 구매고객</p>
                                        <p>지급 : 사이즈 무료교환 1장</p>
                                        <p>매월 : 6만원 이상 구매시 10% 할인쿠폰 X 2</p>
                                    </div>
                                </div>
                                <div class="grade">
                                    <div class="tit-wrap">
                                        <p class="tit"><%=UCASE(GroupName(2))%></p>
                                        <p class="ratio">기본 적립율 <span class="bold"><%=PointRate(2)%>%</span></p>
                                    </div>
                                    <div class="explain">
                                        <p>대상 : <%=MAmount(2)%> 구매고객</p>
                                        <p>지급 : 사이즈 무료교환 2장</p>
                                        <p>매월 : 6만원 이상 구매시 12% 할인쿠폰 X 2</p>
                                    </div>
                                </div>
                                <div class="grade">
                                    <div class="tit-wrap">
                                        <p class="tit"><%=UCASE(GroupName(3))%></p>
                                        <p class="ratio">기본 적립율 <span class="bold"><%=PointRate(3)%>%</span></p>
                                    </div>
                                    <div class="explain">
                                        <p>대상 : <%=MAmount(3)%> 구매고객</p>
                                        <p>지급 : 사이즈 무료교환 3장</p>
                                        <p>매월 : 8만원 이상 구매시 15% 할인쿠폰 X 2</p>
                                    </div>
                                </div>
                                <div class="grade">
                                    <div class="tit-wrap">
                                        <p class="tit"><%=UCASE(GroupName(4))%></p>
                                        <p class="ratio">기본 적립율 <span class="bold"><%=PointRate(4)%>%</span></p>
                                    </div>
                                    <div class="explain">
                                        <p>대상 : <%=MAmount(4)%> 구매고객</p>
                                        <p>지급 : 사이즈 무료교환 4장</p>
                                        <p>매월 : 8만원 이상 구매시 20% 할인쿠폰 X 2</p>
                                    </div>
                                </div>
                            </div>
                            <!-- 멤버십 혜택 안내 -->
                            <div class="h-line">
                                <h2 class="h-level4">슈마커 멤버십 혜택 안내</h2>
                                <span class="h-date">모든 등급의 슈마커 온라인 정회원에게 적용</span>
                            </div>
                            <div class="mbs-benefit">
                                <div class="cnt coupon">
                                    <div class="txt-area">
                                        <p>쿠폰</p>
                                        <%IF U_MFLAG = "Y" THEN%>
										<button type="button" onclick="APP_GoUrl('/ASP/Mypage/CouponList.asp')" class="button-ty3 ty-bd-black">
											<span>나의 보유 쿠폰</span>
										</button>
										<%END IF%>
                                    </div>
                                    <div class="explain">
										<p>앱 첫 로그인시 <span class="bold">3% 중복</span> 쿠폰</p>
                                        <p>첫 구매 감사 쿠폰 <span class="bold">5,000</span>원 지급 (2만원 이상 구매시)</p>
										<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="bold">10,000</span>원 지급 (5만원 이상 구매시)</p>
                                        <p>생일 축하 쿠폰 <span class="bold" style="margin-left:12px;">10,000</span>원 쿠폰 지급<br /> (6만원 이상 구매 시, 등록된 생일로 부터 7일전 발급)</p>
                                    </div>
                                </div>
                                <div class="cnt Scash">
                                    <div class="txt-area">
                                        <p>포인트</p>
                                        <%IF U_MFLAG = "Y" THEN%>
                                        <button type="button" onclick="APP_GoUrl('/ASP/Mypage/PointList.asp')" class="button-ty3 ty-bd-black">
											<span>나의 보유 포인트</span>
										</button>
										<%END IF%>
                                    </div>
                                    <div class="explain">
                                        <div class="cnt1">
                                            <p class="bold">상품 후기 작성</p>
                                            <p>일반후기 작성 : <span class="bold"><%=FormatNumber(MALL_REVIEW_POINT_B, 0)%></span> 포인트</p>
                                            <p>포토후기 작성 : <span class="bold"><%=FormatNumber(MALL_REVIEW_POINT_P, 0)%></span> 포인트</p>
                                        </div>
                                        <div class="cnt2">
                                            <p class="bold">출석체크 참여</p>
                                            <p>10일 : <span class="bold">500</span> 포인트</p>
                                            <p>15일 : <span class="bold">1,000</span> 포인트</p>
                                            <p>20일 : <span class="bold">2,000</span> 포인트</p>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </section>

                </div>
            </div>
        </div>
    </main>


	<!-- SNS계정 연결 공통 시작 -->
	<!-- SNS계정 로그인 Form -->
 	<form name="SimpleLoginForm" id="SimpleLoginForm" method="post">
		<input type="hidden" name="UID">
		<input type="hidden" name="Email">
		<input type="hidden" name="KName">
		<input type="hidden" name="SNSKind">
	</form>



	
<!-- #include virtual="/INC/FooterNoBNB.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->


<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>
