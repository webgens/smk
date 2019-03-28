<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'GradeInfo.asp - 등급혜택
'Date		: 2019.01.16
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
PageCode1 = "00"
PageCode2 = "00"
PageCode3 = "00"
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->

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


DIM mType						'# 회원정보 수정타입

DIM GroupName(10)
DIM StartAmount(10)
DIM EndAmount(10)
DIM MAmount(10)
DIM PointRate(10)

DIM TotalOrderPrice	: TotalOrderPrice	= 0
DIM GoalAmount		: GoalAmount		= 0


DIM MyPage_MemberGroup	: MyPage_MemberGroup	= ""

Dim SDate
Dim EDate
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

mType			 = sqlFilter(request("mType"))

SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성

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
%>


<!-- #include virtual="/INC/Header.asp" -->
	<style type="text/css">
		#OrderMenu .selector { margin-bottom: 0; }
		#OrderMenu .selector.is-focus .btn-list:after { background: url("/images/ico/ico_arrow_u2.png")no-repeat; background-size: 100% auto; }
	</style>
    <script type="text/javascript" src="/JS/dev/mypage.js?ver=<%=U_DATE%><%=U_TIME%>"></script>
	<script type="text/javascript" src="/JS/dev/login.js?ver=<%=U_DATE%><%=U_TIME%>"></script>
<!-- #include virtual="/INC/PopupTop.asp" -->


    <!-- Main -->
    <main id="container" class="container">
        <div class="sub_content" style="padding:0 !important;">

            <div class="wrap-mypage" style="border-top: 1px solid #e1e1e1;">


				


                <div class="mypage-membership">
                    <section id="contentList_1" class="accord-mypage">

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
                                        <!--<span class="my-grade">내등급</span>-->
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
                                <span class="h-date">모든 회원에게 지급되는 혜택</span>
                            </div>
                            <div class="mbs-benefit">
                                <div class="cnt coupon">
                                    <div class="txt-area">
                                        <p>쿠폰</p>
                                    </div>
                                    <div class="explain">
										<p>앱 첫 로그인시 <span class="bold">3% 중복</span>쿠폰</p>
                                        <p>첫 구매 감사 쿠폰 <span class="bold">5,000</span>원 지급 (2만원 이상 구매시)</p>
										<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<span class="bold">10,000</span>원 지급 (5만원 이상 구매시)</p>
                                        <p>생일 축하 <span class="bold">10,000</span>원 쿠폰 지급<br /> (6만원 이상 구매 시, 등록된 생일로 부터 7일전 발급)</p>
                                    </div>
                                </div>
                                <div class="cnt Scash">
                                    <div class="txt-area">
                                        <p>포인트</p>
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

<!-- #include virtual="/INC/FooterNone.asp" -->
<!-- #include virtual="/INC/PopupBottom.asp" -->



<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>
