<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'MemberShip.asp - 마이페이지 > 나의 멤버십 > 등급별 혜택 안내
'Date		: 2019.01.07
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
PageCode3 = "01"
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



DIM GroupName(10)
DIM StartAmount(10)
DIM EndAmount(10)
DIM MAmount(10)
DIM PointRate(10)
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'



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
		.container .content { min-height: 100%; padding: 80px 0 195px; }
		#OrderMenu .selector { margin-bottom: 0; }
		#OrderMenu .selector.is-focus .btn-list:after { background: url("/images/ico/ico_arrow_u2.png")no-repeat; background-size: 100% auto; }

		.h-line { overflow: initial !important;z-index:2; }
		.selectbox { position: absolute; top: 4px; right: 12px; width: 130px; height: 32px; }
		.selectbox .selector { height:32px !important; border: 1px solid #c8c8c8; margin-bottom:0 !important; }
		.selectbox .btn-list { position: relative; width: 100%; height: 30px; padding: 0 12px; font-size: 12px; color: #282828; line-height: 30px; text-indent: 3px; text-align: left; background-color: #fff; padding-left: 10px; font-size: 12px; }
		.selectbox .btn-list:after { content: ''; display: block; position: absolute; right: 15px; top: 50%; -webkit-transform: translateY(-50%); transform: translateY(-50%); width: 8px; height: 5px; background: url(/Images/ico/ico_arrow_d1.png) no-repeat; background-size: 100% auto; }
		
		.selectbox .my-recode-ct { background: #fff; z-index:5; }
		.selectbox .my-recode-ct li.on { background: #f1f1f1; }
		.selectbox .my-recode-ct li { width: 100%; height: 30px; line-height: 30px; border-left: 1px solid #c8c8c8; border-right: 1px solid #e6e6e6; border-bottom: 1px solid #e6e6e6; box-sizing: border-box; font-size: 12px; padding-left: 10px; cursor: pointer; }


		.ly-title .btn-list:after, .ly-mtitle_sub .btn-list:after { content: ''; display: block; position: absolute; right: 15px; top: 50%; -webkit-transform: translateY(-50%); transform: translateY(-50%); width: 8px; height: 5px; background: initial; background-size: 100% auto; }
	</style>
    <script type="text/javascript" src="/JS/dev/mypage.js?ver=<%=U_DATE%><%=U_TIME%>"></script>
<!-- #include virtual="/INC/TopMypage.asp" -->


    <!-- Main -->
    <main id="container" class="container">
        <div class="content">

            <div class="wrap-mypage">
				<div style="height:8px"></div>


				
                <div id="OrderMenu" class="ly-title accordion">
                    <div class="selector">
	                    <button type="button" class="btn-list" data-target="OrderMenu">등급별 혜택 안내</button>
					</div>
                </div>



                <div class="mypage-membership">
                    <section id="contentList_1" class="accord-mypage">
						
                        <div class="ly-content1">
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




<!-- #include virtual="/INC/Footer.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->


<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>
