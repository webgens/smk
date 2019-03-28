<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'PointList.asp - 마이페이지 > 나의 멤버십 > 포인트
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
PageCode2 = "13"
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

DIM MemberPoint
	
DIM GroupName(10)
DIM PointRate(10)
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


FOR i = 0 TO 10
		PointRate(i) = 0
NEXT


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성


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
		MemberPoint = oRs("Point")
END IF
oRs.Close


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
				GroupName(oRs("GroupCode") - 1000) = oRs("GroupName")
				PointRate(oRs("GroupCode") - 1000) = oRs("PointRate")
				oRs.MoveNext
		Loop
END IF
oRs.Close
%>


<!-- #include virtual="/INC/Header.asp" -->
	<style type="text/css">
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
	</style>
    <script type="text/javascript" src="/JS/dev/mypage.js?ver=<%=U_DATE%><%=U_TIME%>"></script>

<%TopSubMenuTitle = "쇼핑혜택"%>
<!-- #include virtual="/INC/TopSub.asp" -->


    <!-- Main -->
    <main id="container" class="container">
        <div class="sub_content">

            <div class="wrap-mypage">
				<div style="height:8px"></div>


				
                <div id="OrderMenu" class="ly-title accordion">
                    <div class="selector">
	                    <button type="button" class="btn-list clickEvt" data-target="OrderMenu">포인트</button>
					</div>
					<div class="option my-recode">
						<ul>
							<li><a href="/ASP/Mypage/CouponList.asp">쿠폰북</a></li>
							<li><a href="/ASP/Mypage/ShoesGiftList.asp">슈즈상품권</a></li>
							<li><a href="/ASP/Mypage/PointList.asp">포인트</a></li>
							<li><!--<a href="/ASP/Mypage/ShoeMarkerPay.asp">슈마커 PAY</a>--></li>
						</ul>
					</div>
                </div>



                <div class="shopping-benefit">

                     <div id="shoppingBenefit_3" class="accord-mypage">
                        <div class="ly-content1">
                            <p class="possession">보유 포인트 <em class="won"><%=FormatNumber(MemberPoint, 0)%><span>원</span></em></p>


							<div id="PointList">
							</div>

                            <div class="h-line">
                                <h2 class="h-level4">회원 등급별 포인트 적립율</h2>
                            </div>

                            <div class="grade-saving-rate">
                                <ol>
                                    <li class="bronze">
                                        <div class="tit"><%=UCASE(GroupName(0))%></div>
                                        <div class="rate"><em><%=PointRate(0)%></em>%</div>
                                    </li>
                                    <li class="silver">
                                        <div class="tit"><%=UCASE(GroupName(1))%></div>
                                        <div class="rate"><em><%=PointRate(1)%></em>%</div>
                                    </li>
                                    <li class="gold">
                                        <div class="tit"><%=UCASE(GroupName(2))%></div>
                                        <div class="rate"><em><%=PointRate(2)%></em>%</div>
                                    </li>
                                    <li class="vip">
                                        <div class="tit"><%=UCASE(GroupName(3))%></div>
                                        <div class="rate"><em><%=PointRate(3)%></em>%</div>
                                    </li>
                                    <li class="vvip">
                                        <div class="tit"><%=UCASE(GroupName(4))%></div>
                                        <div class="rate"><em><%=PointRate(4)%></em>%</div>
                                    </li>
                                </ol>
                            </div>

                            <div class="inf-type1">
                                <p class="tit">알려드립니다.</p>
                                <ul>
                                    <li class="bullet-ty1">포인트는 상품 구매금액이 1,000원 이상부터 10원 단위로 적립/사용됩니다.</li>
                                    <li class="bullet-ty1">1일 최대 25,000원까지 포인트가 적립됩니다.</li>
                                    <li class="bullet-ty1">출석체크 참여시 10일 500포인트, 15일 1,000포인트, 20일 2,000포인트가 적립됩니다.</li>
                                    <li class="bullet-ty1">포인트 지급이 있는 이벤트 참여 시 적립됩니다.</li>
                                    <li class="bullet-ty1">보유하신 포인트는 상품구매 시 10원 단위부터 현금처럼 결제 시 사용하실 수 있습니다.</li>
                                    <li class="bullet-ty1">사용기한은 적립이 발생한 달로부터 1년이 경과한 달의 1일 자정에 소멸됩니다.</li>
                                </ul>
                            </div>
                        </div>
                    </div>


                </div>



            </div>
        </div>
    </main>



		<script type="text/javascript">
			$(function () {
				get_PointList(1);
			});
		</script>



<!-- #include virtual="/INC/FooterNoBNB.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->


<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>
