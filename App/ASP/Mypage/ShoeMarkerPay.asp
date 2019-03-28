<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'ShoeMarkerPay.asp - 마이페이지 > 나의 멤버십 > 슈마커PAY
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
PageCode2 = "14"
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
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
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
	                    <button type="button" class="btn-list clickEvt" data-target="OrderMenu">슈마커 PAY</button>
					</div>
					<div class="option my-recode">
						<ul>
							<li><a href="/ASP/Mypage/CouponList.asp">쿠폰북</a></li>
							<li><a href="/ASP/Mypage/ShoesGiftList.asp">슈즈상품권</a></li>
							<li><a href="/ASP/Mypage/MemberShip.asp">등급별 혜택 안내</a></li>
							<li><a href="/ASP/Mypage/PointList.asp">POINT</a></li>
							<li><a href="/ASP/Mypage/ShoeMarkerPay.asp">슈마커 PAY</a></li>
							<li></li>
						</ul>
					</div>
                </div>


                <div class="shopping-benefit">

                </div>



            </div>
        </div>
    </main>




<!-- #include virtual="/INC/FooterNoBNB.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->


<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>
