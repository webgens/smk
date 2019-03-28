<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'CouponList.asp - 마이페이지 > 나의 멤버십 > 쿠폰북
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
PageCode2 = "10"
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
	                    <button type="button" class="btn-list clickEvt" data-target="OrderMenu">쿠폰북</button>
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

                    <div id="shoppingBenefit_1" class="accord-mypage">

                        <div class="ly-content1">
							<form name="cForm" id="cForm" method="post">
                            <div class="register-coupon">
                                <p class="tit">쿠폰 등록</p>
                                <span class="input">
									<input type="text" name="CouponNum" id="CouponNum" maxlength="20" placeholder="쿠폰 번호를 입력해 주세요.">
								</span>
                                <button type="button" onclick="chg_Coupon()" class="button-ty2 is-expand ty-black">쿠폰 등록</button>
                            </div>
							</form>

                            <div id="tabs" class="tab">
                                <div class="tab-mypage">
                                    <ul class="tab-selector">
                                        <li class="part-2 "><a href="javascript:clk_CouponType(0);" data-target="tabs-col1">MY 쿠폰</a></li>
                                        <!-- 탭메뉴 갯수에 따른 클래스명 지정 : part-탭메뉴 갯수 예) 탭메뉴가 3개일 때, part-3 -->
                                        <li class="part-2"><a href="javascript:clk_CouponType(1);" data-target="tabs-col1">쿠폰북</a></li>
                                    </ul>
                                    <!-- 사용 가능 쿠폰 -->
                                    <div id="tabs-col1" class="tab-panel">
                                        <div class="h-line">
                                            <h2 class="h-level4">쿠폰수</h2>
                                            <span class="h-num"><span id="CouponCnt">0</span>개</span>

											<div class="selectbox">
												<div class="selector">
													<button type="button" class="btn-list clkCouponType">사용 가능한 쿠폰</button>
												</div>
												<div class="option2 my-recode-ct" style="display:none">
													<ul>
														<li data-useable="Y" class="on">사용 가능한 쿠폰</li>
														<li data-useable="N">지난 쿠폰</li>
													</ul>
												</div>
											</div>
                                        </div>

                                        <div class="ly-available" id="CouponList">
                                        </div>

                                        <div class="inf-type1">
                                            <p class="tit">알려드립니다.</p>
                                            <ul>
                                                <li class="bullet-ty1">할인쿠폰은 한 주문, 한 상품에 한해서 사용이 가능합니다.</li>
                                                <li class="bullet-ty1">할인쿠폰은 쿠폰마다 사용기간과 혜택(할인율, 금액) 내에서만 사용이 가능합니다.</li>
                                                <li class="bullet-ty1">쿠폰은 일부상품(이벤트/세일상품)에는 사용하실 수 없습니다.</li>
                                                <li class="bullet-ty1">주문 후 반품/환불/취소의 경우 한번 사용하신 할인쿠폰은 다시 사용하실 수 없습니다.</li>
                                            </ul>
                                        </div>
                                    </div>

                                </div>
                            </div>
                        </div>
                    </div>
                </div>



            </div>
        </div>
    </main>


	<form name="form" id="form">
		<input type="hidden" name="Page" id="Page" value="1" />
		<input type="hidden" name="Useable" id="Useable" value="Y" />
	</form>


	<script type="text/javascript">
		$(function () {
			get_CouponList(1);

			$(".clkCouponType").on("click", function () {
				if ($(".selectbox").hasClass("is-focus")) {
					$(".selectbox").removeClass("is-focus");
					$(".my-recode-ct").hide();
				}
				else {
					$(".selectbox").addClass("is-focus");
					$(".my-recode-ct").show();
				}
			});

			$(".my-recode-ct li").on("click", function () {
				$(".my-recode-ct li").removeClass("on");
				$(this).addClass("on");
				var txt = $(this).text();
				$(".clkCouponType").text(txt);
				$(".my-recode-ct").hide();
				$(".selectbox").removeClass("is-focus");

				var useable = $(this).data("useable");
				$("#Useable").val(useable);

				get_CouponList(1);
			});
		});
	</script>



<!-- #include virtual="/INC/FooterNoBNB.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->


<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>
