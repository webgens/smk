<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'ShoesGiftList.asp - 마이페이지 > 쇼핑혜택 > 슈즈상품권
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
PageCode2 = "11"
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

DIM SDate
DIM EDate

DIM MemberSCash
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


EDate			 = Date
SDate			 = DateAdd("m", -1, EDate)


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성


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
		MemberSCash = oRs("SCash")
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
	                    <button type="button" class="btn-list clickEvt" data-target="OrderMenu">슈즈상품권</button>
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

                    <div id="shoppingBenefit_2" class="accord-mypage">
                        <div class="ly-content1">
							
							<form name="formAddShoesGift" id="formAddShoesGift" method="post">
                            <div class="register-coupon">
                                <p class="tit">슈즈 상품권 등록</p>
                                <span class="input">
									<input type="text" name="cpno" id="cpno" maxlength="12" placeholder="쿠폰 번호를 입력해 주세요.">
								</span>
                                <button type="button" onclick="chk_ShoesGift()" class="button-ty2 is-expand ty-black">쿠폰 등록</button>
                            </div>
							</form>

                            <p class="possession">보유 슈즈 상품권 <em class="won"><%=FormatNumber(MemberSCash, 0)%><span>원</span></em></p>
							
							<form name="form" id="form">
							<input type="hidden" name="Page" id="Page" value="1" />
                            <div class="ly-calendar">
                                <div class="tit">
                                    <span>시작일</span>
                                    <span>종료일</span>
                                </div>
                                <div class="wrap">
                                    <div class="date-picker">
                                        <input type="text" name="SDate" id="SDate" value="<%=SDate%>" class="date-from" readonly="readonly">
                                    </div>
                                    <div class="date-picker">
                                        <input type="text" name="EDate" id="EDate" value="<%=EDate%>" class="date-to" readonly="readonly">
                                    </div>
                                </div>
                                <div class="area-radio">
                                    <span class="rad-ty1">
										<input type="radio" id="oneMonth" name="period_1" checked>
										<label for="oneMonth" onclick="setDate('1m', 'SDate', 'EDate')">1개월</label>
									</span>
                                    <span class="rad-ty1">
										<input type="radio" id="threeMonth" name="period_1">
										<label for="threeMonth" onclick="setDate('3m', 'SDate', 'EDate')">3개월</label>
									</span>
                                    <span class="rad-ty1">
										<input type="radio" id="sixMonth" name="period_1">
										<label for="sixMonth" onclick="setDate('6m', 'SDate', 'EDate')">6개월</label>
									</span>
                                </div>

                                <button type="button" onclick="get_ShoesGiftList(1)" class="button-ty2 is-expand ty-bd-gray">조회</button>
                            </div>
							</form>

							<div id="ShoesGiftList">
							</div>

                        </div>
                    </div>

                </div>



            </div>
        </div>
    </main>


	<script type="text/javascript">
		$(function () {
			$('#SDate').datepicker($.datepicker.regional['ko']);
			$('#EDate').datepicker($.datepicker.regional['ko']);


			get_ShoesGiftList(1);
		});
	</script>



<!-- #include virtual="/INC/FooterNoBNB.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->


<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>
