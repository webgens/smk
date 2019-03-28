<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'TimeSale.asp - 타임세일
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
DIM oRs1						'# ADODB Recordset 개체
DIM oCmd						'# ADODB Command 개체

DIM i
DIM j
DIM x
DIM y

DIM wQuery						'# WHERE 절
DIM sQuery						'# SORT 절

Dim TimeSaleCount
Dim TimeSale
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성

SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Event_Category_Product_Select_By_TimeSale"
		.Parameters.Append .CreateParameter("@Today",		 adVarchar, adParaminput, 12	, R_YEAR & R_MONTH & R_DAY & R_HOUR & R_MIN)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

TimeSaleCount = oRs.RecordCount

TimeSale = R_YEAR & "-" & R_MONTH & "-" & R_DAY & "T" & R_HOUR & ":00:00"

IF NOT oRs.EOF THEN
	TimeSale = Left(oRs("EDate"), 4) & "-" & Mid(oRs("EDate"), 5, 2) & "-" & Mid(oRs("EDate"), 7, 2) & "T" & Mid(oRs("EDate"), 9, 2) & ":00:00"
	Response.Write "OK|||||"
%>
            <div class="ly-timeSale">
                <div class="tit">
                    <span class="head">TIME SALE</span>
                    <div class="cnt">
                        <p>남은시간</p>
                        <div id="timesale" class="time-sale"></div>
                    </div>
                    <button type="button" class="btn-hide" onclick="close_TimeSale();">닫기</button>
                </div>

                <div class="content">
                    <div class="timeSale-slide">
                        <div class="swiper-container">
                            <div class="swiper-wrapper">
							<%
								Do While Not oRs.EOF	
							%>
                                <div class="swiper-slide">
                                    <!--<span class="item-length">3개 한정</span>-->
                                    <div class="thumbNail">
                                        <img src="<%=oRs("ImageUrl")%>" alt="<%=oRs("ProductName")%>" onclick="javascript:location.href='/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>';">
                                    </div>

                                    <div class="cont">
                                        <div class="brand"><%=oRs("BrandName")%></div>
                                        <p class="item-name"><%=oRs("ProductName")%></p>
                                        <span class="delete"><%=FormatNumber(oRs("TagPrice"), 0)%>원</span>
                                        <div class="price"><em><%=FormatNumber(oRs("SalePrice"), 0)%></em>원</div>
                                        <div class="discount-rate"><em><%=FormatNumber(oRs("DiscountRate"), 0)%></em>%</div>
                                    </div>

                                    <a href="/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>" class="a-more">상품 상세보기</a>
                                </div>
							<%
									oRs.MoveNext
								Loop	
							%>
                            </div>
                            <div class="swiper-pagination"></div>
                        </div>
                    </div>

                </div>
            </div>

			<script type="text/javascript">
				//timesale
				function timesaleNew() {
					var now = new Date().getTime();
					var countDownDate = new Date('<%=TimeSale%>');

					var distance = countDownDate - now;

					var hours = Math.floor((distance % (1000 * 60 * 60 * 60 * 24)) / (1000 * 60 * 60));
					var minutes = Math.floor((distance % (1000 * 60 * 60)) / (1000 * 60));
					var seconds = Math.floor((distance % (1000 * 60)) / 1000);

					var h = hours.toString();
					var m = minutes.toString();
					var s = seconds.toString();

					if (h.length == 1) { h = "0" + h }
					if (m.length == 1) { m = "0" + m }
					if (s.length == 1) { s = "0" + s }

					document.getElementById('timesale').innerHTML = h + ':' + m + ':' + s;
				}

				var swiper = new Swiper('.timeSale-slide .swiper-container', {
					slidesPerView: 1,
					<% If TimeSaleCount > 1 Then %>loop: true,<% End If %>
					direction: 'horizontal',
					autoplay: {
						delay: 2500,
						disableOnInteraction: false,
					},
					pagination: {
						el: '.swiper-pagination',
						clickable: true,
					},
					observer: true,
					observeParents: true
				});

			</script>

			<script type="text/javascript">
				$(document).ready(function () {
					setInterval(timesaleNew, 1);
				});
			</script>

<%
Else
	Response.Write "NONE|||||"
End If
oRs.Close

SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>





