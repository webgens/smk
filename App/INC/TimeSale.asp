<%
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

Dim TimeSale
TimeSale = R_YEAR & "-" & R_MONTH & "-" & R_DAY & "T" & R_HOUR & ":00:00"

Dim TimeSaleCount
TimeSaleCount = oRs.RecordCount

IF NOT oRs.EOF THEN
	TimeSale = Left(oRs("EDate"), 4) & "-" & Mid(oRs("EDate"), 5, 2) & "-" & Mid(oRs("EDate"), 7, 2) & "T" & Mid(oRs("EDate"), 9, 2) & ":00:00"
%>
   	<section class="wrap-pop timesale">
        <!-- 수정 190116 : 클래스 추가 -->
        <div class="area-dim"></div>

        <div class="area-pop">
            <div class="ly-timeSale">
                <div class="tit">
                    <span class="head">TIME SALE</span>
                    <div class="cnt">
                        <p>남은시간</p>
                        <div id="timesale" class="time-sale"></div>
                    </div>
                    <button type="button" class="btn-hide" onclick="TimeSaleClose();">닫기</button>
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
                                        <div class="discount-rate"><em>33</em>%</div>
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
        </div>
    </section>
    <!--// PopUp -->

    <!-- 수정 190116 : 타임세일 아이콘 추가 -->
    <div class="ico-timesale">
        <button type="button">TIME<br>SALE</button>
    </div>
    <!-- // 수정 190116 : 타임세일 아이콘 추가 -->

	<script type="text/javascript">
		//timesale
		function timesaleNew() {
			var now = new Date().getTime();
//			var countDownDate = new Date('<%=TimeSale%>');
			var countDownDate = new Date('<%=TimeSale%>' + 'Z');
			countDownDate = countDownDate.getTime() + countDownDate.getTimezoneOffset() * 60 * 1000


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

		function TimeSaleClose() {
			var _this = $(".btn-hide");
			_this.closest('.wrap-pop').next().removeClass('hide');
			_this.closest('.wrap-pop').css('display', 'none');
			_this.closest('body').removeClass('posFixed');
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
End If
oRs.Close	
%>