<%
If Trim(Request.Cookies("SlidePopupToday")) = "" Then
	SET oCmd = Server.CreateObject("ADODB.Command")
	WITH oCmd
			.ActiveConnection	 = oConn
			.CommandType		 = adCmdStoredProc
			.CommandText		 = "USP_Mobile_EShop_Slide_Popup_Select_By_Today"
			.Parameters.Append .CreateParameter("@Today",		 adVarchar, adParaminput, 12	, R_YEAR & R_MONTH & R_DAY & R_HOUR & R_MIN)
	END WITH
	oRs.CursorLocation = adUseClient
	oRs.Open oCmd, , adOpenStatic, adLockReadOnly
	SET oCmd = Nothing

	Dim TodaySlidePopupCount
	TodaySlidePopupCount = oRs.RecordCount

	If Not oRs.eof Then
%>        
        <div class="area-pop" id="slidePopup" style="height:0px;">
            <div class="ly-banner" style="top:190px;">
                <div class="pop-banner-slide">
                    <div class="swiper-container">
                        <div class="swiper-wrapper">
						<% 
							Do While Not oRs.eof
						%>
							<div class="swiper-slide">
							
								<div class="img">
									<img src="<%=oRs("MobileImage")%>" alt="<%=oRs("Title")%>" style="cursor:pointer;" onclick="APP_GoUrl('<%=oRs("LinkUrl")%>');">
								</div>
							
							</div>
						<%
								oRs.MoveNext
							Loop
						%>
                        </div>
                        <div class="swiper-pagination ty-red"></div>
                    </div>
                </div>

                <div class="close">
                    <span class="chk-today">
                        <input type="checkbox" id="chk-today-closed">
                        <label for="chk-today-closed">오늘 하루 보지 않기</label>
                    </span>
                    <button type="button" class="btn-hide-pop" onclick="close_slidePopup();">닫기</button>
                </div>
            </div>
        </div>

		<script type="text/javascript">
			var bannerSlide = new Swiper('.pop-banner-slide .swiper-container', {
				slidesPerView: 1,
				<% If TodaySlidePopupCount > 1 Then %>loop: true,<% End If %>
				loopedSlides: 10,
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
			function close_slidePopup() {
				if ($("input:checkbox[id='chk-today-closed']").prop("checked")) {
					var d = new Date();
					d.setDate(d.getDate() + 1);
					document.cookie = "SlidePopupToday=Y; path=/; expires=" + d.toGMTString() + ";";
				}

				$("#slidePopup").hide();
			}
		</script>
<%
	End If
	oRs.Close	
End If
%>