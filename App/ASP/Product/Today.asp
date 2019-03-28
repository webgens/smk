<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'Today.asp - 투데이딜
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
PageCode1 = "TD"
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

Dim NextDate 
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

NextDate = DateADD("d", 1, Date)

SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성

SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Mobile_EShop_Event_Category_Product_Select_By_TodayDeal"
		.Parameters.Append .CreateParameter("@Today",		 adVarchar, adParaminput, 12	, R_YEAR & R_MONTH & R_DAY & R_HOUR & R_MIN)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing
%>

<!-- #include virtual="/INC/Header.asp" -->
<!-- #include virtual="/INC/TopMain.asp" -->

    <main id="container" class="container">
        <div class="content">
            <section class="wrap-item-list">
                <div class="today-time">
                    <strong>TODAY'S DEAL</strong>
                    <div class="text">
                        <p>오늘의딜</p>
                        <p>남은시간</p>
                    </div>
                    <div id="remaintime"></div>
                </div>
<%
If Not oRs.EOF Then
%>
                <div class="item-list">
                    <div class="today-item">
                        <ul class="listview">
						<%
							Do While Not oRs.EOF
						%>
                            <li>
                                <div class="salebadge"><%=FormatNumber(oRs("DiscountRate"), 0)%>%</div>
                                <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>')" class="listitems">
                                    <div class="thumbnail"><img src="<%=oRs("ImageUrl")%>" alt="<%=oRs("ProductName")%>"></div>
                                    <div class="today-txt">
                                        <p class="brand-name"><%=oRs("BrandName")%></p>
                                        <div class="iteminfo">
                                            <h1 class="product-name pname"><%=oRs("ProductName")%></h1>
                                            <p class="price"><strong><%=FormatNumber(oRs("SalePrice"), 0)%></strong>원</p>
                                        </div>
								</a>
								<a nohref class="listitems">
                                        <p class="optional-info">
											<button type="button" class="btn-size" onclick="SizeLayerOpen('<%=oRs("ProductCode")%>');">SIZE</button>
											<span class="icon ico-fav"><%=FormatNumber(oRs("WishCnt"), 0)%></span>
											<span class="icon ico-cmt"><%=FormatNumber(oRs("ReviewCnt"), 0)%></span>
                                        </p>
                                    </div>
                                </a>
                            </li>
						<%
								oRs.MoveNext
							Loop	
						%>
						</ul>
					</div>
				</div>
<%
End If
oRs.Close

SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Mobile_EShop_Event_Category_Product_Select_By_TodayDeal_NextDay"
		.Parameters.Append .CreateParameter("@Today",		 adVarchar, adParaminput, 12	, Replace(NextDate, "-", "") & "2400")
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

If Not oRs.EOF Then
%>
                    <div class="tomorrow-tit">
                        <p class="head">SEE U TOMORROW!</p>
                        <p class="inform">내일 딜 예정상품 입니다.</p>
                    </div>
                    <div class="tomorrow-preview">
                        <div class="item-list">
                            <div class="today-item">
                                <ul class="listview">
								<%
									i = 1
									Do While Not oRs.EOF
								%>

                                    <li>
                                        <a href="/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>" class="listitems">
                                            <div class="thumbnail"><img src="<%=oRs("ImageUrl")%>" alt="<%=oRs("ProductName")%>"></div>
                                            <div class="today-txt">
                                                <p class="brand-name"><%=oRs("BrandName")%></p>
                                                <div class="iteminfo">
                                                    <h1 class="product-name pname"><%=oRs("ProductName")%></h1>
                                                    <p class="price"><strong><%=FormatNumber(oRs("Next_SalePrice"), 0)%></strong>원</p>
                                                </div>
										</a>
										<a nohref class="listitems">
                                                <p class="optional-info">
                                                    <button type="button" class="btn-size" onclick="SizeLayerOpen('<%=oRs("ProductCode")%>');">SIZE</button>
                                                    <span class="icon ico-fav"><%=FormatNumber(oRs("WishCnt"), 0)%></span>
                                                    <span class="icon ico-cmt"><%=FormatNumber(oRs("ReviewCnt"), 0)%></span>
                                                </p>
                                            </div>
                                        </a>
                                        <div class="tomorrow-dim">
                                            <p class="txt">내일 만나요!</p>
                                        </div>
                                    </li>
								<%
										i = i + 1
										If i > 3 Then Exit Do
										oRs.MoveNext
									Loop	
								%>

                                </ul>
                            </div>
                        </div>
                    </div>
<%
End If
oRs.Close	
%>
            </section>
    
	   </div>

    </main>

	<script type="text/javascript">
		//timedeal
		function timedealNew() {
			var now = new Date();
			//var dday = new Date('<%=NextDate & "T00:00:00"%>');
			var dday = new Date('<%=NextDate & "T00:00:00"%>' + 'Z');
			dday = dday.getTime() + dday.getTimezoneOffset() * 60 * 1000


			var days = (dday - now) / 1000 / 60 / 60 / 24;
			var daysRound = Math.floor(days);
			var hours = (dday - now) / 1000 / 60 / 60 - (24 * daysRound);
			var hoursRound = Math.floor(hours);
			var minutes = (dday - now) / 1000 / 60 - (24 * 60 * daysRound) - (60 * hoursRound);
			var minutesRound = Math.floor(minutes);
			var seconds = (dday - now) / 1000 - (24 * 60 * 60 * daysRound) - (60 * 60 * hoursRound) - (60 * minutesRound);
			var secondsRound = Math.floor(seconds);
			var miliseconds = ((dday - now) - (24 * 60 * 60 * 1000 * daysRound) - (60 * 60 * 1000 * hoursRound) - (60 * 1000 * minutesRound) - (1000 * secondsRound)) / 10;
			var milisecondsRound = Math.round(miliseconds);

			if (minutesRound < 10) minutesRound = '0' + minutesRound;
			if (secondsRound < 10) secondsRound = '0' + secondsRound;
			if (milisecondsRound < 10) milisecondsRound = '0' + milisecondsRound;

			var todayresult = hoursRound + ':' + minutesRound + ':' + secondsRound + ':' + milisecondsRound;
			
			document.getElementById('remaintime').innerHTML = todayresult;
		}
	</script>

	<script type="text/javascript">
		$(document).ready(function () {
			setInterval(timedealNew, 1);
		});
	</script>
<!-- #include virtual="/INC/Footer.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>
