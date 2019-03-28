<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'Index.asp - Street306
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
DIM oCmd						'# ADODB Command 개체

DIM i
DIM j
DIM x
DIM y

DIM wQuery						'# WHERE 절
DIM sQuery						'# SORT 절

DIM SubBanner	 : SubBanner	 = ""
DIM SubLinkUrl	 : SubLinkUrl	 = ""

DIM PCode

DIM ToDay : ToDay = R_YEAR & R_MONTH & R_DAY & R_HOUR & R_MIN
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성


wQuery = "WHERE BCode = 'S' "
sQuery = "ORDER BY IDX ASC "
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_Street306_Banner_Select_By_wQuery"

		.Parameters.Append .CreateParameter("@wQuery", adVarChar, adParamInput, 1000, wQuery)
		.Parameters.Append .CreateParameter("@sQuery", adVarChar, adParamInput,  100, sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		SubBanner	 = oRs("MobileImage")
		SubLinkUrl	 = oRs("LinkUrl")
END IF
oRs.Close
%>

<!-- #include virtual="/INC/Header.asp" -->
<!-- #include virtual="/INC/Top_Street306.asp" -->

    <!-- Main -->
    <main id="container" class="container">
        <div class="content">
            <section class="street-contents">
				<article class="main-style1">
<%
wQuery = "WHERE BCode = 'M' AND DelFlag = 'N' AND SDate <= '" & ToDay & "' AND EDate >= '" & ToDay & "' "
sQuery = "ORDER BY DisplayNum DESC "
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_Street306_Banner_Select_By_wQuery"
		.Parameters.Append .CreateParameter("@wQuery", adVarChar, adParamInput, 1000, wQuery)
		.Parameters.Append .CreateParameter("@sQuery", adVarChar, adParamInput, 100, sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

If Not oRs.EOF Then
%>
					<div class="main-slider">
						<div class="swiper-container main-swiper">
							<ul class="swiper-wrapper">
							<%
							Do While Not oRs.EOF
							%>
								<li class="swiper-slide">
									<a href="javascript:void(0)" onclick="LinkgoUrl('<%=oRs("LinkUrl")%>')" class="listitems">
										<div class="thumbnail">
											<img src="<%=oRs("MobileImage")%>" alt="Street306">
										</div>
									</a>
								</li>
							<%
								oRs.MoveNext
							Loop
							%>
							</ul>
							<div class="swiper-pagination"></div>
						</div>
					</div>
<%
End If
oRs.Close	
%>
				</article>

                <article class="street-main">
                    <img class="banner" src="<%=SubBanner%>" alt="Street306" <%IF SubLinkUrl <> "" THEN%>onclick="LinkgoUrl('<%=SubLinkUrl%>');" style="cursor:pointer;"<%END IF%>>
                </article>
                <article class="street-product">
                    <div class="contents-area">
                        <div id="tabs" class="tab" data-use="">
                            <ul class="tab-selector main-tab-btn">
                                <li class="part-2 active"><a href="javascript:ProductList('B', '1');;" data-target="tabs-col1">BEST SELLER</a></li>
                                <li class="part-2"><a href="javascript:ProductList('N', '1');;" data-target="tabs-col2">NEW ARRIVALS</a></li>
                            </ul>
                            <div id="tabs-col1" class="tab-panel active">
                            </div>
                            <div id="tabs-col2" class="tab-panel active">
                            </div>
                        </div>
                    </div>
                </article>

<%
wQuery = "WHERE DelFlag = 'N' AND BCode = 'B' "
sQuery = "ORDER BY DisplayNum ASC "
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_Street306_Banner_Select_By_wQuery"
		.Parameters.Append .CreateParameter("@wQuery", adVarChar, adParamInput, 1000, wQuery)
		.Parameters.Append .CreateParameter("@sQuery", adVarChar, adParamInput, 100, sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

If Not oRs.EOF Then
%>
                <article class="street-slide">
                    <div class="swiper-container street-swiper">
                        <ul class="swiper-wrapper">
						<%
						Do While Not oRs.EOF
						%>
                            <li class="swiper-slide"><a href="javascript:void(0)" onclick="APP_GoUrl('<%=oRs("LinkUrl")%>')"><img src="<%=oRs("MobileImage")%>" alt="Street306"></a></li>
						<%
							oRs.MoveNext
						Loop
						%>
                        </ul>
                        <div class="swiper-pagination"></div>
                    </div>
                </article>
<%
End If
oRs.Close



wQuery = "WHERE DelFlag = 'N' AND MFlag = 'Y' "
sQuery = "ORDER BY DisplayNum DESC "
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Street306_LookBook_Select_By_wQuery_For_Top"
		.Parameters.Append .CreateParameter("@Top", adVarChar, adParamInput, 10, "1000")
		.Parameters.Append .CreateParameter("@wQuery", adVarChar, adParamInput, 1000, wQuery)
		.Parameters.Append .CreateParameter("@sQuery", adVarChar, adParamInput, 100, sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

If Not oRs.EOF Then
%>
                <article class="lookbook wrap-item-list">
                    <p class="section-tit">LOOKBOOK</p>
                    <ul id="grid" class="looklist">
					<%
					Do While Not oRs.EOF
					%>
                        <li class="card">
                            <a href="javascript:void(0)" onclick="LookBookOpen('<%=oRs("IDX")%>');">
                                <img src="<%=oRs("PC_ListImage")%>" alt="<%=oRs("Title1")%>">
                                <div class="txt">
                                    <span><%=oRs("BrandName")%></span>
                                    <strong><%=oRs("Title1")%></strong>
                                    <p><%=oRs("Title2")%></p>
                                </div>
                            </a>
                        </li>
					<%
						oRs.MoveNext
					Loop	
					%>
                    </ul>
                </article>
<%
End If
oRs.Close	
%>
            </section>
        </div>
    </main>

	<form name="form" method="get">
		<input type="hidden" name="PCode" value="<%=PCode%>" />
		<input type="hidden" name="Page" />
		<input type="hidden" name="PageSize" value="4" />
		<input type="hidden" name="ISTopN" />
	</form>


	<script type="text/javascript">
		function ProductList(pcode, page) {
			document.form.PCode.value = pcode;
			document.form.Page.value = page;
			$.ajax({
				type		 : "post",
				url			 : "/ASP/Street306/Ajax/Main_ProductList.asp",
				async		 : false,
				data		 : $("form[name='form']").serialize(),
				dataType	 : "text",
				success		 : function (data) {
								arrData	 = data.split("|||||");
								Data	 = arrData[0];
								RecCnt	 = arrData[1];
								PageCnt	 = arrData[2];

								if (pcode == 'B') {
									$("#tabs-col1").html(Data);
								}

								if (pcode == 'N') {
									$("#tabs-col2").html(Data);
								}
				},
				error		 : function (data) {
								openAlertLayer("alert", "처리중 오류가 발생하였습니다.", "closePop('alertPop', '');", "");
				}
			});
		}

		function getNextPage(pcode) {
			var page = document.form.Page.value;
			ProductList(pcode, parseInt(page) + 1);
		}

		function getPrevPage(pcode) {
			var page = document.form.Page.value;
			ProductList(pcode, parseInt(page) - 1);
		}

		function LookBookOpen(idx) {
			$.ajax({
				type		 : "post",
				url			 : "/ASP/Street306/Ajax/LookBookView.asp",
				async		 : false,
				data		 : "IDX=" + idx,
				dataType	 : "text",
				success		 : function (data) {
								$("#LookBookContents").html(data);
								$("#lookbookView").show();
								$("#LookBookContents").scrollTop(0);
				},
				error		 : function (data) {
								openAlertLayer("alert", "처리중 오류가 발생하였습니다.", "closePop('alertPop', '');", "");
				}
			});
		}

		function LookBookClose() {
			$("#lookbookView").hide();
		}

		$(function () {
			var mainSwiper = new Swiper('.main-swiper', {
				slidesPerView: 1,
				loop: true,
				spaceBetween: 5,
				centeredSlides: true,
				autoplay: {
					delay: 2000,
				},
				autoplayDisableOnInteraction: true,
				observer: true,
				observeParents: true,
				pagination: {
					el: '.swiper-pagination',
					clickable: true
				},
			});

			var $man = $('#grid').masonry({
				initLayout: true,
				columnWidth: '.card',
				itemSelector: '.card',
				gutter: 5,
				horizontalOrder: true,
				percentPosition: true,
				transitionDuration: '0.5s'
			});

			setTimeout(function () {
				$man.masonry("layout");
			}, 600);

			ProductList('B', 1);
		});
	</script>


<!-- #include virtual="/INC/Footer.asp" -->

    <section class="wrap-pop" id="lookbookView">
        <div class="area-pop">
            <div class="top-exposed vertical">
                <!-- 더블 클래스 vertical로 팝업 호출/ 닫기 -->
                <div class="tit-pop">
                    <p class="tit">LOOKBOOK</p>
                    <button class="btn-hide-pop" type="button" onclick="LookBookClose();">닫기</button>
                </div>

                <div class="container-pop">
                    <div class="contents" id="LookBookContents">

                    </div>
                </div>
            </div>
        </div>
    </section>

<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>