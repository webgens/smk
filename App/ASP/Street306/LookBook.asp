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
PageCode1 = "ST"
PageCode2 = "LB"
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

Dim MainBanner
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>

<!-- #include virtual="/INC/Header.asp" -->
<!-- #include virtual="/INC/Top_Street306.asp" -->

    <!-- Main -->
    <main id="container" class="container">
        <div class="content">
            <section class="wrap-street">
                <div class="item-bg">
                    <img src="/images/img/@street_lookbook_1.jpg" alt="Street306 LOOKBOOK">
                    <p>LOOKBOOK</p>
                </div>

                <article class="lookbook wrap-item-list">

                    <ul id="grid" class="looklist">
                    </ul>

					<div class="buttongroup" id="morebtn">
						<button type="button" class="button is-expand" onclick="getNextLookBookList();">
							<span class="icon is-right is-arrow-d2">더보기</span>
						</button>

						<span class="pagination">
							<span class="current" id="nowpage"></span>/<span class="all" id="totalpage"></span>
						</span>
					</div>

                </article>

			</section>
        </div>
    </main>

	<form name="form" method="get">
		<input type="hidden" name="Page" />
		<input type="hidden" name="PageSize" value="10" />
		<input type="hidden" name="ISTopN" />
	</form>


	<script type="text/javascript">
		function LookBookList(page) {
			$("form[name='form'] > input[name='Page']").val(page);

			var $man = $('#grid').masonry({
				initLayout: true,
				columnWidth: '.card',
				itemSelector: '.card',
				gutter: 5,
				horizontalOrder: true,
				percentPosition: true,
				transitionDuration: '0.5s'
			});

			$.ajax({
				type		 : "post",
				url			 : "/ASP/Street306/Ajax/LookBookList.asp",
				async		 : false,
				data		 : $("form[name='form']").serialize(),
				dataType	 : "text",
				success		 : function (data) {
								if (data == "") {
									$("form[name='form'] > input[name='Page']").val(page - 1);
									return;
								}

								arrData	 = data.split("|||||");
								Data	 = arrData[0];
								RecCnt	 = arrData[1];
								PageCnt	 = arrData[2];

								$("#nowpage").html(page);
								$("#totalpage").html(PageCnt);

								$("#morebtn").show();
								if (parseInt(page) >= parseInt(PageCnt)) {
									$("#morebtn").hide();
								}

								if (page == 1) {
									var $items = $(Data);

									setTimeout(function() {
										$man.append($items).masonry("appended", $items, true).masonry("layout");
									}, 600);
						
								} else {
									var $items = $(Data);
									setTimeout(function () {
										$man.append($items).masonry("appended", $items, true).masonry("layout");
									}, 600);
								}
				},
				error		 : function (data) {
								openAlertLayer("alert", "처리중 오류가 발생하였습니다.", "closePop('alertPop', '');history.back();", "");
				}
			});
		}

		function getNextLookBookList() {
			var page = document.form.Page.value;
			LookBookList(parseInt(page) + 1);
		}

		function pushHash() {
			document.location.hash = $("form[name='form'] > input[name='Page']").val() + "|" + $(window).scrollTop();
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
	</script>

	<script type="text/javascript">
		$(document).ready(function () {

			LookBookList(1);
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