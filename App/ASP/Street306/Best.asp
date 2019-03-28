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
PageCode2 = "BP"
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
                    <img src="/images/img/@street_best_1.jpg" alt="Street306 BEST">
                    <p>BEST</p>
                </div>

                <div class="sale-item">
                    <div class="item-list">
                        <ul class="listview" id="ProductList">
							
                        </ul>
                    </div>
                </div>

                <div class="buttongroup" id="morebtn">
                    <button type="button" class="button is-expand" onclick="getNextPage('B');">
						<span class="icon is-right is-arrow-d2">더보기</span>
					</button>

                    <span class="pagination">
						<span class="current" id="nowpage"></span>/<span class="all" id="totalpage"></span>
                    </span>
                </div>

			</section>
        </div>
    </main>

	<form name="form" method="get">
		<input type="hidden" name="PCode" value="B" />
		<input type="hidden" name="Page" />
		<input type="hidden" name="PageSize" value="10" />
		<input type="hidden" name="ISTopN" />
	</form>


	<script type="text/javascript">
		function ProductList(pcode, page) {
			document.form.PCode.value = pcode;
			$("form[name='form'] > input[name='Page']").val(page);

			var arrHash = "";
			if (document.location.hash.replace("#", "") == "")						//해시값이 없을경우
			{
				$("form[name='form'] > input[name='ISTopN']").val("N");		// history.back() 으로 올경우 전체 페이지 다시가져오기
			}
			else																	//해시값이 있을경우
			{
				$("form[name='form'] > input[name='ISTopN']").val("Y");		// history.back() 으로 올경우 전체 페이지 다시가져오기
				arrHash = document.location.hash.replace("#", "").split("|");
			}

			$.ajax({
				type		 : "post",
				url			 : "/ASP/Street306/Ajax/ProductList.asp",
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

								// 목록 로딩시키기
								if (arrHash != "") {
									$("#ProductList").html(Data);
									$(window).scrollTop(arrHash[1]);
									document.location.hash = 0;
								} else {
									if (page == 1) {
										$("#ProductList").html(Data);
									} else {
										$("#ProductList").append(Data);
									}
								}
				},
				error		 : function (data) {
								openAlertLayer("alert", "처리중 오류가 발생하였습니다.", "closePop('alertPop', '');history.back();", "");
				}
			});
		}

		function getNextPage(pcode) {
			var page = document.form.Page.value;
			ProductList(pcode, parseInt(page) + 1);
		}

		function pushHash() {
			document.location.hash = $("form[name='form'] > input[name='Page']").val() + "|" + $(window).scrollTop();
		}
	</script>

	<script type="text/javascript">
		$(document).ready(function () {
			// history.back() 시 카테고리로 다시 페이지로딩
			if (document.location.hash) {
				var arrHash = document.location.hash.replace("#", "").split("|")
				if (arrHash.length == 2) {
					document.form.Page.value = arrHash[0];
					ProductList('B', arrHash[0]);
				} else {
					ProductList('B', 1);
				}
			} else {
				ProductList('B', 1);
			}
		});
	</script>


<!-- #include virtual="/INC/Footer.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>