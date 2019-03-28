<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'CouponApplyProductList.asp - 상품검색
'Date		: 2019.01.25
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
PageCode3 = "01"
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->
<!-- #include virtual="/Common/CheckID.asp" -->

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

DIM Page
DIM PageSize : PageSize = 20
DIM RecCnt
DIM PageCnt

DIM Idx

'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

Page			 = sqlFilter(Request("Page"))
Idx				 = sqlFilter(Request("Idx"))

IF Page			 = "" THEN Page		 = 1


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>


<!-- #include virtual="/INC/Header.asp" -->
<%
TopSubMenuTitle = "쿠폰적용 상품"
%>
<!-- #include virtual="/INC/TopSub.asp" -->


    <main id="container" class="container">
        <div class="content add">

            <section class="wrap-item-list">

                <div class="sale-item">
                    <div class="item-list">
                        <ul class="listview" id="CouponApplyProductList">
							
                        </ul>
                    </div>
                </div>

                <div class="buttongroup" id="morebtn">
                    <button type="button" class="button is-expand" onclick="getNextPage();">
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
		<input type="hidden" name="Page" />
		<input type="hidden" name="Idx" value="<%=Idx%>" />
		<input type="hidden" name="ISTopN" />
	</form>

	<script type="text/javascript">
		function get_CouponApplyProductList(page) {
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
				url: '/ASP/Mypage/Ajax/CouponApplyProductList.asp',
				data: $("form[name='form']").serialize(),
				async: false,
				type: 'get',
				dataType: 'html',
				success: function (data, textStatus, jqXHR) {
					if (data == "") {
						$("form[name='form'] > input[name='Page']").val(page - 1);
						return;
					}

					arrData = data.split("|||||");
					Data = arrData[0];
					RecCnt = arrData[1];
					PageCnt = arrData[2];

					$("#nowpage").html(page);
					$("#totalpage").html(PageCnt);

					$("#morebtn").show();
					if (parseInt(page) >= parseInt(PageCnt)) {
						$("#morebtn").hide();
					}

					// 목록 로딩시키기
					if (arrHash != "") {
						$("#CouponApplyProductList").html(Data);
						$(window).scrollTop(arrHash[1]);
						document.location.hash = 0;
					} else {
						if (page == 1) {
							$("#CouponApplyProductList").html(Data);
						} else {
							$("#CouponApplyProductList").append(Data);
						}
					}
				},
				error: function (data, textStatus, jqXHR) {
					//alert(jqXHR);
					alert(data.responseText);
					alert("상품 리스트를 가져오는 도중 오류가 발생하였습니다.");
				}
			});
		}

		function pushHash() {
			document.location.hash = $("form[name='form'] > input[name='Page']").val() + "|" + $(window).scrollTop();
		}

		function getNextPage() {
			var page = document.form.Page.value;
			get_CouponApplyProductList(parseInt(page) + 1);
		}

	</script>

	<script type="text/javascript">
		$(document).ready(function () {
			// history.back() 시 카테고리로 다시 페이지로딩
			if (document.location.hash) {
				var arrHash = document.location.hash.replace("#", "").split("|")
				if (arrHash.length == 2) {
					document.form.Page.value = arrHash[0];
					get_CouponApplyProductList(arrHash[0]);
				} else {
					get_CouponApplyProductList(1);
				}
			} else {
				get_CouponApplyProductList(1);
			}
		});
	</script>

<!-- #include virtual="/INC/Footer.asp" -->


	<script type="text/javascript">

	</script>

<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs1 = Nothing
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>
