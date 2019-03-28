<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'Top100.asp - Top100 상품리스트
'Date		: 2019.01.06
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
PageCode1 = "T1"
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

DIM Page
DIM PageSize : PageSize = 20
DIM RecCnt
DIM PageCnt

DIM SCode1

Dim CategoryName

DIM ImageUrl
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

Page			 = sqlFilter(Request("Page"))
SCode1			 = sqlFilter(Request("SCode1"))

IF Page			 = "" THEN Page			 = 1

SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성

CategoryName = "전체"
IF SCode1 <> "" THEN
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Front_EShop_Category1_Select_By_CategoryCode1"

				.Parameters.Append .CreateParameter("@CategoryCode1", adChar, adParamInput, 2, SCode1)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		SET oCmd = Nothing

		IF NOT oRs.EOF THEN
				CategoryName = oRs("CategoryName1")
		ELSE
				oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
				Call AlertMessage("없는 분류 정보 입니다.", "/")
				Response.End
		END IF
		oRs.Close
END IF
%>

<!-- #include virtual="/INC/Header.asp" -->
<!-- #include virtual="/INC/TopMain.asp" -->

    <main id="container" class="container">
        <div class="content">
            <section class="wrap-item-list">
                <div class="item-bg">
                    <img src="/images/img/mo_top100.jpg" alt="TOP100">
                    <p>Top100</p>
                </div>

                <div class="sale-item">
                    <div class="item-list">
                        <ul class="listview" id="ProductList">
							
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
		<input type="hidden" name="SCode1" value="<%=SCode1%>" />
		<input type="hidden" name="Page" />
		<input type="hidden" name="ISTopN" />
	</form>

	<script type="text/javascript">
		function get_ProductList(page) {
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
				url: '/ASP/Product/Ajax/Top100.asp',
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
				error: function (data, textStatus, jqXHR) {
					//alert(jqXHR);
					//alert(data.responseText);
					alert("상품 리스트를 가져오는 도중 오류가 발생하였습니다.");
				}
			});
		}

		function pushHash() {
			document.location.hash = $("form[name='form'] > input[name='Page']").val() + "|" + $(window).scrollTop();
		}

		function getNextPage() {
			var page = document.form.Page.value;
			get_ProductList(parseInt(page) + 1);
		}

		//분류 레이어
		function CategoryLayerOpen() {
			$("#CateogryPop").show();
		}
	</script>

	<script type="text/javascript">
		$(document).ready(function () {
			// history.back() 시 카테고리로 다시 페이지로딩
			if (document.location.hash) {
				var arrHash = document.location.hash.replace("#", "").split("|")
				if (arrHash.length == 2) {
					document.form.Page.value = arrHash[0];
					get_ProductList(arrHash[0]);
				} else {
					get_ProductList(1);
				}
			} else {
				get_ProductList(1);
			}
		});
	</script>

<!-- #include virtual="/INC/Footer.asp" -->

    <section class="wrap-pop" id="CateogryPop">
        <div class="area-dim"></div>

        <div class="area-pop">
            <div class="alert">
                <div class="tit-pop">
                    <p class="tit">분류</p>
                    <button type="button" class="btn-hide-pop" onclick="$('#CateogryPop').hide();">닫기</button>
                </div>
                <div class="container-pop">
                    <div class="contents">
                        <div class="pop-category" id="subcategorylist">
							<button type="button" onclick="location.href='/ASP/Product/Top100.asp';" <% If SCode1 = "" Then %>class="on"<% End If %>><span>전체</span></button>
<%
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Category1_Select"
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
	Do While Not oRs.EOF
%>
							<button type="button" onclick="location.href='/ASP/Product/Top100.asp?SCode1=<%=oRs("CategoryCode1")%>';" <% If SCode1 = oRs("CategoryCode1") Then %>class="on"<% End If %>><span><%=oRs("CategoryName1")%></span></button>
<%
		oRs.MoveNext
	Loop
End If
oRs.Close
%>            
						</div>
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
