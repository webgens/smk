<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'Store.asp - 고객센터 > 전국매장안내
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
PageCode1 = "06"
PageCode2 = "03"
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


'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>


<!-- #include virtual="/INC/Header.asp" -->
	<script type="text/javascript" src="//dapi.kakao.com/v2/maps/sdk.js?appkey=<%=KAKAO_LOGIN_CLIENTID%>&libraries=services"></script>
	<script type="text/javascript">
		function storeList(PageSize) {
			$.ajax({
				type		 : "post",
				url			 : "/ASP/Customer/Ajax/StoreList.asp?PageSize=" + PageSize,
				async		 : true,
				data		 : $("form[name=StoreForm]").serialize(),
				dataType	 : "text",
				success		 : function (data) {
								var splitData	 = data.split("|||||");
								var result		 = splitData[0];
								var cont		 = splitData[1];

								if (result == "OK") {
									$("#StoreList").html(cont);
					
									var RecCnt = parseInt($("#RecCnt").val());
									var PageSize = parseInt($("#PageSize").val());
									$("#TotalCount").html(" - 총 "+ RecCnt +"건");

									if (RecCnt > 5 && RecCnt > PageSize) {
										$("#customer-btn-more > button").css("display","block");
									}
									else {
										$("#customer-btn-more > button").css("display","none");
									}
									return;
								}
				},
				error		 : function (data) {
								openAlertLayer("alert", "매장 정보를 가져오는 도중 오류가 발생하였습니다.", "closePop('alertPop', '');", "");
				}
			});

		}

		function searchStore(getValue){
			$("form[name=StoreForm] input[name=ChannelNM]").val(getValue);
			storeList(5);
		}

		function storeView(xPoint, yPoint, shopNM, addr, tel) {
			$("form[name=StoreView] input[name=ShopNM]").val(shopNM);
			$("form[name=StoreView] input[name=ADDR]").val(addr);
			$("form[name=StoreView] input[name=TEL]").val(tel);
			$("form[name=StoreView] input[name=XPoint]").val(xPoint);
			$("form[name=StoreView] input[name=YPoint]").val(yPoint);

			if (shopNM != "" && addr != "") {
				APP_PopupGoUrl("/ASP/Customer/StoreView.asp?" + $("#StoreView").serialize());
			}	
		}
	</script>

<%TopSubMenuTitle = "전국 매장 안내"%>
<!-- #include virtual="/INC/TopSub.asp" -->


    <!-- Main -->
    <main id="container" class="container">
        <div class="sub_content">
            <div class="slider-for">
                <section>
                    <div class="customer customer-store" id="StoreList">
                    </div>
                </section>
            </div>
        </div>
    </main>


	<script type="text/javascript">
		$(function () {
			storeList(5);
		})
	</script>

<!-- #include virtual="/INC/FooterNoBNB.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>