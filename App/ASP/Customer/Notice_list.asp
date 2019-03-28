<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'Notice_List.asp - 고객센터 > 공지사항 리스트
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
PageCode1 = "06"
PageCode2 = "06"
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
		<script>
			function noticeList(PageSize){
				$.ajax({
					type: "post",
					url: "/ASP/Customer/Ajax/NoticeList.asp",
					async: true,
					data: "PageSize="+PageSize,
					dataType: "text",
					success: function (data) {
						var splitData = data.split("|||||");
						var result = splitData[0];
						var cont = splitData[1];

						if (result == "OK") {
							$("#NoticeList").html(cont);

							var RecCnt = parseInt($("#RecCnt").val());
							var PageSize = parseInt($("#PageSize").val());

							if (RecCnt > 5 && RecCnt > PageSize){
								$("#customer-btn-more").css("display","block");
							}else{
								$("#customer-btn-more").css("display","none");
							}
							return;
						}
					},
					error: function (data) {
						//alert(data.responseText);
						common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.");
					}
				});
				
			}

			function noticeView(idx){
				$("form[name=NoticeListForm] input[name=Idx]").val(idx);
				common_PopOpen('DimDepth1', 'NoticeView');
			}
		</script>

<%TopSubMenuTitle = "슈마커소식"%>
<!-- #include virtual="/INC/TopSub.asp" -->


    <!-- Main -->
    <main id="container" class="container">
        <div class="sub_content">
            <div class="slider-for">
                <section>
                    <div class="customer-news">
						<div class="customer-form">
                            <div class="h-line">
                                <h2 class="h-level4">슈마커 소식</h2>
                                <p>슈마커의 새로운 소식을 전해드립니다.</p>
                            </div>
                        </div>

                        <section class="news-list" id="NoticeList">
                        </section>

                        <div class="customer-btn-more" id="customer-btn-more">
                            <button type="button" onclick="noticeList(parseInt($('#PageSize').val())+10);" class="button-ty2 is-expand ty-bd-gray">더보기</button>
                        </div>
                    </div>

                </section>

            </div>
        </div>
    </main>


	<script>
		noticeList(10);
	</script>
		

<!-- #include virtual="/INC/FooterNoBNB.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>

