<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'Index.asp - 고객센터 메인
'Date		: 2018.12.28
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
PageCode2 = "01"
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


DIM sKeyword
DIM sClassName
DIM Idx
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


sKeyword	= sqlFilter(request("Keyword"))
sClassName	= sqlFilter(request("ClassName"))
Idx			= sqlFilter(request("Idx"))
	
SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>


<!-- #include virtual="/INC/Header.asp" -->
	<script type="text/javascript">
		function faqList(PageSize) {
			$.ajax({
				type: "post",
				url: "/ASP/Customer/Ajax/FaqList.asp?PageSize=" + PageSize,
				async: true,
				data: $("form[name=FaqForm]").serialize(),
				dataType: "text",
				success: function (data) {
					var splitData = data.split("|||||");
					var result = splitData[0];
					var cont = splitData[1];

					if (result == "OK") {
						$("#FaqList").html(cont);
						var RecCnt = parseInt($("#RecCnt").val());
						var PageSize = parseInt($("#PageSize").val());
						$("#TotalCount").html(" - 총 "+ RecCnt +"건");

						if (RecCnt > 5 && RecCnt > PageSize){
							$("#customer-btn-more > button").css("display","block");
						}else{
							$("#customer-btn-more > button").css("display","none");
						}
						return;
					}
				},
				error: function (data) {
					openAlertLayer("alert", "처리 도중 오류가 발생하였습니다.", "closePop('alertPop', '');", "");
				}
			});
		}

		function searchFaq(no, className){
			$("#TotalCount").html("");
			if (no == "99"){
				if(!$("form[name=FaqForm] input[name=Keyword]").val()){
					common_msgPopOpen("", "질문 키워드를 입력하여 주세요.", "", "Keyword");
					return;
				}
				$("form[name=FaqForm] input[name=ClassName]").val("");
				$("#ClassName").html("");
				$("form[name=FaqForm] input:radio[name=question_keyword]").prop("checked", false)
				faqList(5);
			}else{
				$("form[name=FaqForm] input[name=ClassName]").val(className);
				$("#ClassName").html(className);
				$("form[name=FaqForm] input[name=Keyword]").val("");
				faqList(5);
			}
		}
	</script>
<%TopSubMenuTitle = "FAQ"%>
<!-- #include virtual="/INC/TopSub.asp" -->


    <!-- Main -->
    <main id="container" class="container">
        <div class="sub_content">
            <div class="slider-for">
                <section>
                    <div class="FAQ customer">
                        <section class="FAQ-search">
                            <form name="FaqForm" id="FaqForm" method="post" action="javascript:searchFaq('99','');">
							<input type="hidden" name="ClassName" value="<%=sClassName%>" />
							<input type="hidden" name="Idx" value="<%=Idx%>" />
                            <div class="fieldset">
                                <label class="fieldset-label">질문검색</label>
                                <div class="fieldset-row">
                                    <span class="input is-expand">
                                        <input type="text" maxlength="20" name="Keyword" id="Keyword" value="<%=sKeyword%>" placeholder="질문 키워드를 입력하세요." onkeypress="if(event.keyCode == '13') { searchFaq('99',''); }">
                                    </span>
                                </div>
                            </div>
                            <button type="button" onclick="searchFaq('99','');" class="button-ty2 is-expand ty-bd-gray">검색</button>
                            <div class="customer-select">
                                <p>직접선택</p>
                                <div class="input-area">
                                    <input type="radio" id="FAQ-select1" name="question_keyword" <%IF sClassName="상품" THEN Response.Write "checked"%> onclick="searchFaq('3','상품')"><label for="FAQ-select1"><span>상품</span></label>
                                    <input type="radio" id="FAQ-select2" name="question_keyword" <%IF sClassName="배송" THEN Response.Write "checked"%> onclick="searchFaq('2','배송')" value="배송"><label for="FAQ-select2"><span>배송</span></label>
                                    <input type="radio" id="FAQ-select3" name="question_keyword" <%IF sClassName="주문/결제" THEN Response.Write "checked"%> onclick="searchFaq('5','주문/결제')" value="주문/결제"><label for="FAQ-select3"><span>주문/결제</span></label>
                                    <input type="radio" id="FAQ-select4" name="question_keyword" <%IF sClassName="취소/환불" THEN Response.Write "checked"%> onclick="searchFaq('6','취소/환불')" value="취소/환불"><label for="FAQ-select4"><span>취소/환불</span></label>
                                    <input type="radio" id="FAQ-select5" name="question_keyword" <%IF sClassName="교환/반품" THEN Response.Write "checked"%> onclick="searchFaq('1','교환/반품')" value="교환/반품"><label for="FAQ-select5"><span>교환/반품</span></label>
                                    <input type="radio" id="FAQ-select6" name="question_keyword" <%IF sClassName="AS" THEN Response.Write "checked"%> onclick="searchFaq('9','AS')" value="AS관련"><label for="FAQ-select6"><span>AS관련</span></label>
                                    <input type="radio" id="FAQ-select7" name="question_keyword" <%IF sClassName="쿠폰/적립금" THEN Response.Write "checked"%> onclick="searchFaq('7','쿠폰/적립금')" value="쿠폰/적립금"><label for="FAQ-select7"><span>쿠폰/적립금</span></label>
                                    <input type="radio" id="FAQ-select8" name="question_keyword" <%IF sClassName="슈즈상품권/금액할인권" THEN Response.Write "checked"%> onclick="searchFaq('4','슈즈상품권/금액할인권')" value="슈즈상품권"><label for="FAQ-select8"><span>슈즈상품권</span></label>
                                    <input type="radio" id="FAQ-select9" name="question_keyword" <%IF sClassName="회원/기타" THEN Response.Write "checked"%> onclick="searchFaq('8','회원/기타')" value="회원/기타"><label for="FAQ-select9"><span>회원/기타</span></label>
                                </div>
                            </div>
							</form>
                        </section>
                        <section class="customer-Q">
                            <div class="h-line">
                                <h2 class="h-level4" id="ClassName"><%=sClassName%></h2> <span id="TotalCount"></span>
                            </div>
                            <div class="ly-accord-sub" id="FaqList">
                            </div>
                        </section>
                        <div class="customer-btn-more" id="customer-btn-more">
                            <button type="button" onclick="faqList(parseInt($('#PageSize').val())+5);" class="button-ty2 is-expand ty-bd-gray" style="display:none;">더보기</button>
                        </div>
                    </div>
                </section>
            </div>
        </div>
    </main>

		<script>
			faqList(5);
		</script>
		

<!-- #include virtual="/INC/FooterNoBNB.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>

