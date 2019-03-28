<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'index.asp - 고객센터 메인
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


'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>


<!-- #include virtual="/INC/Header.asp" -->
	<script type="text/javascript">
		function searchFaq(no, className) {
			if (no == "99"){
				if (!$("form[name=FaqForm] input[name=Keyword]").val()) {
					openAlertLayer("alert", "질문 키워드를 입력하여 주십시오.", "closePop('alertPop', 'Keyword');", "");
					return;
				}
				$("form[name=FaqForm] input[name=ClassName]").val("");
			}
			else {
				$("form[name=FaqForm] input[name=ClassName]").val(className);
			}

			APP_GoUrl("/ASP/Customer/Faq_List.asp?" + $("#FaqForm").serialize());
			//FaqForm.action = "/ASP/Customer/Faq_List.asp";
			//FaqForm.submit();
		}

		function mainFaqTop5List(){
			$.ajax({
				type: "post",
				url: "/ASP/Customer/Ajax/MainFaqTop5List.asp",
				async: true,
				data: "",
				dataType: "text",
				success: function (data) {
					var splitData = data.split("|||||");
					var result = splitData[0];
					var cont = splitData[1];

					if (result == "OK") {
						$("#MainFaqTop5List").html(cont);
						return;
					}
				},
				error: function (data) {
					//alert(data.responseText);
					common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.");
				}
			});
		}

	</script>
<%TopSubMenuTitle = "고객센터"%>
<!-- #include virtual="/INC/TopCustomer.asp" -->


    <!-- Main -->
    <main id="container" class="container">
        <div class="content">
            <div class="slider-for">
                <section>
                    <div class="customer">
                        <section class="center">
                            <div class="tel">
                                <p>슈마커 고객센터</p>
                                <strong>080-030-2809</strong>
                                <span>평일 10:00~17:00 (점심 12:00~13:00)</span>
                            </div>
							<!--
                            <div class="btn">
                                <a href="javascript:common_PopOpen('DimDepth2','MyInfoModify');" class="btn1">실시간 상담하기</a>
                                <a href="javascript:popMtmQnaAdd();" class="btn2">1:1 문의하기</a>
                            </div>
							-->
                        </section>
						<!--
                        <section class="go">
                            <a href="/ASP/Mypage/OrderList.asp"><span>주문/배송조회</span></a>
                            <a href="/ASP/Mypage/MyReview.asp"><span>내 상품 후기</span></a>
                            <a href="/ASP/Mypage/MyPickList.asp"><span>찜한 상품</span></a>
                            <a href="/ASP/Mypage/MyReentry.asp"><span>재 입고 알림 신청</span></a>
                            <a href="/ASP/Mypage/AddressList.asp"><span>배송지관리</span></a>
                            <a href="#"><span>참여이벤트</span></a>
                        </section>
						-->
                        <div class="inquire">
                            <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Customer/Store.asp')" class="">
                                <p><span>전국<br>매장안내</span></p>
                            </a>
                            <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Customer/PartnerShip.asp')" class="">
                                <p><span>입점/제휴<br>문의</span></p>
                            </a>
                            <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Customer/GroupPurchase.asp')" class="">
                                <p><span>단체구매<br>문의</span></p>
                            </a>
                        </div>
                        <section class="FAQ-search">
							<form name="FaqForm" id="FaqForm" method="post" action="javascript:searchFaq('99','')">
							<input type="hidden" name="ClassName" value="" />
                            <div class="fieldset">
                                <label class="fieldset-label black">FAQ 검색</label>
                                <div class="fieldset-row">
                                    <span class="input is-expand">
                                        <input type="text" maxlength="20" name="Keyword" id="Keyword" placeholder="질문 키워드를 입력하세요.">
                                    </span>
                                </div>
                            </div>
                            <button type="button" onclick="searchFaq('99','')" class="button-ty2 is-expand ty-bd-gray">검색</button>
							</form>
                            <div class="customer-select">
                                <p>직접선택</p>
                                <div class="input-area">
                                    <input type="radio" id="FAQ-select1" name="ClassName" value="상품" onclick="searchFaq('3','상품');"><label for="FAQ-select1"><span>1개월</span></label>
                                    <input type="radio" id="FAQ-select2" name="ClassName" value="배송" onclick="searchFaq('2','배송');"><label for="FAQ-select2"><span>배송</span></label>
                                    <input type="radio" id="FAQ-select3" name="ClassName" value="주문/결제" onclick="searchFaq('5','주문/결제');"><label for="FAQ-select3"><span>주문/결제</span></label>
                                    <input type="radio" id="FAQ-select4" name="ClassName" value="취소/환불" onclick="searchFaq('6','취소/환불');"><label for="FAQ-select4"><span>취소/환불</span></label>
                                    <input type="radio" id="FAQ-select5" name="ClassName" value="교환/반품" onclick="searchFaq('1','교환/반품');"><label for="FAQ-select5"><span>교환/반품</span></label>
                                    <input type="radio" id="FAQ-select6" name="ClassName" value="AS관련" onclick="searchFaq('9','AS');"><label for="FAQ-select6"><span>AS관련</span></label>
                                    <input type="radio" id="FAQ-select7" name="ClassName" value="쿠폰/적립금" onclick="searchFaq('7','쿠폰/적립금');"><label for="FAQ-select7"><span>쿠폰/적립금</span></label>
                                    <input type="radio" id="FAQ-select8" name="ClassName" value="슈즈상품권" onclick="searchFaq('4','슈즈상품권/금액할인권');"><label for="FAQ-select8"><span>슈즈상품권</span></label>
                                    <input type="radio" id="FAQ-select9" name="ClassName" value="회원/기타" onclick="searchFaq('8','회원/기타');"><label for="FAQ-select9"><span>회원/기타</span></label>
                                </div>
                            </div>
                        </section>
                        <section class="customer-Q">
                            <div class="h-line">
                                <h2 class="h-level4">자주 묻는 질문 TOP5</h2>
                                <button type="button" onclick="APP_GoUrl('/ASP/Customer/Faq_List.asp')" class="more">더보기</button>
                            </div>
                            <div class="ly-accord-sub" id="MainFaqTop5List">
                            </div>
							<div style="height:20px;"></div>
                        </section>
                    </div>
                </section>
            </div>

        </div>
    </main>


<!-- #include virtual="/INC/FooterNoBNB.asp" -->
    <script>
		mainFaqTop5List();
    </script>
<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>
