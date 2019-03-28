<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'OrderCRXList.asp - 마이페이지 > 주문 취소/교환/반품 조회
'Date		: 2019.01.03
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
PageCode2 = "02"
PageCode3 = "00"
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->

<%
IF U_NUM = "" AND N_NAME = "" THEN
		Response.Redirect "/ASP/Member/SubLogin.asp?ProgID=" & Server.URLEncode(ProgID)
		Response.End
END IF

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

DIM SDate
DIM EDate
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

SDate			 = sqlFilter(Request("SDate"))
EDate			 = sqlFilter(Request("EDate"))


IF SDate	= "" THEN SDate		= DateAdd("m", -1, Date)
IF EDate	= "" THEN EDate		= Date

	
SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>


<!-- #include virtual="/INC/Header.asp" -->
	<style type="text/css">
		.informItem .cont .oneplusone { position: absolute; right: 0; top: 0; font-size: 9px; color: #e62019; display: block; width: 84px; line-height: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; text-align: right; }
		.detailView .select { width: 100%; }
		.delivery-price { margin-bottom: 7px; line-height: 18px; border-bottom: 1px solid #e1e1e1; }
		.delivery-price .list { padding: 15px 12px 12px; }
		.delivery-price .list:first-child { border-top: 1px solid #e1e1e1; }
		.delivery-price .tit { margin-bottom: 3px; }
		.delivery-price .tit>span { vertical-align: top; }
		.delivery-price .tit>span:nth-child(2):before { content: ''; display: inline-block; width: 1px; height: 8px; margin-left: 8px; margin-right: 8px; margin-top: -2px; background-color: #e1e1e1; vertical-align: middle; }
		.delivery-price .tit .ty-red { font-size: 14px; color: #e62019; margin-left: 5px; }
		.addr-list .tit>span:nth-child(3):before { content: ''; display: inline-block; width: 1px; height: 8px; margin-left: 8px; margin-right: 8px; margin-top: -2px; background-color: #e1e1e1; vertical-align: middle; }
		.area-radio .rad-ty1:nth-child(3n+1) { border-left: 1px solid #c8c8c8; }
	</style>
	<%IF Request.ServerVariables("HTTPS") = "on" Then%>
	<script src="//ssl.daumcdn.net/dmaps/map_js_init/postcode.v2.js"></script>
	<%ELSE%>
	<script src="//dmaps.daum.net/map_js_init/postcode.v2.js"></script>
	<%END IF%>
	<script type="text/javascript">
		function execDaumPostcode(zipCode, addr1, addr2) {
			new daum.Postcode({
				oncomplete: function(data) {
					// 팝업에서 검색결과 항목을 클릭했을때 실행할 코드를 작성하는 부분.
					//alert(data.addressType + "-" + data.userSelectedType + "\n\n1 : " + data.roadAddress + "\n2 : " + data.autoRoadAddress + "\n\n3 : " + data.jibunAddress + "\n4 : " + data.autoJibunAddress);

					// data.addressType : 주소검색방법 R=도로명검색, J=지번검색
					// data.userSelectedType : 주소선택구분 R=도로명주소선택, J=지번주소선택
					var roadAddress = data.roadAddress;
					var jibunAddress = data.jibunAddress;
					if (data.addressType == "R") {
						if (jibunAddress == "" && data.userSelectedType == "R") {
							jibunAddress = data.autoJibunAddress;
						}
					} else {
						if (roadAddress == "" && data.userSelectedType == "J") {
							roadAddress = data.autoRoadAddress;
						}
					}

					// 도로명 주소의 노출 규칙에 따라 주소를 조합한다.
					// 내려오는 변수가 값이 없는 경우엔 공백('')값을 가지므로, 이를 참고하여 분기 한다.
					var fullRoadAddr = roadAddress; // 도로명 주소 변수
					var extraRoadAddr = ''; // 도로명 조합형 주소 변수

					// 법정동명이 있을 경우 추가한다. (법정리는 제외)
					// 법정동의 경우 마지막 문자가 "동/로/가"로 끝난다.
					if(data.bname !== '' && /[동|로|가]$/g.test(data.bname)){
						extraRoadAddr += data.bname;
					}
					// 건물명이 있고, 공동주택일 경우 추가한다.
					if(data.buildingName !== '' && data.apartment === 'Y'){
					   extraRoadAddr += (extraRoadAddr !== '' ? ', ' + data.buildingName : data.buildingName);
					}
					// 도로명, 지번 조합형 주소가 있을 경우, 괄호까지 추가한 최종 문자열을 만든다.
					if(extraRoadAddr !== ''){
						extraRoadAddr = ' (' + extraRoadAddr + ')';
					}
					// 도로명, 지번 주소의 유무에 따라 해당 조합형 주소를 추가한다.
					if(fullRoadAddr !== ''){
						fullRoadAddr += extraRoadAddr;
					}

					// 우편번호와 주소 정보를 해당 필드에 넣는다.
					document.getElementById(zipCode).value = data.zonecode; //5자리 새우편번호 사용
					document.getElementById(addr1).value = fullRoadAddr;
					document.getElementById(addr2).focus();
					//document.getElementById('sample4_jibunAddress').value = jibunAddress;

					// iframe을 넣은 element를 안보이게 한다.
					// (autoClose:false 기능을 이용한다면, 아래 코드를 제거해야 화면에서 사라지지 않는다.)
					//document.getElementById('post_layer').style.display = 'none';
					closePop('PopupPostSearch');
				},
				width: '100%',
				height: '100%',
				maxSuggestItems: 5
			}).embed(document.getElementById('PopupPostContents'));

			openPop('PopupPostSearch');
		}
	</script>
	<script type="text/javascript" src ="/ASP/Mypage/JS/OrderCRX.js?ver=<%=U_DATE & U_TIME%>"></script>

<%TopSubMenuTitle = "쇼핑내역"%>
<!-- #include virtual="/INC/TopSub.asp" -->


    <!-- Main -->
    <main id="container" class="container">
        <div class="sub_content">

            <div class="wrap-mypage">
				<div style="height:8px"></div>
                <div class="shopping-list">
                    <section>
                        <div id="MypageSubMenu" class="ly-title accordion">
                            <div class="selector">
	                            <button type="button" class="btn-list clickEvt" data-target="MypageSubMenu">주문취소/반품/교환 내역</button>
							</div>
							<div class="option my-recode">
								<!-- #include virtual="/ASP/Mypage/SubMenu_Order.asp" -->
							</div>
                        </div>
                        <div>
                            <div id="tabs">
                                <div class="tab-mypage">

									<form name="form" id="form">
										<input type="hidden" name="Page"		id="Page"			value="1"	/>
										<input type="hidden" name="SCancelType" id="SCancelType"	value=""	/>
										<input type="hidden" name="SOPIdx"		id="SOPIdx"			value=""	/>

                                    <div>
                                        <div class="ly-calendar">
                                            <div class="tit">
                                                <span>시작일</span>
                                                <span>종료일</span>
                                            </div>
                                            <div class="wrap">
                                                <div class="date-picker">
                                                    <input type="text" name="SDate" id="SDate" value="<%=SDate%>" class="date-from" readonly="readonly" />
                                                </div>
                                                <div class="date-picker">
                                                    <input type="text" name="EDate" id="EDate" value="<%=EDate%>" class="date-to" readonly="readonly" />
                                                </div>
                                            </div>
                                            <div class="area-radio">
                                                <span class="rad-ty1">
													<input type="radio" id="oneMonth" name="period_1" onclick="setDate('1m', 'SDate', 'EDate')" checked />
													<label for="oneMonth">1개월</label>
												</span>
                                                <span class="rad-ty1">
													<input type="radio" id="threeMonth" name="period_1" onclick="setDate('3m', 'SDate', 'EDate')" />
													<label for="threeMonth">3개월</label>
												</span>
                                                <span class="rad-ty1">
													<input type="radio" id="sixMonth" name="period_1" onclick="setDate('6m', 'SDate', 'EDate')" />
													<label for="sixMonth">6개월</label>
												</span>
                                            </div>

                                            <button type="button" onclick="getOrderList(1, '', '')" class="button-ty2 is-expand ty-bd-gray">조회</button>
                                        </div>
									</div>

									</form>

                                    <div id="OrderList">
									</div>
                                </div>
                            </div>
                        </div>
                    </section>
                </div>
            </div>
        </div>
    </main>



	<script type="text/javascript">
		$(function () {
			getOrderList(1, "", "");
		});
	</script>



<!-- #include virtual="/INC/FooterNoBNB.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>