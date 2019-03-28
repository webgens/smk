<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'OrderList.asp - 마이페이지 > 쇼핑내역 > 주문/배송 조회
'Date		: 2018.12.31
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
PageCode2 = "01"
PageCode3 = "03"
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->

<%
IF U_NUM = "" AND N_NAME = "" THEN
		Response.Redirect "/ASP/Mypage/Login.asp?ProgID=" & Server.URLEncode(ProgID)
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
	<script type="text/javascript" src ="/ASP/Mypage/JS/Order.js?ver=<%=U_DATE & U_TIME%>"></script>

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
	                            <button type="button" class="btn-list clickEvt" data-target="MypageSubMenu">A/S 신청내역</button>
							</div>
							<div class="option my-recode">
								<!-- #include virtual="/ASP/Mypage/SubMenu_Order.asp" -->
							</div>
                        </div>
                        <div>
                            <div id="tabs">
                                <div class="tab-mypage">
                                    
                                    <div>
										<form name="form" id="form">
										<input type="hidden" name="Page"		id="Page"			value="1"	/>
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
													<input type="radio" id="oneMonth" name="period_1" onclick="setDate('1m', 'SDate', 'EDate')" checked>
													<label for="oneMonth">1개월</label>
												</span>
                                                <span class="rad-ty1">
													<input type="radio" id="threeMonth" name="period_1" onclick="setDate('3m', 'SDate', 'EDate')">
													<label for="threeMonth">3개월</label>
												</span>
                                                <span class="rad-ty1">
													<input type="radio" id="sixMonth" name="period_1" onclick="setDate('6m', 'SDate', 'EDate')">
													<label for="sixMonth">6개월</label>
												</span>
                                            </div>

                                            <button type="button" onclick="getOrderAsList(1, '', '')" class="button-ty2 is-expand ty-bd-gray">조회</button>
                                        </div>
										</form>

                                        <div class="h-line">
                                            <h2 class="h-level4">A/S 신청내역</h2>
                                        </div>

                                        <ul class="informView" id="OrderAsList">
                                        </ul>
                                    </div>
                                    <!-- // A/S 신청하기 -->
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
			getOrderAsList(1, "", "");
		});

		function getOrderAsList() {

			$.ajax({
				type		 : "post",
				url			 : "/ASP/Mypage/Ajax/OrderAsList.asp",
				async		 : false,
				data		 : $("#form").serialize(),
				dataType	 : "text",
				success		 : function (data) {
								var splitData	 = data.split("|||||");
								var result		 = splitData[0];
								var cont		 = splitData[1];


								if (result == "OK") {
									$("#OrderAsList").html(cont);
									return;
								}
								else {
									openAlertLayer("alert", cont, "closePop('alertPop', '')", "");
									return;
								}
				},
				error		 : function (data) {
								alert(data.responseText);
								openAlertLayer("alert", "처리 도중 오류가 발생하였습니다.", "closePop('alertPop', '')", "");
				}
			});

		}
	</script>



<!-- #include virtual="/INC/Footer.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>