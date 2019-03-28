<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'MainFaqTop10List.asp - 고객센터 메인 > 자주 묻는 질문 TOP 5
 'Date		: 2019.01.06
'Update	: 
'/****************************************************************************************'

'//페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'//---------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->

<%
'/****************************************************************************************'
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

DIM IDX
DIM Title
DIM Contents
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성

Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Mobile_EShop_FAQ_Select_TOP5"

'		.Parameters.Append .CreateParameter("@PAGE",		adInteger,	adParamInput,	  ,		Page)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing

Response.Write "OK|||||"
IF oRs.EOF THEN
%>
                                <div id="qna1" class="qna">
                                    <div class="accord-sub-mypage">
                                        <div class="ly-title_sub">
                                            <button type="button">등록된 FAQ가 없습니다.</button>
                                        </div>
                                    </div>
                                </div>
<%
ELSE
	j = 1
	DO UNTIL oRs.EOF
		IDX				= oRs("IDX")
		Title			= oRs("Title")
		Contents		= oRs("Contents")
%>
                                <div id="qna<%=j+1%>" class="qna">
                                    <div class="accord-sub-mypage">
                                        <div class="ly-title_sub">
                                            <button type="button" class="clickEvt_sub btn-list" data-target="qna<%=j+1%>"><%=Title%></button>
                                        </div>
                                        <div class="ly-content_sub">
                                            <p><%=Contents%></p>
                                        </div>
                                    </div>
                                </div>
<%
	j = j + 1
	oRs.MoveNext
	LOOP

END IF
%>
								<!-- //script 공통 -->
								<script>
									// 아코디언
									var mypageAccodion = function(){
										var selector,
											module;

										selector = {
											parent : '.accord-mypage',
											button : '.clickEvt',
											toggler : '.ly-title',
											panel : '.ly-content',
											// 주문취소/반품/교환 내 아코디언 안에 아코디언
											parent_sub : '.accord-sub-mypage',
											button_sub : '.clickEvt_sub',
											toggler_sub : '.ly-title_sub',
											panel_sub : '.ly-content_sub'
										};

										module = {
											init : function(){
												$(selector.button).on('click', function(){
													module.accordion(this);
												});
												$(selector.button_sub).on('click', function(){
													module.accordion_sub(this);
												});
												$(window).trigger('scroll');
												$(selector.parent).eq(0).find($(selector.panel)).show();
												$(selector.parent).eq(0).find($(selector.toggler)).addClass('is-on');
												$(selector.parent_sub).eq(0).find($(selector.panel_sub)).show();
												$(selector.parent_sub).eq(0).find($(selector.toggler_sub)).addClass('is-on');
											},
											accordion : function(el){
												var target = $(el).data('target');

												$(selector.panel).slideUp(400);
												$(selector.toggler).removeClass('is-on');

												if($(selector.panel, '#' + target).css('display') === 'none'){
													$(selector.panel, '#' + target).slideDown(400);
													$(selector.toggler, '#' + target).addClass('is-on');
												}
											},
											accordion_sub : function(el){
												var target = $(el).data('target');

												$(selector.panel_sub).slideUp(300);
												$(selector.toggler_sub).removeClass('is-on');

												if($(selector.panel_sub, '#' + target).css('display') === 'none'){
													$(selector.panel_sub, '#' + target).slideDown(300);
													$(selector.toggler_sub, '#' + target).addClass('is-on');
												}
											}
										};
										module.init();
									}();
								</script>
<%
oRs.Close
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>