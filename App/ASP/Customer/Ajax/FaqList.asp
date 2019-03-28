<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'FaqList.asp - 고객센터 > Faq 리스트
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

DIM Page	 : Page = 1
DIM PageSize
DIM RecCnt
DIM PageCnt


DIM IDX
DIM ClassName
DIM Title
DIM Contents
DIM CreateDT

DIM sClassName
DIM sKeyword
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


PageSize		 = Request("PageSize")
If PageSize		 = "" Then PageSize = 5

sClassName		 = sqlFilter(Request("ClassName"))
sKeyword		 = sqlFilter(Request("Keyword"))
Idx				 = sqlFilter(Request("Idx"))


SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성

wQuery = "WHERE A.DelFlag = 'N' "
IF sClassName <> "" THEN
	wQuery = wQuery & "AND A.ClassName = '"& sClassName &"' "
ELSEIF sKeyword <> "" THEN
	wQuery = wQuery & "AND A.Title Like '%"& sKeyword &"%' "
ELSEIF IDX <> "" THEN
	wQuery = wQuery & "AND A.IDX = '"& IDX &"' "
END IF

sQuery = "ORDER BY A.TopFlag DESC, A.IDX DESC "
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Front_EShop_FAQ_Select"

		.Parameters.Append .CreateParameter("@PAGE",		adInteger,	adParamInput,	  ,		Page)
		.Parameters.Append .CreateParameter("@PAGE_SIZE",	adInteger,	adParamInput,	  ,		PageSize)
		.Parameters.Append .CreateParameter("@WQUERY",		adVarchar,	adParamInput, 1000,		wQuery)
		.Parameters.Append .CreateParameter("@SQUERY",		adVarchar,	adParamInput,  100,		sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing

RecCnt	 = oRs(0)

SET oRs = oRs.NextrecordSet

Response.Write "OK|||||"
IF oRs.EOF THEN
%>
                                <div id="qna0" class="qna">
                                    <div class="accord-sub-mypage">
                                        <div class="ly-content_sub">
                                            <p>등록된 FAQ가 없습니다.</p>
                                        </div>
                                    </div>
                                </div>
<%
ELSE
%>
<%
	j = 0
	DO UNTIL oRs.EOF
		ClassName		= oRs("ClassName")
		Title			= oRs("Title")
		Contents		= oRs("Contents")
		CreateDT		= LEFT(oRs("CreateDT"), 10)
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
								<input type="hidden" name="RecCnt" id="RecCnt" value="<%=RecCnt%>" />
								<input type="hidden" name="PageSize" id="PageSize" value="<%=PageSize%>" />
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