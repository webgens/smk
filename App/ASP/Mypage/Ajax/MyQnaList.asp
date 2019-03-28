<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'MyQnaList.asp - 마이페이지 > 상품문의
'Date		: 2018.12.26
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
PageCode2 = "03"
PageCode3 = "05"
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->
<!-- #include virtual="/Common/CheckID_Ajax.asp" -->

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
DIM PageSize : PageSize = 1000
DIM RecCnt
DIM PageCnt
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'



'Page			 = Request("page")
'If Page = "" Then Page = 1
'
'
SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
'
'wQuery = "WHERE B.IsShowFlag = 'Y' AND B.ProductType IN ('P','O') AND B.OrderState = '7' "
'IF U_NUM <> "" THEN
'		wQuery = wQuery & "AND A.UserID = '" & U_NUM & "' "
'ELSEIF N_NAME <> "" THEN
'		wQuery = wQuery & "AND (A.UserID = '' OR A.UserID IS NULL) AND A.OrderName = '" & N_NAME & "' AND A.OrderHp = '" & N_HP & "' AND A.OrderEmail = '" & N_EMAIL & "' "
'END IF
'sQuery = "ORDER BY A.OrderCode DESC, B.Idx "
'
'SET oCmd = Server.CreateObject("ADODB.Command")
'WITH oCmd
'		.ActiveConnection	 = oConn
'		.CommandType		 = adCmdStoredProc
'		.CommandText		 = "USP_Mobile_EShop_Order_MyReview_Select"
'
'		.Parameters.Append .CreateParameter("@PAGE",		 adInteger, adParaminput,		, Page)
'		.Parameters.Append .CreateParameter("@PAGE_SIZE",	 adInteger, adParaminput,		, PageSize)
'		.Parameters.Append .CreateParameter("@wQuery",		 adVarchar, adParaminput, 1000	, wQuery)
'		.Parameters.Append .CreateParameter("@sQuery",		 adVarchar, adParaminput, 100	, sQuery)
'END WITH
'oRs.CursorLocation = adUseClient
'oRs.Open oCmd, , adOpenStatic, adLockReadOnly
'SET oCmd = Nothing
'
'RecCnt	 = oRs(0)
'PageCnt	 = FIX((RecCnt+(PageSize-1))/PageSize)
'
'SET oRs = oRs.NextrecordSet

Response.Write "OK|||||"
%>


                            <div id="tabs_qna" class="tab">
                                <div class="tab-mypage">
                                    <ul class="tab-selector">
                                        <li class="part-2"><a href="javascript:productQnaList(1);" data-target="tabs-col4">상품 Q&amp;A관리</a></li>
                                        <li class="part-2"><a href="javascript:mtmQnaList(1);" data-target="tabs-col5">1:1 문의</a></li>
                                    </ul>
                                    <!-- 상품 Q&A 관리 -->
                                    <div id="tabs-col4" class="tab-panel">
                                    </div>
	                                <!-- // 상품 Q&A 관리 -->

									<!-- 1:1 문의 -->
									<div id="tabs-col5" class="tab-panel">
									</div>
									<!-- // 1:1 문의 -->
                                </div>
                            </div>


							<script>
								$(function () {
									var Tabs = {
										selector: {
											container: '.tab',
											panel: '.tab-panel',
											list: '.tab-selector li',
											item: '.tab-selector a'
										},
										build: function(i, el) {
											if (Tabs.getQueryStringUse()) {
												var
													pathname = location.search.queryStringToObject(),
													getLoc = $(Tabs.selector.container).data('use');

												if (pathname[getLoc] == undefined) {
													$(el).find(Tabs.selector.panel).eq(0).addClass('active');
													$(el).find(Tabs.selector.list + ':first-child').addClass('active');
												} else {
													$(el).find(Tabs.selector.panel).eq(parseInt(pathname[getLoc])).addClass('active');
													$(el).find(Tabs.selector.list).eq(parseInt(pathname[getLoc])).addClass('active');
												}
											} else {
												$(el).find(Tabs.selector.panel).eq(0).addClass('active');
												$(el).find(Tabs.selector.list + ':first-child').addClass('active');
											}
										},
										getQueryStringUse: function() {
											if ($(Tabs.selector.container).data('use') == undefined) {
												return false;
											} else {
												return true;
											}
										},
										getData: function(el) {
											$wrap = '#' + $(el).closest(Tabs.selector.container).attr('id');
											$target = $($wrap).find($('#' + $(el).data('target')));
											Tabs.openPanel($target, $wrap);
											Tabs.classChange(el, $wrap);
										},
										classChange: function(el, wr) {
											$(wr).find(Tabs.selector.list).removeClass('active');
											$(el).parent(Tabs.selector.list).addClass('active');
										},
										openPanel: function(el, wr) {
											$(wr).find(Tabs.selector.panel).removeClass('active');
											$(el).addClass('active');
										}
									};


									var tabBuild = function() {
										$(Tabs.selector.container).each(function(i, el) {
											Tabs.build(i, el)
										});
										$(Tabs.selector.item).on('click', function() {
											Tabs.getData(this)
										});
									};

									tabBuild();
								});

								productQnaList(1);
							</script>
<%
'oRs.Close
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>