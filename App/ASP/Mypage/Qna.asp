<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'Qna.asp - 상품문의 / 1:1문의
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
PageCode1 = "05"
PageCode2 = "03"
PageCode3 = "00"
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->
<!-- #include virtual="/Common/SubCheckID.asp" -->

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
DIM PageSize : PageSize = 10
DIM RecCnt
DIM PageCnt

DIM QnaType						'# 탭구분
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

Page			 = sqlFilter(Request("Page"))
IF Page = "" THEN Page = 1

QnaType			 = sqlFilter(request("QnaType"))
IF QnaType = "" THEN QnaType	= "1"


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>


<!-- #include virtual="/INC/Header.asp" -->
	<script type="text/javascript">
		// 탭 선택시
		function chgTab(listType) {
			$("#tabs .tab-selector li").removeClass("active");
			//$("#tabs-col1").removeClass("active");
			//$("#tabs-col2").removeClass("active");

			if (listType == "1") {
				$("#tabs .tab-selector li").eq(0).addClass("active");
				location.replace("/ASP/Mypage/Qna.asp?QnaType=1");
				//$("#tabs-col1").addClass("active");
			}
			else {
				$("#tabs .tab-selector li").eq(1).addClass("active");
				location.replace("/ASP/Mypage/Qna.asp?QnaType=2");
				//$("#tabs-col2").addClass("active");
			}
		}

		function answerView(num) {
			var target = $("#question" + num).data('target');

			if ($("#question" + num).closest(".ly-title_sub").hasClass("is-on")) {
				$("#question" + num).closest(".ly-title_sub").removeClass("is-on");
				$("#" + target).find(".ly-content_sub").slideUp(300);
			} else {
				$(".ly-title_sub").removeClass("is-on");
				$(".ly-content_sub").slideUp(300);
				$("#question" + num).closest(".ly-title_sub").addClass("is-on");
				$("#" + target).find(".ly-content_sub").slideDown(300);
			}
			/*
			$(".ly-title_sub").removeClass("is-on");
			$("#question" + num).closest(".ly-title_sub").addClass("is-on");

			var _toggler_sub = $(".ly-title_sub");
			var _panel_sub = $(".ly-content_sub");

			var target = _this.data('target');

			$(_panel_sub).slideUp(300);
			$(_toggler_sub).removeClass('is-on');

			if ($('#' + target).css('display') === 'none') {
				$('#' + target).slideDown(300);
				$('#' + target).addClass('is-on');
			}
			*/
		}
	</script>

<%TopSubMenuTitle = "MY슈마커"%>
<!-- #include virtual="/INC/TopSub.asp" -->


    <!-- Main -->
    <main id="container" class="container">
        <div class="sub_content">

            <div class="wrap-mypage">
				<div style="height:8px"></div>
				
                <div id="MypageSubMenu" class="ly-title accordion">
                    <div class="selector">
	                    <button type="button" class="btn-list clickEvt" data-target="MypageSubMenu">상품문의</button>
					</div>
					<div class="option my-recode">
						<!-- #include virtual="/ASP/Mypage/SubMenu_MyShoeMarker.asp" -->
					</div>
                </div>


                <div class="mypage-my-inquire">
                    <div id="shoppingList">
                        <div>
                            <div id="tabs">
                                <div class="tab-mypage">
                                    <ul class="tab-selector">
                                        <li class="part-2<%IF QnaType = "1" THEN%> active<%END IF%>"><a href="javascript:void(0)" onclick="chgTab('1')">상품 Q&amp;A관리</a></li>
                                        <li class="part-2<%IF QnaType = "2" THEN%> active<%END IF%>"><a href="javascript:void(0)" onclick="chgTab('2')">1:1 문의</a></li>
                                    </ul>
<%
'# 상품 Q&A 관리
IF QnaType = "1" THEN
%>
<%
		wQuery = "WHERE A.DelFlag = 0 AND B.UserID = '" & U_ID & "' "
		sQuery = "ORDER BY A.IDX DESC "

		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Admin_EShop_Product_QNA_Select"

				.Parameters.Append .CreateParameter("@PAGE",		 adInteger, adParaminput,		, Page)
				.Parameters.Append .CreateParameter("@PAGE_SIZE",	 adInteger, adParaminput,		, PageSize)
				.Parameters.Append .CreateParameter("@wQuery",		 adVarchar, adParaminput, 1000	, wQuery)
				.Parameters.Append .CreateParameter("@sQuery",		 adVarchar, adParaminput, 100	, sQuery)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		SET oCmd = Nothing

		RecCnt	 = oRs(0)
		PageCnt	 = FIX((RecCnt+(PageSize-1))/PageSize)

		SET oRs = oRs.NextrecordSet
%>
                                    <!-- 상품 Q&A 관리 -->
                                    <div id="tabs-col1" class="tab-panel active">
                                        <div class="h-line">
                                            <h2 class="h-level4">내가 쓴 상품문의</h2>
                                            <span class="h-num"><%=FormatNumber(RecCnt,0)%>건</span>
                                            <span class="h-date is-right">
                                            </span>
                                        </div>
<%
		IF NOT oRs.EOF THEN
%>
                                        <ul class="informView">
<% 
				i = 1
				Do Until oRs.EOF
%>
                                            <li class="informItem">
                                                <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>')">
                                                    <span class="head-tit">
                                                        <span class="tit">[<%=oRs("ClassName")%>]</span>
                                                        <span class="date">작성일 : <%=Left(oRs("CreateDT"),10)%></span>
                                                    </span>
                                                    <span class="cont">
                                                        <span class="thumbNail">
                                                            <span class="img">
                                                                <img src="<%=oRs("ProductImg_180")%>" alt="상품 이미지">
                                                            </span>
                                                        </span>
                                                        <span class="detail">
                                                            <span class="brand">
                                                                <span class="name"><%=oRs("BrandName")%></span>
                                                            </span>
                                                            <span class="product-name"><em><%=oRs("ProductName")%></em></span>
                                                        </span>
                                                    </span>
                                                </a>
                                                <div class="inquire-cnt">
                                                    <p class="tit"><span class="bold">Q.</span><%=oRs("Title")%></p>
                                                    <p class="cnt"><%=oRs("Contents")%></p>
                                                </div>
												<%IF oRs("Reply_Flag") = "Y" THEN%>
                                                <div class="ly-title_sub">
                                                    <button type="button" onclick="answerView('<%=i%>')" id="question<%=i%>" class="btn-list" data-target="answer-complete<%=i%>">답변 완료</button>
                                                </div>
                                                <div class="ly-accord-sub">
                                                    <div id="answer-complete<%=i%>">

                                                        <div class="ly-content_sub">
                                                            <div class="inquire-cnt answer-area">
                                                                <p class="tit"><span class="bold">A.</span>고객님 문의사항에 답변 드립니다.</p>
                                                                <p class="cnt"><%=oRs("Reply_Contents")%></p>
                                                            </div>
                                                        </div>
                                                    </div>
                                                </div>
												<%ELSE%>
                                                <span class="answer-wait">답변 대기</span>
												<%END IF%>
                                            </li>
<%
						oRs.MoveNext
						i = i + 1
				LOOP
%>
                                        </ul>

				                        <div class="area-pagination">
<%
											y = Int((Page-1) / 5) * 5 + 1

											IF y = 1 THEN
													Response.Write "						<a class=""btn-prev"">이전</a>"&vbLf
											ELSE
													Response.Write "						<a href=""?QnaType=" & QnaType & "&Page=" & y - 10 & """ class=""btn-prev1"">이전</a>"&vbLf
											END IF

											x = 1
											Do Until x > 5 OR y > PageCnt

													IF y = int(Page) THEN
															IF CDbl(x) < 5 AND CDbl(y) < CDbl(PageCnt) THEN
																	Response.Write "<span class=""page-num  current""><a href=""?QnaType=" & QnaType & "&Page=" & y & """>" & y & "</a></span>"
															ELSE
																	Response.Write "<span class=""page-num1 current""><a href=""?QnaType=" & QnaType & "&Page=" & y & """>" & y & "</a></span>"
															END IF
													ELSE
															IF CDbl(x) < 5 AND CDbl(y) < CDbl(PageCnt) THEN
																	Response.Write "<span class=""page-num  point""><a href=""?QnaType=" & QnaType & "&Page=" & y & """>" & y & "</a></span>"
															ELSE
																	Response.Write "<span class=""page-num1 point""><a href=""?QnaType=" & QnaType & "&Page=" & y & """>" & y & "</a></span>"
															END IF
													END IF

													y = y + 1
													x = x + 1
											Loop

											IF y > PageCnt THEN
													Response.Write "						<a class=""btn-next1"">다음</a>"&vbLf
											ELSE
													Response.Write "						<a href=""?QnaType=" & QnaType & "&Page=" & y & """ class=""btn-next"">다음</a>"&vbLf
											END IF
%>
							            </div>
<%
		ELSE
%>
										<div class="area-empty">
											<span class="icon-empty"></span>
											<p class="tit-empty">문의한 내역이 없습니다.</p>
										</div>
<%
		END IF
		oRs.close
%>
                                    </div>
                                    <!-- // 상품 Q&A 관리 -->
<%
'# 1:1 문의
ELSEIF QnaType = "2" THEN
%>
<%
		wQuery = "WHERE A.DeleteFlag = 'N' AND B.UserID = '" & U_ID & "' "
		sQuery = "ORDER BY A.IDX DESC "
		SET oCmd = Server.CreateObject("ADODB.Command")
		WITH oCmd
				.ActiveConnection	 = oConn
				.CommandType		 = adCmdStoredProc
				.CommandText		 = "USP_Admin_EShop_QNA_Select"

				.Parameters.Append .CreateParameter("@PAGE",		 adInteger, adParaminput,		, Page)
				.Parameters.Append .CreateParameter("@PAGE_SIZE",	 adInteger, adParaminput,		, PageSize)
				.Parameters.Append .CreateParameter("@wQuery",		 adVarchar, adParaminput, 1000	, wQuery)
				.Parameters.Append .CreateParameter("@sQuery",		 adVarchar, adParaminput, 100	, sQuery)
		END WITH
		oRs.CursorLocation = adUseClient
		oRs.Open oCmd, , adOpenStatic, adLockReadOnly
		SET oCmd = Nothing

		RecCnt	 = oRs(0)
		PageCnt	 = FIX((RecCnt+(PageSize-1))/PageSize)

		SET oRs = oRs.NextrecordSet
%>
									<!-- 1:1 문의 -->
									<div id="tabs-col2" class="tab-panel active">
										<div class="h-line">
											<h2 class="h-level4">1:1 문의</h2>
											<span class="h-num"><%=FormatNumber(RecCnt,0)%>건</span>
											<span class="h-date is-right">
												<button type="button" onclick="popMtmQnaAdd()" class="button-ty3 ty-bd-black">
													<span class="icon ico-inquire">1:1 문의하기</span>
											</button>
											</span>
										</div>
<%
		IF NOT oRs.EOF THEN
				i = 1
				Do Until oRs.EOF
%>
										<div class="inquire">
											<div class="inquire-cnt">
												<p class="tit"><span class="bold">Q.</span><%=oRs("Title")%></p>
												<p class="cnt"><%=oRs("Contents")%></p>
												<span class="date">작성일 : <%=Left(oRs("CreateDT"),10)%></span>
											</div>
											<%IF oRs("ReplyFlag") = "Y" THEN%>
											<div class="ly-title_sub">
												<button type="button" onclick="answerView('<%=i%>')" id="question<%=i%>" class="btn-list" data-target="answer-complete<%=i%>">답변 완료</button>
											</div>
											<div class="ly-accord-sub">
												<div id="answer-complete<%=i%>">

													<div class="ly-content_sub">
														<div class="inquire-cnt answer-area">
															<p class="tit"><span class="bold">A.</span>고객님 문의사항에 답변 드립니다.</p>
															<p class="cnt"><%=oRs("Reply_Contents")%></p>
														</div>
													</div>
												</div>
											</div>
											<%ELSE%>
											<span class="answer-wait">답변 대기</span>
											<%END IF%>
										</div>
<%
						oRs.MoveNext
						i = i + 1
				LOOP
%>
				                        <div class="area-pagination">
<%
											y = Int((Page-1) / 5) * 5 + 1

											IF y = 1 THEN
													Response.Write "						<a class=""btn-prev"">이전</a>"&vbLf
											ELSE
													Response.Write "						<a href=""?QnaType=" & QnaType & "&Page=" & y - 10 & """ class=""btn-prev1"">이전</a>"&vbLf
											END IF

											x = 1
											Do Until x > 5 OR y > PageCnt

													IF y = int(Page) THEN
															IF CDbl(x) < 5 AND CDbl(y) < CDbl(PageCnt) THEN
																	Response.Write "<span class=""page-num  current""><a href=""?QnaType=" & QnaType & "&Page=" & y & """>" & y & "</a></span>"
															ELSE
																	Response.Write "<span class=""page-num1 current""><a href=""?QnaType=" & QnaType & "&Page=" & y & """>" & y & "</a></span>"
															END IF
													ELSE
															IF CDbl(x) < 5 AND CDbl(y) < CDbl(PageCnt) THEN
																	Response.Write "<span class=""page-num  point""><a href=""?QnaType=" & QnaType & "&Page=" & y & """>" & y & "</a></span>"
															ELSE
																	Response.Write "<span class=""page-num1 point""><a href=""?QnaType=" & QnaType & "&Page=" & y & """>" & y & "</a></span>"
															END IF
													END IF

													y = y + 1
													x = x + 1
											Loop

											IF y > PageCnt THEN
													Response.Write "						<a class=""btn-next1"">다음</a>"&vbLf
											ELSE
													Response.Write "						<a href=""?QnaType=" & QnaType & "&Page=" & y & """ class=""btn-next"">다음</a>"&vbLf
											END IF
%>
							            </div>
<%
		ELSE
%>
										<div class="area-empty">
											<span class="icon-empty"></span>
											<p class="tit-empty">문의 내역이 없습니다.</p>
										</div>
<%
		END IF
		oRs.close
%>
									</div>
									<!-- // 1:1 문의 -->

<%
END IF
%>
								</div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </main>

<!-- #include virtual="/INC/FooterNoBNB.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>