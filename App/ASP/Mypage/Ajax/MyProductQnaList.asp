<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'MyMtmQnaList.asp - 마이페이지 > 상품문의 > 상품Q&A관리
'Date		: 2018.12.27
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
PageCode4 = "01"
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
DIM oConn							'# ADODB Connection 개체
DIM oRs								'# ADODB Recordset 개체
DIM oCmd							'# ADODB Command 개체

DIM wQuery							'# WHERE 절
DIM sQuery							'# SORT 절

Dim i
Dim x
Dim y

DIM Page
DIM PageSize : PageSize = 10
DIM RecCnt
DIM PageCnt
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'



Page			 = Request("page")
If Page = "" Then Page = 1


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성

wQuery = "WHERE A.DelFlag = 0 "' AND B.UserID = '" & U_ID & "' "
sQuery = "ORDER BY A.IDX DESC "
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Mobile_EShop_Product_QNA_Select"

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

Response.Write "OK|||||"
%>
                                        <div class="h-line">
                                            <h2 class="h-level4">내가 쓴 상품문의</h2>
                                            <span class="h-num"><%=RecCnt%>건</span>
                                            <span class="h-date is-right">
                                                <button type="button" class="button-ty3 ty-bd-black">
                                                    <span class="icon ico-inquire" onclick="insert_MyMtmQna('add');">1:1 문의하기</span>
                                            </button>
                                            </span>
                                        </div>
                                        <ul class="informView">
<%
IF NOT oRs.EOF THEN
	i = 0
	DO UNTIL oRs.EOF
%>
                                            <li class="informItem">
                                                <a href="#">
                                                    <span class="head-tit">
                                                        <span class="tit">[<%=oRs("ClassName")%>]</span>
                                                        <span class="date">작성일 : <%=LEFT(oRs("CreateDT"), 10)%></span>
                                                    </span>
                                                    <span class="cont">
                                                        <span class="thumbNail">
                                                            <span class="img">
                                                                <img src="<%=oRs("ProductImg")%>" alt="<%=oRs("ProductName")%>">
                                                            </span>
                                                        </span>
                                                        <span class="detail">
                                                            <span class="brand">
                                                                <span class="name"><%=oRs("BrandName") %></span>
                                                                <span class="item-code"><%=oRs("ProdCD") %></span>
                                                            </span>
                                                            <span class="product-name"><em><%=oRs("ProductName")%></em></span>

                                                            <span class="product-more">
                                                                <span class="price"><strong><%=FormatNumber(oRs("SalePrice"),0) %></strong>원</span>
                                                            </span>
                                                        </span>
                                                    </span>
                                                </a>
                                                <div class="inquire-cnt">
                                                    <p class="tit"><span class="bold">Q.</span><%=oRs("Title")%></p>
                                                    <p class="cnt"><%=oRs("Contents")%></p>
                                                </div>
<%
		If oRs("Reply_Flag") = "Y" Then
%>
                                                <div class="ly-qtitle_sub">
                                                    <button type="button" class="clickEvt_sub btn-list" data-target="answer-complete1">답변 완료</button>
                                                </div>
                                                <div class="ly-accord-sub">
                                                    <div id="answer-complete<%=i%>" class="accord-sub-mypage">

                                                        <div class="ly-qcontent_sub">
                                                            <div class="inquire-cnt answer-area">
                                                                <p class="tit"><span class="bold">A.</span>고객님 문의사항에 답변 드립니다.</p>
                                                                <p class="cnt"><%=oRs("Reply_Contents")%></p>
                                                            </div>
                                                        </div>
                                                    </div>
                                                </div>
<%
		ELSE
%>
                                                <span class="answer-wait">답변 대기</span>
<%
		END IF
%>
                                            </li>
<%
	i = i + 1 
	oRs.MOVENEXT
	LOOP
ELSE
%>
                                            <li class="informItem">
                                                <div class="inquire-cnt" style="text-align:center;">
                                                    <p class="tit"><span class="bold"></span>등록된 문의글이 없습니다.</p>
                                                </div>
                                            </li>
<%
END IF

oRs.Close
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>