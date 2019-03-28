<%@ Language=VBScript codepage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'ShoesGiftList.asp - 슈즈 상품권 리스트
'Date		: 2019.01.07
'Update		: 
'*****************************************************************************************'
	
'# 페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include Virtual = "/ADO/ADODBCommon.asp" -->
<!-- #include Virtual = "/Common/Common.asp" -->
<!-- #include Virtual = "/Common/SetInfo.asp" -->
<!-- #include Virtual = "/Common/CheckID_Ajax.asp" -->

<%
'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn							'# ADODB Connection 개체
DIM oRs								'# ADODB Recordset 개체
DIM oCmd							'# ADODB Command 개체

DIM wQuery							'# WHERE 절
DIM sQuery							'# SORT 절

DIM x
DIM y

DIM Page
DIM PageSize : PageSize = 5
DIM RecCnt
DIM PageCnt

DIM SDate
DIM EDate

DIM AvailableDT
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'
	
	
Page			 = sqlFilter(Request("Page"))
SDate			 = sqlFilter(Request("SDate"))
EDate			 = sqlFilter(Request("EDate"))
IF Page			 = "" THEN Page	 = 1


SET oConn				 = ConnectionOpen()							'# 커넥션 생성
SET oRs					 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성




Response.Write "OK|||||"




wQuery = "WHERE A.MemberNum = " & U_NUM & "  AND A.DelFlag = 'N' "
'wQuery = "WHERE 1 = 1 "
IF SDate <> "" THEN
		wQuery = wQuery & "AND A.CreateDT >= '" & SDate & "' "
END IF
IF EDate <> "" THEN
		wQuery = wQuery & "AND A.CreateDT < '" & DateAdd("d", 1, EDate) & "' "
END IF

sQuery = "ORDER BY A.Idx DESC "




SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_Member_SCash_Select"

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
                            <div class="h-line">
                                <h2 class="h-level4">등록/사용내역</h2>
                                <span class="h-num"><%=RecCnt%>개</span>
                                <span class="h-date is-right"><%=SDate%> ~ <%=EDate%></span>
                            </div>
<%
IF NOT oRs.EOF THEN	
%>
                            <div class="ly-gift-card">
                                <ul>
<%
		Do Until oRs.EOF
				AvailableDT = oRs("AvailableDT")
				IF ISNULL(oRs("AvailableDT")) THEN AvailableDT = "1970-01-01"

				IF oRs("AddSCash") > 0 AND CDate(AvailableDT) > Date THEN
%>
                                    <li>
                                        <div class="h-gift-card">
                                            <span class="name">SHOES GIFT</span>
                                            <em class="price"><%=FormatNumber(ABS(oRs("AddSCash")), 0)%></em>
                                            <p class="tit"><%=oRs("Memo")%></p>
                                        </div>
                                        <div class="f-gift-card">
                                            <div class="date">등록일 : <%=REPLACE(LEFT(oRs("CreateDT"), 10), "-", ".")%></div>
                                            <div class="condition">적립</div>
                                        </div>
                                    </li>
<%
				ELSE
%>
                                    <li>
                                        <div class="h-gift-card">
                                            <span class="name">SHOES GIFT</span>
                                            <em class="price"><%=FormatNumber(ABS(oRs("AddSCash")), 0)%></em>
                                            <p class="tit"><%=oRs("Memo")%></p>
                                        </div>
                                        <div class="f-gift-card">
                                            <div class="date">등록일 : <%=REPLACE(LEFT(oRs("CreateDT"), 10), "-", ".")%></div>
                                            <div class="condition">사용</div>
                                        </div>
                                    </li>
<%
				END IF

				oRs.MoveNext
		Loop
%>
                                <ul>
							</div>
<%
ELSE
%>
							<div class="area-empty">
								<span class="icon-empty"></span>
								<p class="tit-empty">등록된 정보가 없습니다.</p>
							</div>
<%
END IF
oRs.Close
%>


<%
IF RecCnt > 0 THEN	
%>
                        <div class="area-pagination">
<%
							y = Int((Page-1) / 5) * 5 + 1

							IF y = 1 THEN
									Response.Write "						<a class=""btn-prev"">이전</a>"&vbLf
							ELSE
									Response.Write "						<a href=""javascript:get_ShoesGiftList(" & y - 10 & ")"" class=""btn-prev1"">이전</a>"&vbLf
							END IF

							x = 1
							Do Until x > 10 OR y > PageCnt

									IF y = int(Page) THEN
											IF CDbl(x) < 10 AND CDbl(y) < CDbl(PageCnt) THEN
													Response.Write "<span class=""page-num  current""><a href=""javascript:get_ShoesGiftList(" & y & ")"">" & y & "</a></span>"
											ELSE
													Response.Write "<span class=""page-num1 current""><a href=""javascript:get_ShoesGiftList(" & y & ")"">" & y & "</a></span>"
											END IF
									ELSE
											IF CDbl(x) < 10 AND CDbl(y) < CDbl(PageCnt) THEN
													Response.Write "<span class=""page-num  point""><a href=""javascript:get_ShoesGiftList(" & y & ")"">" & y & "</a></span>"
											ELSE
													Response.Write "<span class=""page-num1 point""><a href=""javascript:get_ShoesGiftList(" & y & ")"">" & y & "</a></span>"
											END IF
									END IF

									y = y + 1
									x = x + 1
							Loop

							IF y > PageCnt THEN
									Response.Write "						<a class=""btn-next1"">다음</a>"&vbLf
							ELSE
									Response.Write "						<a href=""javascript:get_ShoesGiftList(" & y & ")"" class=""btn-next"">다음</a>"&vbLf
							END IF
%>
                        </div>
<%
END IF
	
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>