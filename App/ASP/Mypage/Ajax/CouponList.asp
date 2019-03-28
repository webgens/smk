<%@ Language=VBScript codepage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'CouponList.asp - 쿠폰 리스트
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

DIM i
DIM x
DIM y

DIM Page
DIM PageSize : PageSize = 6
DIM RecCnt
DIM PageCnt

DIM Useable

DIM ToDay : ToDay = R_YEAR & R_MONTH & R_DAY & R_HOUR & R_MIN
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'
	
	
Page			 = sqlFilter(Request("Page"))
Useable			 = sqlFilter(Request("Useable"))
IF Page			 = "" THEN Page		 = 1
IF Useable		 = "" THEN Useable	 = "Y"


SET oConn				 = ConnectionOpen()							'# 커넥션 생성
SET oRs					 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성



Response.Write "OK|||||"




wQuery = "WHERE A.MemberNum = " & U_NUM & "  AND A.ReceiveFlag = 'Y' AND A.CollectFlag = 'N' AND B.PCFlag = 'Y' "
'wQuery = "WHERE 1 = 1 "

IF Useable = "Y" THEN
		wQuery = wQuery & "AND A.UseFlag = 'N' AND A.StartDT <= '" & ToDay & "' AND A.EndDT >= '" & ToDay & "' "
ELSE
		wQuery = wQuery & "AND (A.UseFlag = 'Y' OR A.EndDT < '" & ToDay & "') "
END IF

sQuery = "ORDER BY A.Idx DESC "




SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Coupon_Member_Select"

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


IF NOT oRs.EOF THEN	
%>
                                            <div class="coupon-lists">
<%
		Do Until oRs.EOF
				IF Useable = "Y" THEN
%>
                                                <div class="coupon-list">
                                                    <div class="tit">
                                                        <div class="inn">
															<%IF oRs("MoneyType") = "P" THEN%>
                                                            <div class="off"><%=oRs("Discount")%>% OFF</div>
															<%ELSE%>
															<div class="off"><%=FormatNumber(oRs("Discount"), 0)%><em>원</em></div>
															<%END IF%>
                                                            <div class="name"><%=oRs("CouponName")%></div>
                                                        </div>
                                                    </div>
                                                    <div class="time-limit">
                                                        <em>사용기한</em><%=LEFT(oRs("StartDT"), 4)%>.<%=MID(oRs("StartDT"), 5, 2)%>.<%=MID(oRs("StartDT"), 7, 2)%>(<%=MID(oRs("StartDT"), 9, 2)%>:<%=MID(oRs("StartDT"), 11, 2)%>) ~
														<%IF LEFT(oRs("EndDT"),1) = "9" THEN%>
														<br />기간제한 없음
														<%ELSE%>
														<br><%=LEFT(oRs("EndDT"), 4)%>.<%=MID(oRs("EndDT"), 5, 2)%>.<%=MID(oRs("EndDT"), 7, 2)%>(<%=MID(oRs("EndDT"), 9, 2)%>:<%=MID(oRs("EndDT"), 11, 2)%>)
														<%END IF%>
                                                    </div>
                                                    <div class="issue">
														<%IF oRs("DeliveryCouponFlag") <> "Y" THEN%>
                                                        <a href="/ASP/Mypage/CouponApplyProductList.asp?Idx=<%=oRs("Idx")%>" class="btn-down-coupon link">적용상품</a>
														<%ELSE%>
														<span class="btn-down-coupon">교환반품</span>
														<%END IF%>
                                                    </div>
                                                </div>
<%
				ELSE
%>
                                                <div class="coupon-list">
                                                    <div class="tit overdue">
                                                        <div class="inn">
															<%IF oRs("MoneyType") = "P" THEN%>
                                                            <div class="off"><%=oRs("Discount")%>% OFF</div>
															<%ELSE%>
															<div class="off"><%=FormatNumber(oRs("Discount"), 0)%><em>원</em></div>
															<%END IF%>
                                                            <div class="name"><%=oRs("CouponName")%></div>
                                                        </div>
                                                    </div>
                                                    <div class="time-limit">
                                                        <em>사용기한</em><%=LEFT(oRs("StartDT"), 4)%>.<%=MID(oRs("StartDT"), 5, 2)%>.<%=MID(oRs("StartDT"), 7, 2)%>(<%=MID(oRs("StartDT"), 9, 2)%>:<%=MID(oRs("StartDT"), 11, 2)%>) ~
														<%IF LEFT(oRs("EndDT"),1) = "9" THEN%>
														<br />기간제한 없음
														<%ELSE%>
														<br><%=LEFT(oRs("EndDT"), 4)%>.<%=MID(oRs("EndDT"), 5, 2)%>.<%=MID(oRs("EndDT"), 7, 2)%>(<%=MID(oRs("EndDT"), 9, 2)%>:<%=MID(oRs("EndDT"), 11, 2)%>)
														<%END IF%>
                                                    </div>
                                                    <div class="issue">
														<%IF oRs("UseFlag") = "Y" THEN%>
                                                        <span class="btn-down-coupon">사용완료</span>
														<%ELSE%>
														<span class="btn-down-coupon">기한만료</span>
														<%END IF%>
                                                    </div>
                                                </div>
<%
				END IF
%>
<%
				oRs.MoveNext
		Loop
%>
                                        </ul>
                                    </div>
<%
ELSE
%>
                                    <div class="area-empty">
                                        <span class="icon-empty"></span>
                                        <p class="tit-empty">보유중인 쿠폰 내역이 없습니다.</p>
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
												Response.Write "						<a href=""javascript:get_CouponList(" & y - 10 & ")"" class=""btn-prev1"">이전</a>"&vbLf
										END IF

										x = 1
										Do Until x > 10 OR y > PageCnt

												IF y = int(Page) THEN
														IF CDbl(x) < 10 AND CDbl(y) < CDbl(PageCnt) THEN
																Response.Write "<span class=""page-num  current""><a href=""javascript:get_CouponList(" & y & ")"">" & y & "</a></span>"
														ELSE
																Response.Write "<span class=""page-num1 current""><a href=""javascript:get_CouponList(" & y & ")"">" & y & "</a></span>"
														END IF
												ELSE
														IF CDbl(x) < 10 AND CDbl(y) < CDbl(PageCnt) THEN
																Response.Write "<span class=""page-num  point""><a href=""javascript:get_CouponList(" & y & ")"">" & y & "</a></span>"
														ELSE
																Response.Write "<span class=""page-num1 point""><a href=""javascript:get_CouponList(" & y & ")"">" & y & "</a></span>"
														END IF
												END IF

												y = y + 1
												x = x + 1
										Loop

										IF y > PageCnt THEN
												Response.Write "						<a class=""btn-next1"">다음</a>"&vbLf
										ELSE
												Response.Write "						<a href=""javascript:get_CouponList(" & y & ")"" class=""btn-next"">다음</a>"&vbLf
										END IF
%>
							        </div>


<%
END IF
	
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>

|||||<%=RecCnt%>