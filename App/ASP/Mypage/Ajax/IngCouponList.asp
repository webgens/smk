<%@ Language=VBScript codepage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'IngCouponList.asp - 배포중 쿠폰 리스트
'Date		: 2019.01.05
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

DIM ToDay : ToDay = R_YEAR & R_MONTH & R_DAY & R_HOUR & R_MIN
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'
	

SET oConn				 = ConnectionOpen()							'# 커넥션 생성
SET oRs					 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성



Response.Write "OK|||||"




SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Coupon_Select_For_Available_Coupon"
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

RecCnt	 = oRs.RecordCount

IF NOT oRs.EOF THEN	
%>
                                            <div class="coupon-lists">
<%
		Do Until oRs.EOF
%>
                                                <div class="coupon-list">
                                                    <div class="tit">
                                                        <div class="inn">
															<%IF oRs("DeliveryCouponFlag") = "Y" THEN%>
															<div class="off">교환/반품 무료</div>
															<%ELSE%>
																<%IF oRs("MoneyType") = "P" THEN%>
                                                            <div class="off"><%=oRs("Discount")%>% OFF</div>
																<%ELSE%>
															<div class="off"><%=FormatNumber(oRs("Discount"), 0)%><em>원</em></div>
																<%END IF%>
															<%END IF%>
                                                            <div class="name"><%=oRs("CouponName")%></div>
                                                        </div>
                                                    </div>
                                                    <div class="time-limit">
														<%IF oRs("UseDateType") = "U" THEN%>
														&nbsp;<br />기간제한 없음
														<%ELSEIF oRs("UseDateType") = "P" THEN%>
                                                        <em>사용기한</em><%=LEFT(oRs("UseSDate"), 4)%>.<%=MID(oRs("UseSDate"), 5, 2)%>.<%=MID(oRs("UseSDate"), 7, 2)%> ~
															<%IF LEFT(oRs("UseEDate"),1) = "9" THEN%>
														<br />기간제한 없음
															<%ELSE%>
														<br /><%=LEFT(oRs("UseEDate"), 4)%>.<%=MID(oRs("UseEDate"), 5, 2)%>.<%=MID(oRs("UseEDate"), 7, 2)%>
															<%END IF%>
														<%ELSEIF oRs("UseDateType") = "D" THEN%>
														&nbsp;<br />발급일로 <%=oRs("UseDay")%>일 이내
														<%END IF%>
                                                    </div>
                                                    <div class="issue">
                                                        <a href="javascript:void(0)" onclick="couponDown('<%=oRs("Idx")%>')" class="btn-down-coupon link">쿠폰받기</a>
                                                    </div>
                                                </div>
<%
				oRs.MoveNext
		Loop
%>
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



	
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>

|||||<%=RecCnt%>