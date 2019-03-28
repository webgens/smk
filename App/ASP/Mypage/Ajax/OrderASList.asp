<%@ Language=VBScript codepage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'OrderAsList.asp - 마이페이지 > A/S신청내역
'Date		: 2019.01.24
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
DIM j
DIM x
DIM y

DIM Page
DIM PageSize : PageSize = 100
DIM RecCnt
DIM PageCnt

DIM SDate
DIM EDate
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'
	
	
Page			 = sqlFilter(Request("Page"))
IF Page			 = "" THEN Page	 = 1

SDate			 = sqlFilter(Request("SDate"))
EDate			 = sqlFilter(Request("EDate"))

IF SDate		 = "" THEN SDate		= DateAdd("m", -1, Date)
IF EDate		 = "" THEN EDate		= Date


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성






	wQuery = "WHERE 0=0 "

	IF SDate <> "" AND EDate <> "" THEN
			wQuery = wQuery & "AND Convert(Varchar(10), A.CreateDT, 121) BetWeen '" & SDate & "' AND '" & EDate & "' "
	END IF
	IF U_NUM <> "" THEN
			wQuery = wQuery & "AND B.UserID = '" & U_NUM & "' "
	ELSEIF N_NAME <> "" THEN
			wQuery = wQuery & "AND (B.UserID = '' OR B.UserID IS NULL) AND B.OrderName = '" & N_NAME & "' AND B.OrderHp = '" & N_HP & "' AND B.OrderEmail = '" & N_EMAIL & "' "
	END IF
	sQuery = "ORDER BY A.IDX DESC "


	SET oCmd = Server.CreateObject("ADODB.Command")
	WITH oCmd
			.ActiveConnection	 = oConn
			.CommandType		 = adCmdStoredProc
			.CommandText		 = "USP_Front_Order_AfterService_Select"

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

	IF oRs.EOF THEN	
%>
                                            <li class="informItem">
                                                <a>
													<span class="cont" style="text-align:center;padding-top:60px;">
														신청한 내용이 없습니다.
													</span>
												</a>
											</li>
<%
	ELSE
		DO UNTIL oRs.EOF
%>
                                            <li class="informItem">
                                                <a>
													<span class="head-tit">
														<span class="tit">주문번호 : <%=oRs("OrderCode")%></span>
														<span class="date"><%=GetDateYMD(oRs("OrderDate"))%></span>
													</span>
													<span class="cont">
														<span class="thumbNail">
															<span class="img">
																<img src="<%=oRs("ProductImage")%>" alt="<%=oRs("ProductName")%>">
															</span>
															<span class="about">
																<span class="process"></span>
																<span class="date"><%=oRs("StateName")%> (<%=oRs("RequestName")%>)</span>
															</span>
														</span>
							
														<span class="detail">
															<span class="brand">
																<span class="name"><%=oRs("BrandName")%></span>
																<span class="item-code"><%=oRs("ProdCD")%>-<%=oRs("ColorCD")%></span>
															</span>
															<span class="product-name"><em><%=oRs("ProductName")%></em></span>
															
															<span class="inform">
																<span class="list">
																	<span class="tit">옵션</span>
																	<span class="opt"><%=oRs("SizeCD")%></span>
																</span>
																<span class="list">
																	<span class="tit">수량</span>
																	<span class="opt"><%=oRs("OrderCnt")%></span>
																</span>
																<span class="list">
																	<span class="tit">결제금액</span>
																	<span class="opt price"><em><%=FormatNumber(oRs("OrderPrice"),0)%></em>원</span>
																</span>
																<span class="list">
																	<span class="tit">주문일</span>
																	<span class="opt"><%=GetDateYMD(oRs("OrderDate"))%></span>
																</span>
															</span>
														</span>
													</span>
												</a>

                                            </li>
<%
			oRs.MoveNext
			Loop

	END IF
	oRs.Close

SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>