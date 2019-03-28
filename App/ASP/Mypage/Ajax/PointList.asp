<%@ Language=VBScript codepage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'PointList.asp - 포인트 리스트
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
DIM PageSize : PageSize = 5
DIM RecCnt
DIM PageCnt

DIM RsCnt : RsCnt = 0
DIM ArrRs

DIM AvailableDT
DIM LiClass
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'
	
	
Page			 = sqlFilter(Request("Page"))
IF Page			 = "" THEN Page	 = 1


SET oConn				 = ConnectionOpen()							'# 커넥션 생성
SET oRs					 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성




SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_Member_Point_PCode_Select"
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

IF NOT oRs.EOF THEN
		RsCnt = oRs.RecordCount
		ArrRs = oRs.GetRows(RsCnt)
END IF
oRs.Close


FUNCTION get_PointPName(ByVal pCode)
		DIM RetVal : RetVal = ""
		IF RsCnt > 0 THEN
				FOR i = 0 TO UBound(ArrRs, 2)
						IF pCode = ArrRs(0, i) THEN
								RetVal = ArrRs(1, i)
								EXIT FOR
						END IF
				NEXT
		END IF
		get_PointPName = RetVal
END FUNCTION





Response.Write "OK|||||"




wQuery = "WHERE A.MemberNum = " & U_NUM & "  AND A.DelFlag = 'N' "
'wQuery = "WHERE 1 = 1 "

sQuery = "ORDER BY A.Idx DESC "




SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_Member_Point_Select"

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
                                <h2 class="h-level4">적립/사용내역</h2>
                                <span class="h-num"><%=RecCnt%>건</span>
                            </div>
<%
IF NOT oRs.EOF THEN	
%>
                            <div class="point-lists">
                                <ul>
<%
		x = 1
		Do Until oRs.EOF
				AvailableDT = oRs("AvailableDT")
				IF ISNULL(oRs("AvailableDT")) THEN AvailableDT = "1970-01-01"

				LiClass = ""
				IF oRs("PCode") = "A01" OR (oRs("AddPoint") > 0 AND CDate(AvailableDT) < Date) THEN LiClass = "bgc-type2"
%>
                                    <li>
                                        <div class="cont">
                                            <span class="date"><%=REPLACE(LEFT(oRs("CreateDT"), 10), "-", ". ")%></span>
                                            <p class="tit"><%=get_PointPName(oRs("PCode"))%></p>
                                            <p class="item-name ellipsis"><%=oRs("Memo")%></p>

											<%IF oRs("AddPoint") > 0 THEN%>
                                            <div class="point saving">+<%=FormatNumber(ABS(oRs("AddPoint")), 0)%></div>
											<%ELSE%>
											<div class="point used"><%=FormatNumber(ABS(oRs("AddPoint")), 0)%></div>
											<%END IF%>
                                        </div>
										<%IF oRs("AddPoint") > 0 THEN%>
                                        <button type="button" onclick="pointMore(<%=X%>)" class="btn-more" data-target="pointList_<%=x%>">내역보기</button>
                                        <div class="check" id="pointList_<%=x%>">
                                            <p>(총 적립 <%=FormatNumber(oRs("AddPoint"), 0)%>원) + (사용 -<%=FormatNumber(oRs("UsePoint"), 0)%>원) = 잔여 <%=FormatNumber(oRs("AddPoint") - oRs("UsePoint"), 0)%>원</p>
                                        </div>
										<%END IF%>
                                    </li>
<%
				oRs.MoveNext
				x = x + 1
		Loop
%>
                                </ul>
                            </div>
<%
ELSE
%>
							<div class="area-empty" style="margin-top: 30px;">
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
							y = Int((Page-1) / 10) * 10 + 1

							IF y = 1 THEN
									Response.Write "						<a class=""btn-prev"">이전</a>"&vbLf
							ELSE
									Response.Write "						<a href=""javascript:get_PointList(" & y - 10 & ")"" class=""btn-prev1"">이전</a>"&vbLf
							END IF

							x = 1
							Do Until x > 10 OR y > PageCnt

									IF y = int(Page) THEN
											IF CDbl(x) < 10 AND CDbl(y) < CDbl(PageCnt) THEN
													Response.Write "<span class=""page-num  current""><a href=""javascript:get_PointList(" & y & ")"">" & y & "</a></span>"
											ELSE
													Response.Write "<span class=""page-num1 current""><a href=""javascript:get_PointList(" & y & ")"">" & y & "</a></span>"
											END IF
									ELSE
											IF CDbl(x) < 10 AND CDbl(y) < CDbl(PageCnt) THEN
													Response.Write "<span class=""page-num  point""><a href=""javascript:get_PointList(" & y & ")"">" & y & "</a></span>"
											ELSE
													Response.Write "<span class=""page-num1 point""><a href=""javascript:get_PointList(" & y & ")"">" & y & "</a></span>"
											END IF
									END IF

									y = y + 1
									x = x + 1
							Loop

							IF y > PageCnt THEN
									Response.Write "						<a class=""btn-next1"">다음</a>"&vbLf
							ELSE
									Response.Write "						<a href=""javascript:get_PointList(" & y & ")"" class=""btn-next"">다음</a>"&vbLf
							END IF
%>
                        </div>


						<script>
							function pointMore(num) {
								$(".check").slideUp(200);
								if ($("#pointList_" + num).css("display") == "none") {
									$("#pointList_" + num).slideDown(200);
								}
								else {
									$("#pointList_" + num).slideUp(200);
								}
							}
						</script>

<%
END IF
	
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>