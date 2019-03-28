<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'EventList.asp - 이벤트 리스트
'Date		: 2019.01.12
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
PageCode1 = "00"
PageCode2 = "00"
PageCode3 = "00"
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->

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

Dim StoreProcName

Dim ISTopN
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

Page			 = sqlFilter(Request("Page"))
IF Page			 = "" THEN Page			 = 1

ISTopN			 = sqlFilter(Request("ISTopN"))

SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성


wQuery = "WHERE A.DelFlag = 'N' AND A.DisplayFlag = 'Y' AND A.EType = 'E' AND A.SDate <= '" & U_DATE&LEFT(U_TIME, 4) & "' AND A.EDate >= '" & U_DATE&LEFT(U_TIME, 4) & "' "
sQuery = "ORDER BY A.IDX DESC "

If IsTopN = "Y" Then
	StoreProcName = "USP_Admin_EShop_Event_Select_HistoryBack"
Else
	StoreProcName = "USP_Admin_EShop_Event_Select"
End If

SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = StoreProcName

		.Parameters.Append .CreateParameter("@Page", adInteger, adParamInput, , Page)
		.Parameters.Append .CreateParameter("@PageSize", adInteger, adParamInput, , PageSize)
		.Parameters.Append .CreateParameter("@wQuery", adVarChar, adParamInput, 1000, wQuery)
		.Parameters.Append .CreateParameter("@sQuery", adVarChar, adParamInput, 100, sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

RecCnt	 = oRs(0)
PageCnt	 = FIX((RecCnt+(PageSize-1))/PageSize)

Set oRs = oRs.NextRecordset
%>
                        <li class="thumbNail">
                            <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Event/EventAttend.asp')">
								<img src="/Upload/etc/attend_m.jpg" alt="출석체크">
							</a>
                        </li>
<%
Do While Not oRs.EOF
%>
                        <li class="thumbNail">
                            <a href="javascript:void(0)" onclick="pushHash();APP_GoUrl('/ASP/Event/EventView.asp?EventIDX=<%=oRs("IDX")%>')">
								<img src="<%=oRs("MobileListBanner")%>" alt="">
							</a>
                        </li>
<%
	oRs.MoveNext
Loop
oRs.Close

Response.Write "|||||" & RecCnt & "|||||" & PageCnt

SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>