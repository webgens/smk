<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'StoreList.asp - 고객센터 > 전국매장안내 리스트
'Date		: 2019.01.07
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
DIM wQuery1						'# WHERE 절
DIM sQuery						'# SORT 절


DIM Page	 : Page = 1
DIM PageSize
DIM RecCnt
DIM PageCnt

Dim TotalCount
Dim LoadCount
Dim OutletCount
Dim MartCount
Dim DepartmentCount
Dim JoinCount

DIM ShopNM
DIM ADDR
DIM TEL
DIM XPoint
DIM YPoint

DIM sChannelNM
DIM sKeyword
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


sChannelNM		 = sqlFilter(Request("ChannelNM"))
sKeyword		 = sqlFilter(Request("Keyword"))
PageSize		 = sqlFilter(Request("PageSize"))
IF PageSize		 = "" THEN PageSize = 5


SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성

wQuery = "WHERE A.DisplayFlag = 'Y' AND A.UseFlag='Y' AND A.Circulation IN ('01', '02') AND A.ChannelNM IN ('로드샵','아울렛','마트','백화점') AND A.OpenDT <= '" & U_DATE & "' "
wQuery1 = ""
IF sKeyword <> "" THEN
	wQuery = wQuery & "AND A.ShopNM LIKE '%"& sKeyword &"%' "
END IF
IF sChannelNM = "가맹점" THEN
	wQuery1 = wQuery1 & "AND A.Circulation = '02' "
END IF
IF sChannelNM <> "가맹점" AND sChannelNM <> "" THEN
	wQuery1 = wQuery1 & "AND A.ChannelNM = '"& sChannelNM &"' "
END IF

sQuery = "ORDER BY A.IDX DESC "


Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Front_EShop_Store_Select_For_Count"

		.Parameters.Append .CreateParameter("@WQUERY",		adVarchar,	adParamInput, 1000,		wQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing

IF NOT oRs.EOF THEN
	TotalCount		= oRs("TotalCount")
	LoadCount		= oRs("LoadCount")
	OutletCount		= oRs("OutletCount")
	MartCount		= oRs("MartCount")
	DepartmentCount	= oRs("DepartmentCount")
	JoinCount		= oRs("JoinCount")
END IF
oRs.Close

Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Front_EShop_Store_Select"

		.Parameters.Append .CreateParameter("@PAGE",		adInteger,	adParamInput,	  ,		Page)
		.Parameters.Append .CreateParameter("@PAGE_SIZE",	adInteger,	adParamInput,	  ,		PageSize)
		.Parameters.Append .CreateParameter("@WQUERY",		adVarchar,	adParamInput, 1000,		wQuery)
		.Parameters.Append .CreateParameter("@WQUERY1",		adVarchar,	adParamInput, 1000,		wQuery1)
		.Parameters.Append .CreateParameter("@SQUERY",		adVarchar,	adParamInput,  100,		sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing


RecCnt	 = oRs(0)
PageCnt	 = FIX((RecCnt+(PageSize-1))/PageSize)

SET oRs = oRs.NextrecordSet

Response.Write "OK|||||"
%>
						<form name="StoreForm" id="StoreForm" method="post" action="javascript:searchStore('');">
                        <input type="hidden" name="ChannelNM" id="ChannelNM" value="<%=sChannelNM%>" />
                        <section class="FAQ-search">
                            <div class="fieldset">
                                <label class="fieldset-label">원하시는 지역의 슈마커 매장정보를 확인하세요.</label>
                                <div class="fieldset-row">
                                    <span class="input is-expand">
                                        <input type="text" name="Keyword" id="faq_enter_txt" title="검색어 입력" placeholder="검색어 입력" value="<%=sKeyword%>">
                                    </span>
                                </div>
                            </div>
                            <button type="button" class="button-ty2 is-expand ty-bd-gray" onclick="searchStore('');" tabindex="0">검색</button>
                        </section>
						</form>
                        <section class="store-type">
                            <div class="h-line">
                                <h2 class="h-level4"><%IF sKeyword<>"" THEN Response.Write "검색 &lsquo;"& sKeyword &"&rsquo; " %> <%=sChannelNM%></h2>
                                <span id="TotalCount"></span>
                            </div>
                            <div class="customer-select">
                                <ul class="select-list">
                                    <li <%IF sChannelNM = ""		THEN Response.Write "class='active'"%>><a href="javascript:searchStore('');">전체(<%=TotalCount%>)</a></li>
                                    <li <%IF sChannelNM = "로드샵"	THEN Response.Write "class='active'"%>><a href="javascript:searchStore('로드샵');">로드샵(<%=LoadCount%>)</a></li>
                                    <li <%IF sChannelNM = "아울렛"	THEN Response.Write "class='active'"%>><a href="javascript:searchStore('아울렛');">아울렛(<%=OutletCount%>)</a></li>
                                    <li <%IF sChannelNM = "마트"	THEN Response.Write "class='active'"%>><a href="javascript:searchStore('마트');">마트(<%=MartCount%>)</a></li>
                                    <li <%IF sChannelNM = "백화점"	THEN Response.Write "class='active'"%>><a href="javascript:searchStore('백화점');">백화점(<%=DepartmentCount%>)</a></li>
                                    <li <%IF sChannelNM = "가맹점"	THEN Response.Write "class='active'"%>><a href="javascript:searchStore('가맹점');">가맹점(<%=JoinCount%>)</a></li>
                                </ul>
                            </div>
                        </section>

                        <section class="store-list">
							<ul>
<%
IF oRs.EOF THEN
%>
                                <li>
									<p class="tit" style="text-align:center;">등록된 매장 정보가 없습니다.</p>
                                </li>
<%
ELSE
	j = 1
	DO UNTIL oRs.EOF
		ShopNM		= oRs("ShopNM")
		ADDR		= oRs("ADDR")
		TEL			= oRs("TEL")
		XPoint		= oRs("XPoint")
		YPoint		= oRs("YPoint")
%>

                                <li onclick="storeView('<%=XPoint%>', '<%=YPoint%>', '<%=ShopNM%>', '<%=ADDR%>', '<%=TEL%>')">
                                    <a class="right-arrow-bg">
                                        <p class="tit"><%=ShopNM%></p>
                                        <span class="cnt">
                                            <span class="address"><%=ADDR%></span>
		                                    <span><%=TEL%></span>
										</span>
                                    </a>
                                </li>
<%
	j = j + 1
	oRs.MoveNext
	LOOP
END IF
%>
                            </ul>
                        </section>

                        <div class="customer-btn-more" id="customer-btn-more">
                            <button type="button" onclick="storeList(<%=PageSize+5%>);" class="button-ty2 is-expand ty-bd-gray" style="display:none;">더보기</button>
                        </div>
						<input type="hidden" name="RecCnt" id="RecCnt" value="<%=RecCnt%>" />
						<input type="hidden" name="PageSize" id="PageSize" value="<%=PageSize%>" />

						<form name="StoreView" id="StoreView">
							<input type="hidden" name="ShopNM"	value="" />
							<input type="hidden" name="ADDR"	value="" />
							<input type="hidden" name="TEL"		value="" />
							<input type="hidden" name="XPoint"	value="" />
							<input type="hidden" name="YPoint"	value="" />
						</form>


<%
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>