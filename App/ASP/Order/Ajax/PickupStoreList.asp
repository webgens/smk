<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'PickupStoreList.asp - 픽업매장 리스트 가져오기 페이지
'Date		: 2018.12.30
'Update	: 
'/****************************************************************************************'

'//페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'//---------------------------------------------------------------------------------------'
%>

<!-- #include Virtual = "/ADO/ADODBCommon.asp" -->
<!-- #include Virtual = "/Common/Common.asp" -->
<!-- #include Virtual = "/Common/SetInfo.asp" -->

<%
'/****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn											'# ADODB Connection 개체
DIM oRs												'# ADODB Recordset 개체
DIM oCmd											'# ADODB Command 개체

DIM i
DIM j
DiM x
DIM y
DIM z

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절

DIM Page											'# 페이지 넘버
DIM PageSize : PageSize = 100						'# 페이지 사이즈
DIM RecCnt											'# 전체 레코드 카운트
DIM PageCnt											'# 페이지 카운트


DIM ProductCode
DIM SizeCD
DIM Sido
DIM Gugun
DIM StoreName
DIM Values											'# 변수값들
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'

ProductCode		= sqlFilter(Request("ProductCode"))
SizeCD			= sqlFilter(Request("SizeCD"))
Sido			= sqlFilter(Request("Sido"))
Gugun			= sqlFilter(Request("Gugun"))
StoreName		= sqlFilter(Request("StoreName"))

Page			= sqlFilter(Request("Page"))
IF Page			= "" THEN Page	= 1

Values			= ""
Values			= Values & "ProductCode="	& ProductCode
Values			= Values & "&SizeCD="		& SizeCD
Values			= Values & "&Sido="			& Sido
Values			= Values & "&Gugun="		& Gugun
Values			= Values & "&StoreName="	& StoreName


SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성



Response.Write "OK|||||"
%>					
<%
wQuery = ""
wQuery = wQuery & "WHERE A.ProductCode = " & ProductCode & " "
wQuery = wQuery & "AND   A.SizeCD = '" & SizeCD & "' "
wQuery = wQuery & "AND   A.UseFlag = 'Y' "
wQuery = wQuery & "AND   A.RestQty > 0 "
wQuery = wQuery & "AND   E.ShopCD IS NULL "
IF Sido <> "" THEN
		wQuery = wQuery & "AND   S.AreaNM = '" & Sido & "' "
END IF
IF Gugun <> "" THEN
		wQuery = wQuery & "AND   S.Addr1 LIKE '%" & Gugun & "%' "
END IF
IF StoreName <> "" THEN
		wQuery = wQuery & "AND   S.ShopNM LIKE '%" & StoreName & "%' "
END IF

sQuery = "ORDER BY S.AreaNM, S.ShopNM"

Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Front_EShop_Stock_Select_For_PickupStoreList"

		.Parameters.Append .CreateParameter("@PAGE",		adInteger,	adParamInput, ,			Page)
		.Parameters.Append .CreateParameter("@PAGE_SIZE",	adInteger,	adParamInput, ,			PageSize)
		.Parameters.Append .CreateParameter("@WQUERY",		adVarChar,	adParamInput, 1000,		wQuery)
		.Parameters.Append .CreateParameter("@SQUERY",		adVarChar,	adParamInput,  100,		SQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing

RecCnt	 = oRs(0)
PageCnt	 = FIX((RecCnt+(PageSize-1))/PageSize)


Set oRs = oRs.NextrecordSet

i = 0			
IF NOT oRs.EOF THEN
		Do Until oRs.EOF
%>
                                <li class="current">
                                    <span class="sel-store">
										<input type="hidden" name="StoreNM_<%=oRs("ShopCD")%>" value="<%=oRs("ShopNM")%>" />
										<input type="hidden" name="StoreTel_<%=oRs("ShopCD")%>" value="<%=oRs("ShopTel")%>" />
										<input type="hidden" name="StoreAddr1_<%=oRs("ShopCD")%>" value="<%=oRs("ShopAddr1")%>" />
										<input type="hidden" name="StoreAddr2_<%=oRs("ShopCD")%>" value="<%=oRs("ShopAddr2")%>" />
										<input type="hidden" name="StoreXPoint_<%=oRs("ShopCD")%>" value="<%=oRs("ShopXPoint")%>" />
										<input type="hidden" name="StoreYPoint_<%=oRs("ShopCD")%>" value="<%=oRs("ShopYPoint")%>" />
										<input type="radio" name="StoreCode" id="StoreCode_<%=oRs("ShopCD")%>" value="<%=oRs("ShopCD")%>" onclick="load_Map(<%=oRs("ShopXPoint")%>, <%=oRs("ShopYPoint")%>, '<%=oRs("ShopNM")%>')" />
										<label for="StoreCode_<%=oRs("ShopCD")%>">
											<span class="cont"><span class="name"><%=oRs("ShopNM")%></span><address class="addr"><%=oRs("ShopAddr1")%> <%=oRs("ShopAddr2")%></address> </span>
										</label>
                                    </span>
                                </li>
<%
				oRs.MoveNext
				i = i + 1
		Loop
ELSE
%>
								<li style="padding:14px 0; border-bottom:none;">
									<div class="area-empty" style="height:100%; padding:131px 0;">
										<span class="icon-empty"></span>
										<p class="tit-empty">검색된 매장이 없습니다</p>
									</div>
								</li>
<%
END IF
oRs.Close


Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>