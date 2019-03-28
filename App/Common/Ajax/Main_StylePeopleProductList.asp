<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'Main_StylePeopleProductList.asp - 메인페이지 STYLE PEOPLE 상품 리스트
'Date		: 2018.12.24
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

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절

DIM Idx


DIM Title_06
DIM MobileImage1_06
DIM MobileImage2_06
DIM DisplayTitle1_06
DIM DisplayTitle2_06
DIM DisplayTitle3_06
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'

	
Idx				 = sqlFilter(Request("Idx"))

	

SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성





Response.Write "OK|||||"


wQuery = "WHERE IDX = " & Idx & " "
sQuery = "ORDER BY ReserveMainFlag DESC, DisplayNum ASC, Idx DESC "


SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_MainBanner_Select_For_Ing"

		.Parameters.Append .CreateParameter("@wQuery",		 adVarchar, adParaminput, 1000	, wQuery)
		.Parameters.Append .CreateParameter("@sQuery",		 adVarchar, adParaminput, 100	, sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing


If Not oRs.EOF Then
	Title_06 = REPLACE(oRs("Title"), """", "")
	MobileImage1_06 = oRs("MobileImage1")
	MobileImage2_06 = oRs("MobileImage2")
	DisplayTitle1_06 = oRs("DisplayTitle1")
	DisplayTitle2_06 = oRs("DisplayTitle2")
	DisplayTitle3_06 = oRs("DisplayTitle3")
End If
oRs.Close
%>
                                <div class="ly-img">
                                    <div class="img">
                                        <img src="<%=MobileImage2_06%>" alt="<%=Title_06%>">
                                        <!-- height 고정 320px-->
                                    </div>
                                </div>
                                <div class="subs">
                                    <div class="txt">
                                        <a href="javascript:openExternal('<%=DisplayTitle3_06%>');"><%=DisplayTitle1_06%></a>
                                        <span><%=DisplayTitle2_06%></span>
                                    </div>
                                    <ul class="contView">

<%

SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_MainBanner_Product_Select_Top3_By_MainBannerIdx"

		.Parameters.Append .CreateParameter("@MainBannerIdx", adInteger, adParamInput, 20, Idx)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing
																
IF NOT oRs.EOF THEN
		i = 1
		Do Until oRs.EOF
%>
                                        <li class="contItem">
                                            <a href="javascript:APP_GoUrl('/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>');">
                                                <span class="cont">
                                                    <span class="thumbNail">
                                                        <img src="<%=oRs("ImageUrl_0180")%>" alt="<%=oRs("ProductName")%>">
                                                    </span>

                                                    <span class="detail">
                                                        <span class="brand"><%=oRs("BrandName")%></span>
                                                        <span class="product-name pname"><%=oRs("ProductName")%></span>
                                                         <span class="price"><strong><%=FormatNumber(oRs("SalePrice"), 0)%></strong>원</span>
                                                    </span>
                                                </span>
                                            </a>
                                        </li>
<%
				oRs.MoveNext
				i = i + 1
		Loop
END IF
oRs.Close
%>
                                    </ul>
                                </div>

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>
