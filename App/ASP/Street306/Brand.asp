<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'Index.asp - Street306
'Date		: 2019.01.07
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
PageCode1 = "ST"
PageCode2 = "BR"
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

Dim MainBanner
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>

<!-- #include virtual="/INC/Header.asp" -->
<!-- #include virtual="/INC/Top_Street306.asp" -->

    <!-- Main -->
    <main id="container" class="container">
        <div class="content">
            <section class="wrap-street">
                <div class="item-bg">
                    <img src="/images/img/@street_brand_1.jpg" alt="Street306 BRANDS">
                    <p>BRANDS</p>
                </div>
<%
wQuery = "WHERE DelFlag = 'N' AND BCode = 'R' "
sQuery = "ORDER BY DisplayNum DESC "
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_Street306_Banner_Select_By_wQuery_For_Brand"
		.Parameters.Append .CreateParameter("@wQuery", adVarChar, adParamInput, 1000, wQuery)
		.Parameters.Append .CreateParameter("@sQuery", adVarChar, adParamInput, 100, sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing
	
If NOT oRs.EOF Then
%>
                <div class="list-ty-brand">
                    <ul>
					<%
					Do While Not oRs.EOF						
					%>
                        <li>
                            <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Product/Brand.asp?SBrandCode=<%=oRs("BrandCode")%>')">
								<span class="thumbNail">
									<img src="<%=oRs("MobileImage")%>" alt="<%=oRs("BrandName")%>">
								</span>
								<span class="tit"><%=oRs("BrandNameKor")%></span>
							</a>
                        </li>
					<%
						oRs.MoveNext
					Loop	
					%>
                    </ul>
                </div>
<%
End If
oRs.Close	
%>
            </section>
        </div>
    </main>





<!-- #include virtual="/INC/Footer.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>