<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'EventView.asp - 이벤트 내용
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
PageCode1 = "EV"
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

Dim EventIDX

Dim EType
Dim Title
Dim SDate
Dim EDate
Dim EHour
Dim Banner
Dim MobileBanner
Dim ListBanner
Dim MobileListBanner
Dim Contents
Dim MContents
Dim DisplayFlag
Dim DelFlag

Dim CategoryIDX
Dim CategoryCount

Dim nCategoryView
Dim LinkUrl
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'
EventIDX = sqlFilter(Request("EventIDX"))
CategoryIDX = sqlFilter(Request("CategoryIDX"))

'If U_ID = "jjang2121" Then
'	Response.Redirect "/ASP/Event/EventView_Como_1901.asp?EventIDX=55"
'End If	


If EventIDX = "" Then
		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=잘못된 경로 입니다.&Script=APP_HistoryBack();"
		Response.End
End If

SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성

SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Admin_EShop_Event_Select_By_IDX"
		.Parameters.Append .CreateParameter("@IDX", adInteger, adParamInput, , EventIDX)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

If oRs.EOF Then
		oRs.Close
		Set oRs = Nothing
		oConn.Close
		Set oConn = Nothing
	
		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=해당 이벤트가 존재하지 않습니다.&Script=APP_HistoryBack();"
		Response.End
Else
		EType = oRs("EType")
		Title = oRs("Title")
		SDate = oRs("SDate")
		EDate = oRs("EDate")
		Banner = oRs("Banner")
		MobileBanner = oRs("MobileBanner")
		ListBanner = oRs("ListBanner")
		MobileListBanner = oRs("MobileListBanner")
		Contents = oRs("Contents")
		MContents = oRs("MContents")
		DisplayFlag = oRs("DisplayFlag")
		DelFlag = oRs("DelFlag")
		LinkUrl = oRs("LinkUrl")
End If
oRs.Close

If LinkUrl <> "" Then Response.Redirect LinkUrl

If EType <> "E"  Or DisplayFlag = "N" Or DelFlag = "Y" Then
		Set oRs = Nothing
		oConn.Close
		Set oConn = Nothing
	
		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=해당 이벤트가 존재하지 않습니다.&Script=APP_HistoryBack();"
		Response.End
End If
	
'#If SDate > U_DATE&Left(U_Time, 4) Or EDate <= U_DATE&Left(U_Time, 4) Or DelFlag = "Y" Then
'#		Set oRs = Nothing
'#		oConn.Close
'#		Set oConn = Nothing
'#	
'#		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=종료된 이벤트 입니다.&Script=APP_HistoryBack();"
'#		Response.End
'#End If

'이벤트 ReadCnt Update
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Event_Update_For_ReadCnt"

		.Parameters.Append .CreateParameter("@IDX",			adInteger,	adParamInput,     ,	 EventIDX)
		.Execute, , adExecuteNoRecords
END WITH
SET oCmd = Nothing




DIM cRsCnt : cRsCnt = 0
DIM cArrRs
DIM EventCategory : EventCategory = ""

SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Event_Category_Select_By_EventIDX"

		.Parameters.Append .CreateParameter("@IDX", adInteger, adParamInput, , EventIDX)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing

CategoryCount = oRs.RecordCount

REDIM CategoryIdx(CategoryCount)
REDIM CategoryName(CategoryCount)
REDIM ProductCount(CategoryCount)
IF NOT oRs.EOF THEN

		EventCategory = EventCategory & "                        <div class=""inner-cont"">"&vbLf
		EventCategory = EventCategory & "                            <div class=""sort-ty2"">"&vbLf

		i = 1
		Do Until oRs.EOF
				CategoryIdx(i)	 = oRs("IDX")
				CategoryName(i)	 = oRs("CateName")
				ProductCount(i)	 = oRs("ProductCount")

				EventCategory	 = EventCategory & "								<button type=""button"" class=""length3 cate" & i & """ onclick=""move_Category('" & oRs("IDX") & "')"">" & oRs("CateName") & "</button>"&vbLf
				oRs.MoveNext
				i = i + 1
		Loop

		'IF ((i - 1) Mod 3) = 1 THEN
		'		EventCategory = EventCategory & "								<button type=""button"" class=""length3"">&nbsp;</button><button type=""button"" class=""length3"">&nbsp;</button>"&vbLf
		'ELSEIF ((i - 1) Mod 3) = 2 THEN
		'		EventCategory = EventCategory & "								<button type=""button"" class=""length3"">&nbsp;</button>"&vbLf
		'END IF
	
		EventCategory = EventCategory & "                            </div>"&vbLf
		EventCategory = EventCategory & "                        </div>"&vbLf

END IF
oRs.Close
%>

<!-- #include virtual="/INC/Header.asp" -->
	<script type="text/javascript">
		function CategoryView(eventIDX, categoryIDX, i, t) {
			for (x = 1; x <= t; x++) {
				if (x == i) {
					$("#Category" + x).addClass('active');
				} else {
					$("#Category" + x).removeClass('active');
				}
			}

			$.ajax({
				type: "post",
				url: "/ASP/Event/Ajax/EventProductList.asp",
				async: false,
				data: "EventIDX="+eventIDX+"&CategoryIDX="+categoryIDX,
				success: function (data) {
					arrData = data.split("|||||");
					CateName = arrData[0];
					Data = arrData[1];

					$("#CateName").html(CateName);
					$("#ProductList").html(Data);
				},
				error: function (data) {
					alert(data.responseText);
					alert("처리 도중 오류가 발생하였습니다.");
				}
			});

		}

		var dUrl = document.URL;
		if (dUrl.indexOf("#")) {
			dUrl = dUrl.substring(0, dUrl.indexOf("#"));
		}
		function move_Category(nm) {
			location.href = dUrl + "#" + nm;
		}
	</script>

<%TopSubMenuTitle = "이벤트"%>
<!-- #include virtual="/INC/TopSub.asp" -->




    <main id="container" class="container">
        <div class="sub_content">
            <section class="wrap-event">
                <div class="wrap-event-view" style="padding: 0px 0px 12px 0;">
                    <div class="img">
                        <% If MobileBanner <> "" Then %><img src="<%=MobileBanner%>" alt=""><% End If %>
						<% If MContents <> "" Then %>
							<%=MContents%>
						<% End If %>
                    </div>

<% 
IF EventCategory <> "" THEN
%>
                    <div class="event-item-list">
<%
		FOR i = 1 TO CategoryCount
				SET oCmd = Server.CreateObject("ADODB.Command")
				WITH oCmd
						.ActiveConnection	 = oConn
						.CommandType		 = adCmdStoredProc
						.CommandText		 = "USP_Mobile_EShop_Event_Category_Product_Select_By_EventIDX_CategoryIDX"

						.Parameters.Append .CreateParameter("@EventIDX",	 adInteger, adParamInput, , EventIDX)
						.Parameters.Append .CreateParameter("@CategoryIDX",	 adInteger, adParamInput, , CategoryIdx(i))
				END WITH
				oRs.CursorLocation = adUseClient
				oRs.Open oCmd, , adOpenStatic, adLockReadOnly
				SET oCmd = Nothing

				IF NOT oRs.EOF THEN
%> 
						<div style="position:relative">
						<a name="<%=CategoryIdx(i)%>" style="position:absolute;top:-40px"></a>
						<%IF i = 1 THEN%>
						<%'=Replace(EventCategory, "class=""length3 cate" & i & """", "class=""length3 cate" & i & " active""")%>
						<%=Replace(EventCategory, "class=""length3 cate" & i & """", "class=""length3 cate" & i & """")%>
						<%END IF%>

                        <div class="cont">
                            <div class="tit" id="CateName"><%=CategoryName(i)%> (<%=FormatNumber(oRs.RecordCount, 0)%>)</div>
                            <div class="item-list">
                                <ul class="listview" id="ProductList">
<%
						Do Until oRs.EOF	
%>
									<li>
										<a href="javascript:void(0)" class="listitems" onclick="APP_GoUrl('/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>')">
											<div class="badgegroup">
												<%=ProductBadge(oRs("ProductCode"), oRs("DiscountRate"), oRs("ReserveFlag"), oRs("OPOFlag"), oRs("PickupFlag"), oRs("GiftCnt"))%>
											</div>
											<div class="thumbnail"><img src="<%=oRs("ImageUrl")%>" alt="<%=oRs("ProductName")%>"></div>
											<p class="brand-name"><%=oRs("BrandName")%></p>
											<h1 class="product-name pname"><%=oRs("ProductName")%></h1>
											<p class="price"><strong><%=FormatNumber(oRs("SalePrice"), 0)%></strong>원</p>
										</a>
										<a nohref class="listitems">
											<p class="optional-info">
												<button type="button" class="btn-size" onclick="SizeLayerOpen('<%=oRs("ProductCode")%>');">SIZE</button>
												<span class="icon ico-fav"><%=FormatNumber(oRs("WishCnt"), 0)%></span>
												<span class="icon ico-cmt"><%=FormatNumber(oRs("ReviewCnt"), 0)%></span>
											</p>
										</a>
									</li>
<%
								oRs.MoveNext
						Loop
%>


                                </ul>
                            </div>
<%
				END IF
				oRs.Close
%>
                        </div>
						</div>
<%
		NEXT	
%>
                    </div>
<%
END IF
%>
                </div>
            </section>
        </div>
    </main>


	<form name="form" method="get">
		<input type="hidden" name="Page" />
		<input type="hidden" name="ISTopN" />
	</form>


<!-- #include virtual="/INC/FooterNoBNB.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>
