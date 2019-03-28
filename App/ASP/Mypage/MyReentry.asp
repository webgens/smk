<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'MyReentry.asp - 재입고 알림
'Date		: 2019.01.06
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
PageCode1 = "05"
PageCode2 = "03"
PageCode3 = "00"
PageCode4 = "00"
'-----------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->
<!-- #include virtual="/Common/SubCheckID.asp" -->

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
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>


<!-- #include virtual="/INC/Header.asp" -->
	<script type="text/javascript">
		/* 알림신청 해제 */
		function reentryDel(idx) {
			common_msgPopOpen("", "해당 알림신청 상품을 해제하시겠습니까?", "reentryDel2('" + idx + "');", "msgPopup", "C");
		}
		function reentryDel2(idx) {
			$.ajax({
				type: "post",
				url: "/ASP/Mypage/Ajax/MyReentryDelOk.asp",
				async: true,
				data: "Idx=" + idx,
				dataType: "text",
				success: function (data) {
					var splitData = data.split("|||||");
					var result = splitData[0];
					var cont = splitData[1];

					if (result == "OK") {
						common_msgPopOpen("", cont, "PageReload()", "msgPopup", "N");
						return;
					}
					else if (result == "LOGIN") {
						PageReload();
						return;
					}
					else {
						common_msgPopOpen("", cont, "", "msgPopup", "N");
						return;
					}
				},
				error: function (data) {
					//alert(data.responseText);
					common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
				}
			});
		}

		/* 장바구니 담기 */
		function addCart(productCode, sizeCD) {

			$("form[name='OrderForm'] input[name='OrderType']").val("G");
			$("form[name='OrderForm'] input[name='DelvType']").val("P");
			$("form[name='OrderForm'] input[name='ProductCode']").val(productCode);
			$("form[name='OrderForm'] input[name='SizeCD']").val(sizeCD);
			$("form[name='OrderForm'] input[name='OrderCnt']").val("1");
			$("form[name='OrderForm'] input[name='SalePriceType']").val("1");
			$("form[name='OrderForm'] input[name='ProductType']").val("1");

			common_msgPopOpen("", "해당 상품을 장바구니에 담으시겠습니까?", "addCart2();", "msgPopup", "C");
		}
		function addCart2() {
			$.ajax({
				type: "post",
				url: "/ASP/Order/Ajax/CartProductAddOk.asp",
				async: false,
				data: $("form[name='OrderForm']").serialize(),
				dataType: "text",
				success: function (data) {
					var splitData = data.split("|||||");
					var result = splitData[0];
					var cont = splitData[1];

					if (result == "OK") {
						get_GNB_CartCount();
						alertPopup("AddCart");
						return;
					}
					else if (result == "LOGIN") {
						PageReload();
					}
					else {
						common_msgPopOpen("", cont, "", "msgPopup", "N");
						return;
					}
				},
				error: function (data) {
					alert(data.responseText);
					common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
				}
			});
		}
		// 알림창 띄우기
		function alertPopup(alertType) {
			$.ajax({
				type: "post",
				url: "/Common/Ajax/AlertPopup.asp",
				async: false,
				data: "AlertType=" + alertType,
				dataType: "text",
				success: function (data) {
					$("#msgPopup").html(data);
					openPop('msgPopup');
				},
				error: function (data) {
					alert(data.responseText);
					common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
				}
			});
		}
	</script>

<%TopSubMenuTitle = "MY슈마커"%>
<!-- #include virtual="/INC/TopSub.asp" -->


    <!-- Main -->
    <main id="container" class="container">
        <div class="sub_content">

            <div class="wrap-mypage">
				<div style="height:8px"></div>
				
                <div id="MypageSubMenu" class="ly-title accordion">
                    <div class="selector">
	                    <button type="button" class="btn-list clickEvt" data-target="MypageSubMenu">재입고 알림</button>
					</div>
					<div class="option my-recode">
						<!-- #include virtual="/ASP/Mypage/SubMenu_MyShoeMarker.asp" -->
					</div>
                </div>
<%
SET oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection	 = oConn
		.CommandType		 = adCmdStoredProc
		.CommandText		 = "USP_Front_EShop_Product_Reentry_Select_By_MemberNum"

		.Parameters.Append .CreateParameter("@MemberNum",	 adInteger, adParaminput, , U_NUM)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
SET oCmd = Nothing
%>
                <div class="my-re-entry">
                    <div id="shoppingList">
                        <div>
                            <div class="">
                                <div class="h-line">
                                    <h2 class="h-level4">재입고 알림 신청내역</h2>
                                    <span class="h-num"><%=oRs.RecordCount%>건</span>
                                    <span class="h-date is-right">
                                        <button type="button" onclick="popMtmQnaAdd()" class="button-ty3 ty-bd-black">
                                            <span class="icon ico-inquire">1:1 문의하기</span>
										</button>
                                    </span>
                                </div>
                            </div>
<%
IF NOT oRs.EOF THEN
%>
                            <ul class="informView">
<%
		i=1
		Do While Not oRs.EOF
%>
                                <li class="informItem">
                                    <a href="javascript:void(0)" onclick="APP_GoUrl('/ASP/Product/ProductDetail.asp?ProductCode=<%=oRs("ProductCode")%>')">
                                        <span class="head-tit">
                                            <span class="tit">신청일 : <%=Left(oRs("CreateDT"),10)%></span>
                                        </span>
                                        <span class="cont ty-1">
                                            <span class="thumbNail">
                                                <span class="img">
                                                    <img src="<%=oRs("ImageUrl_180")%>" alt="상품 이미지">
                                                </span>
                                                <span class="about">
                                                    <span class="process"><%IF oRs("RestQty") > 0 THEN%>입고완료<%ELSE%>입고예정<%END IF%></span>
                                                </span>
                                            </span>

                                            <span class="detail">
                                                <span class="brand">
                                                    <span class="name"><%=oRs("BrandName")%></span>
                                                </span>
                                                <span class="product-name"><em><%=oRs("ProductName")%></em></span>
                                                <span class="inform">
                                                    <span class="list">
                                                        <span class="tit">옵션</span>
                                                        <span class="opt"><%=oRs("SizeCD")%></span>
                                                    </span>
                                                </span>
                                                <span class="re-entry-price">
                                                    <span><span class="bold"><%=FormatNumber(oRs("SalePrice"),0)%></span>원</span>
                                                </span>
                                            </span>
                                        </span>
                                    </a>

                                    <div class="buttongroup is-space">
										<%IF oRs("RestQty") > 0 THEN%>
                                        <button type="button" onclick="addCart('<%=oRs("ProductCode")%>', '<%=oRs("SizeCD")%>')" class="button-ty2 is-expand ty-black">장바구니 담기</button>
										<%ELSE%>
                                        <button type="button" class="button-ty2 is-expand ty-gray">장바구니 담기</button>
										<%END IF%>
                                        <button type="button" onclick="reentryDel('<%=oRs("Idx")%>')" class="button-ty2 is-expand ty-bd-grey">알림신청 해제</button>
                                    </div>
                                </li>
<%
		i = i + 1
		oRs.MoveNext
	Loop
%>
                            </ul>
<%
Else
%>
							<div class="area-empty">
								<span class="icon-empty"></span>
								<p class="tit-empty">재입고 신청한 상품이 없습니다.</p>
							</div>
<%
END IF
%>
                            <div class="inf-type1" style="padding-bottom:20px">
                                <p class="tit">알려드립니다.</p>
                                <ul>
                                    <li class="bullet-ty1">품절 상품 중 재입고가 될 경우 신청하신 SMS로 알려드립니다.</li>
									<li class="bullet-ty1">재입고 알림신청은 최대 10개까지 가능합니다.</li>
                                    <li class="bullet-ty1">상품의 입고 시간에 따라 오전 10시 30분, 오후1시, 오후5시 알림 문자가 발송됩니다.</li>
                                    <li class="bullet-ty1">재입고 알림 문자는 1회 발송됩니다.</li>
                                    <li class="bullet-ty1">재입고 알림은 신청자순으로 발송되며, 재입고 알림 후 인기 상품은 조기 품절될 수 있습니다.</li>
                                    <li class="bullet-ty1">신청일 기준 90일까지 저장됩니다.</li>
                                </ul>
                            </div>
                        </div>	<!-- -->
                    </div>	<!--shoppingList-->
                </div>	<!--my-re-entry-->
            </div>	<!--wrap-mypage-->
        </div> <!--content-->
    </main>

	<form name="OrderForm" method="post">
		<input type="hidden" name="OrderType"		value="" />
		<input type="hidden" name="DelvType"		value="" />
		<input type="hidden" name="ProductCode"		value="" />
		<input type="hidden" name="SizeCD"			value="" />
		<input type="hidden" name="OrderCnt"		value="1" />
		<input type="hidden" name="SalePriceType"	value="" />
		<input type="hidden" name="ProductType"		value="" />
	</form>

<!-- #include virtual="/INC/FooterNoBNB.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>