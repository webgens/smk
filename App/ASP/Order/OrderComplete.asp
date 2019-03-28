<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'OrderComplete.asp - 주문완료
'Date		: 2018.12.30
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
PageCode2 = "02"
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

DIM OrderCode

DIM ProductName
DIM OrderCnt
DIM OrderPrice
DIM SalePrice
DIM UseCouponPrice
DIM UsePointPrice
DIM UseScashPrice
DIM DeliveryPrice
DIM PayType
DIM OrderName
DIM OrderTel
DIM OrderHp
DIM OrderEmail
DIM DelvType
DIM ShopNM
DIM ReceiveName
DIM ReceiveTel
DIM ReceiveHp
DIM ReceiveZipCode
DIM ReceiveAddr1
DIM ReceiveAddr2
DIM ReceiptFlag
DIM Memo
DIM OrderDate
DIM OrderTime
DIM LGD_FINANCENAME
DIM LGD_CARDINSTALLMONTH
DIM LGD_ACCOUNTNUM
DIM LGD_TELNO

Dim WiderTracking_ProductInfo
Dim GoogleTag_ProductInfo
Dim Tracking_ProductInfo
Dim Temp_Tracking_ProductInfo
Dim FaceBookTracking_ProductInfo
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

OrderCode		= sqlFilter(Request("OrderCode"))		


IF OrderCode = "" THEN
		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=주문정보가 잘못되었습니다.&Script=APP_TopGoUrl('/');"
		Response.End
END IF


SET oConn		= ConnectionOpen()							'# 커넥션 생성
SET oRs			= Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성




'-----------------------------------------------------------------------------------------------------------'
'주문 정보 검색 START
'-----------------------------------------------------------------------------------------------------------'
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Front_EShop_Order_Select_By_OrderCode"

		.Parameters.Append .CreateParameter("@OrderCode",		adVarChar,		adParamInput, 20,		OrderCode)
		.Parameters.Append .CreateParameter("@UserID",			adVarChar,		adParamInput, 20,		U_NUM)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing

IF NOT oRs.EOF THEN
		OrderName					 = oRs("OrderName")
		OrderHp						 = oRs("OrderHp")
		OrderEmail					 = oRs("OrderEmail")
ELSE
		oRs.Close :  SET oRs = Nothing : oConn.Close : SET oConn = Nothing

		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=잘못된 주문 정보입니다.&Script=APP_TopGoUrl('/');"
		Response.End
END IF
oRs.Close


'-----------------------------------------------------------------------------------------------------------'
'주문 정보 검색 START
'-----------------------------------------------------------------------------------------------------------'
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Front_EShop_Order_Select_For_OrderInfo"

		.Parameters.Append .CreateParameter("@OrderCode",	adVarChar, adParaminput,	20,		OrderCode)
		.Parameters.Append .CreateParameter("@UserID",		adVarChar, adParamInput,	20,		U_NUM)
		.Parameters.Append .CreateParameter("@OrderName",	adVarChar, adParamInput,	50,		OrderName)
		.Parameters.Append .CreateParameter("@OrderHp",		adVarChar, adParamInput,	20,		OrderHp)
		.Parameters.Append .CreateParameter("@OrderEmail",	adVarChar, adParamInput,	50,		OrderEmail)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing

IF NOT oRs.EOF THEN

		ProductName					 = oRs("ProductName")
		OrderCnt					 = oRs("OrderCnt")
		OrderPrice					 = oRs("OrderPrice")
		SalePrice					 = oRs("SalePrice")
		UseCouponPrice				 = oRs("UseCouponPrice")	
		UsePointPrice				 = oRs("UsePointPrice")	
		UseScashPrice				 = oRs("UseScashPrice")	
		DeliveryPrice				 = oRs("DeliveryPrice")	
		PayType						 = oRs("PayType")
		OrderName					 = oRs("OrderName")
		OrderTel					 = oRs("OrderTel")
		OrderHp						 = oRs("OrderHp")
		OrderEmail					 = oRs("OrderEmail")
		DelvType					 = oRs("DelvType")
		ShopNM						 = oRs("ShopNM")
		ReceiveName					 = oRs("ReceiveName")
		ReceiveTel					 = oRs("ReceiveTel")
		ReceiveHp					 = oRs("ReceiveHp")
		ReceiveZipCode				 = oRs("ReceiveZipCode")
		ReceiveAddr1				 = oRs("ReceiveAddr1")
		ReceiveAddr2				 = oRs("ReceiveAddr2")
		ReceiptFlag					 = oRs("ReceiptFlag")
		Memo						 = oRs("Memo")
		OrderDate					 = oRs("OrderDate")
		OrderTime					 = oRs("OrderTime")
		LGD_FINANCENAME				 = oRs("LGD_FINANCENAME")
		LGD_CARDINSTALLMONTH		 = oRs("LGD_CARDINSTALLMONTH")
		LGD_ACCOUNTNUM				 = oRs("LGD_ACCOUNTNUM")
		LGD_TELNO					 = oRs("LGD_TELNO")
ELSE
		oRs.Close : SET oRs = Nothing : oConn.Close : SET oConn = Nothing
	
		Response.Redirect "/ASP/Error/ErrorPopupNone.asp?Title=SHOEMARKER&Msg=잘못된 주문 정보입니다.&Script=APP_TopGoUrl('/');"
		Response.End
End IF
oRs.Close
'-----------------------------------------------------------------------------------------------------------'
'주문정보 검색 END
'-----------------------------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/INC/Header.asp" -->
<!-- #include virtual="/INC/TopSub.asp" -->


    <main id="container" class="container">
        <div class="content">
            <div class="order-complete">
                <p class="p1">주문완료</p>
                <span class="order-num">주문번호 : <%=OrderCode%></span>
                <p class="p2">주문이 정상적으로 완료되었습니다.<br>바로 발송준비하도록 하겠습니다.</p>
            </div>
            <div class="cart cart-show">
                <div class="complete-confirm">
                    <p class="cart-tit">결제 확인</p>
                    <div class="price">
                        <div class="price-info">
                            <div class="info-wrap">
                                <p class="price-tit">결제수단</p>
                                <p class="price-value"><%=GetPayType(PayType)%></p>
                            </div>
					<%IF PayType ="C" THEN  %>
                            <div class="info-wrap">
                                <p class="price-tit">결제카드</p>
                                <p class="price-value">
									<%=LGD_FINANCENAME%>카드
									<%IF LGD_CARDINSTALLMONTH = "00" THEN%>
										(일시불)
									<%ELSE%>
										(<%=FormatNumber(LGD_CARDINSTALLMONTH,0)%>개월 할부)
									<%END IF%>
                                </p>
                            </div>
					<%ELSEIF PayType ="B" THEN  %>
                            <div class="info-wrap">
                                <p class="price-tit">이체은행</p>
                                <p class="price-value"><%=LGD_FINANCENAME%>은행</p>
                            </div>
					<%ELSEIF PayType ="V" THEN  %>
                            <div class="info-wrap">
                                <p class="price-tit">입금은행</p>
                                <p class="price-value"><%=LGD_FINANCENAME%>은행</p>
                            </div>
                            <div class="info-wrap">
                                <p class="price-tit">입금계좌</p>
                                <p class="price-value"><%=LGD_ACCOUNTNUM%></p>
                            </div>
                            <div class="info-wrap">
                                <p class="price-tit">예금주</p>
                                <p class="price-value"><%=MALL_LGD_ACCOUNTOWNER%></p>
                            </div>
					<%ELSEIF PayType ="M" THEN  %>
                            <div class="info-wrap">
                                <p class="price-tit">통신사</p>
                                <p class="price-value"><%=LGD_FINANCENAME%></p>
                            </div>
                            <div class="info-wrap">
                                <p class="price-tit">휴대폰</p>
                                <p class="price-value"><%=LGD_TELNO%></p>
                            </div>
					<%END IF %>
                            <div class="info-wrap">
                                <p class="price-tit">결제금액</p>
                                <p class="price-value"><%=FormatNumber(OrderPrice,0)%>원</p>
                            </div>
                            <div class="info-wrap">
                                <p class="price-tit">결제일시</p>
                                <p class="price-value"><%=GetDateYMD(OrderDate) & " " & GetTimeHMS(OrderTime)%></p>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="complete-confirm">
                    <p class="cart-tit">배송지 확인</p>
                    <div class="price">
                        <div class="price-info">
                            <div class="info-wrap">
                                <p class="price-tit">받는 사람</p>
                                <p class="price-value"><%=ReceiveName%></p>
                            </div>
                            <div class="info-wrap">
                                <p class="price-tit">연락처</p>
                                <p class="price-value"><%=ReceiveHp%></p>
                            </div>
                            <div class="info-wrap">
                                <p class="price-tit">추가 연락처</p>
                                <p class="price-value"><%=ReceiveTel%></p>
                            </div>
                            <div class="info-wrap address-wrap">
                                <p class="price-tit">배송주소</p>
                                <div class="price-value address-result">
									<%IF DelvType = "S" THEN%>
									<p>슈마커 <%=ShopNM%></p>
									<%ELSE%>
                                    <p>(<%=ReceiveZipCode%>)<%=ReceiveAddr1%></p>
                                    <p><%=ReceiveAddr2%></p>
									<%END IF%>
                                </div>
                            </div>
                            <div class="info-wrap">
                                <p class="price-tit">배송 요청사항</p>
                                <p class="price-value require-result"><%=Memo%></p>
                            </div>
                        </div>
                    </div>
                </div>
                <div class="confirm-more">
                    <a href="javascript:void(0)" onclick="APP_TopGoUrl('/');" class="button is-expand ty-red">쇼핑 계속하기</a>
                    <a href="javascript:void(0)" onclick="location.href='/ASP/MyPage/OrderList.asp'" class="button is-expand">주문내역 상세보기</a>
                </div>
            </div>
        </div>
    </main>

<!-- #include virtual="/INC/Footer.asp" -->

<%
wQuery = "WHERE A.OrderCode = '" & OrderCode & "' AND A.ProductType = 'P' "
sQuery = "ORDER BY A.IDX ASC"
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Front_EShop_Order_Product_Select_For_Order_Detail"

		.Parameters.Append .CreateParameter("@wQuery",		adVarChar,		adParamInput, 1000,		wQuery)
		.Parameters.Append .CreateParameter("@sQuery",		adVarChar,		adParamInput, 100,		sQuery)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing

y = 0
Do While Not oRs.EOF
	
	If y = 0 Then
		WiderTracking_ProductInfo = "{i:""" & oRs("ProductCode") & """, t:""" & oRs("ProductName") & """, p:""" & oRs("SalePrice") & """, q:""1""}"
		Tracking_ProductInfo = oRs("SalePrice") & "|,|" & oRs("ProductCode") & "|,|" & oRs("ProductName") & "|,|" & oRs("BrandName")
		FaceBookTracking_ProductInfo = """" & oRs("ProductCode") & """"
	Else
		WiderTracking_ProductInfo = WiderTracking_ProductInfo & ", {i:""" & oRs("ProductCode") & """, t:""" & oRs("ProductName") & """, p:""" & oRs("SalePrice") & """, q:""1""}"
		Tracking_ProductInfo = Tracking_ProductInfo & "|||" & oRs("SalePrice") & "|,|" & oRs("ProductCode") & "|,|" & oRs("ProductName") & "|,|" & oRs("BrandName")
		FaceBookTracking_ProductInfo = FaceBookTracking_ProductInfo & "," & """" & oRs("ProductCode") & """"
	End If

	GoogleTag_ProductInfo = GoogleTag_ProductInfo & "brandIds.push('" & oRs("ProductCode") & "');" 

	y = y + 1

	oRs.MoveNext
Loop
oRs.Close
%>

<!-- WIDERPLANET PURCHASE SCRIPT START 2019.1.8 -->
<div id="wp_tg_cts" style="display:none;"></div>
<script type="text/javascript">
var wptg_tagscript_vars = wptg_tagscript_vars || [];
wptg_tagscript_vars.push(
(function() {
	return {
		wp_hcuid:"<%=U_Num%>",  	/*고객넘버 등 Unique ID (ex. 로그인  ID, 고객넘버 등 )를 암호화하여 대입.
				 *주의 : 로그인 하지 않은 사용자는 어떠한 값도 대입하지 않습니다.*/
		ti:"24585",
		ty:"PurchaseComplete",
		device:"mobile"
		,items:[
			 <%=WiderTracking_ProductInfo%>
		]
	};
}));
</script>
<script type="text/javascript" async src="//cdn-aitg.widerplanet.com/js/wp_astg_4.0.js"></script>
<!-- // WIDERPLANET PURCHASE SCRIPT END 2019.1.8 -->

<!-- Event snippet for MO 구매완료 conversion page -->
<script>
  gtag('event', 'conversion', {
      'send_to': 'AW-815695980/frd0CO2UwZIBEOyQ-oQD',
      'value': <%=OrderPrice%>,
      'currency': 'KRW',
      'transaction_id': ''
  });
</script>

<!-- 전환페이지 설정 -->
<script type="text/javascript" src="//wcs.naver.net/wcslog.js"></script> 
<script type="text/javascript">
	var _nasa = {};
	_nasa["cnv"] = wcs.cnv("1", "<%=OrderPrice%>"); // 전환유형, 전환가치 설정해야함. 설치매뉴얼 참고
</script>

<!-- Google Tag Manager Variable (eMnet) -->
<script type="text/javascript">
	var bprice = '<%=OrderPrice%>';
	var brandIds = [];
	<%=GoogleTag_ProductInfo%>
</script>
<!-- End Google Tag Manager Variable (eMnet) --> 

<%
	'0:금액, 1:코드, 2:상품명, 3:브랜드
	Temp_Tracking_ProductInfo = Split(Tracking_ProductInfo, "|||")
	For i = 0 To UBound(Temp_Tracking_ProductInfo)
		If Trim(Temp_Tracking_ProductInfo(i)) <> "" Then
			Tracking_ProductInfo = Split(Temp_Tracking_ProductInfo(i), "|,|")
%>
<!-- AceCounter Mobile eCommerce (Cart_Inout) v7.5 Start -->
<script type="text/javascript">
	var AM_Cart=(function(){
		var c={pd:'<%=Trim(Tracking_ProductInfo(1))%>',pn:'<%=Trim(Tracking_ProductInfo(2))%>',am:'<%=Trim(Tracking_ProductInfo(0))%>',qy:'1',ct:'<%=Trim(Tracking_ProductInfo(3))%>'};
		var u=(!AM_Cart)?[]:AM_Cart; u[c.pd]=c;return u;
	})();
</script>
<%
		End If
	Next
%>

<script type="text/javascript">
	var m_order_code='<%=OrderCode%>';		// 주문코드 필수 입력 
	var m_buy="finish"; //구매 완료 변수(finish 고정값)
</script>

<!-- Facebook Pixel Code -->
<script>
	fbq('track', 'Purchase', {
	content_type: 'product',
	content_ids: [<%=FaceBookTracking_ProductInfo%>],
	value: <%=OrderPrice%>,
	currency: 'KRW',
	});
</script>
<!-- End Facebook Pixel Code -->

<!-- adinsight 주문 총금액 받아옴. start -->
<script language='javascript'> 
	var TRS_AMT='<%=OrderPrice%>'; 
	var TRS_ORDER_ID='<%=OrderCode%>'; 
</script>
<!-- adinsight 주문 총금액 받아옴. end -->

<!-- GA -->
<script type="text/javascript">
	gtag('event', 'purchase', {
		"transaction_id": "<%=OrderCode%>",
		"affiliation": "슈마커",
		"value": <%=OrderPrice%>,
		"currency": "KRW",
		"tax": 0,
		"shipping": 0,
		"items": [
<%
	'0:금액, 1:코드, 2:상품명, 3:브랜드
For i = 0 To UBound(Temp_Tracking_ProductInfo)
	If Trim(Temp_Tracking_ProductInfo(i)) <> "" Then
	Tracking_ProductInfo = Split(Temp_Tracking_ProductInfo(i), "|,|")
		If i = 0 Then
%>
		{
			"id": "<%=Tracking_ProductInfo(1)%>",
			"name": "<%=Tracking_ProductInfo(2)%>",
			"list_name": "OrderComplete",
			"brand": "<%=Tracking_ProductInfo(3)%>",
			"category": "",
			"variant": "",
			"list_position": 1,
			"quantity": 1,
			"price": '<%=Tracking_ProductInfo(0)%>'
		}
<%
		Else
%>
		, {
		"id": "<%=Tracking_ProductInfo(1)%>",
		"name": "<%=Tracking_ProductInfo(2)%>",
		"list_name": "OrderComplete",
		"brand": "<%=Tracking_ProductInfo(3)%>",
		"category": "",
		"variant": "",
		"list_position": 1,
		"quantity": 1,
		"price": '<%=Tracking_ProductInfo(0)%>'
	}
<%
		End If
	End If
Next
%>
	]
	});
</script>
<!-- GA -->

<!-- kakao pixel script //-->
<script type="text/javascript" charset="UTF-8" src="//t1.daumcdn.net/adfit/static/kp.js"></script>
<script type="text/javascript">
	kakaoPixel('5354511058043421336').pageView();
	kakaoPixel('5354511058043421336').purchase({
		total_quantity: "<%=y%>", // 주문 내 상품 개수(optional)
		total_price: "<%=OrderPrice%>",  // 주문 총 가격(optional)
		currency: "KRW",     // 주문 가격의 화폐 단위(optional, 기본 값은 KRW)
		products: [          // 주문 내 상품 정보(optional)
<%
	'0:금액, 1:코드, 2:상품명, 3:브랜드
For i = 0 To UBound(Temp_Tracking_ProductInfo)
	If Trim(Temp_Tracking_ProductInfo(i)) <> "" Then
	Tracking_ProductInfo = Split(Temp_Tracking_ProductInfo(i), "|,|")
	If i = 0 Then
%>
			{ name: "<%=Tracking_ProductInfo(2)%>", quantity: "1", price: "<%=Tracking_ProductInfo(0)%>"}
<%
		Else
%>
			, { name: "<%=Tracking_ProductInfo(2)%>", quantity: "1", price: "<%=Tracking_ProductInfo(0)%>"}
<%
		End If
	End If
Next
%>           
		]
	});
</script>
<!-- kakao pixel script //-->

<!-- #include virtual="/INC/Bottom.asp" -->

<%
Response.Cookies("NON_ORDER") = ""

SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>