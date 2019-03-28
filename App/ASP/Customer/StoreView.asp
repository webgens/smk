<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'StoreView.asp - 고객센터 > 전국매장안내 뷰
'Date		: 2019.01.07
'Update	: 
'*****************************************************************************************'
	
'# 페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
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

DIM ShopNM
DIM Addr
DIM Tel
DIM XPoint
DIM YPoint
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'


ShopNM			 = sqlFilter(Request("ShopNM"))
Addr			 = sqlFilter(Request("Addr"))
Tel				 = sqlFilter(Request("Tel"))
XPoint			 = sqlFilter(Request("XPoint"))
YPoint			 = sqlFilter(Request("YPoint"))


SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>



<!-- #include virtual="/INC/Header.asp" -->
	<script type="text/javascript" src="//dapi.kakao.com/v2/maps/sdk.js?appkey=<%=KAKAO_LOGIN_CLIENTID%>&libraries=services"></script>
	<script type="text/javascript">
		//카카오맵 처리
		function kakaoMap(XPoint, YPoint, ShopNm, ShopAddr){
			var mapContainer = document.getElementById('KakaoMap'), // 지도를 표시할 div 
				mapOption = { 
					center: new daum.maps.LatLng(YPoint, XPoint), // 지도의 중심좌표
					level: 3 // 지도의 확대 레벨
				};

			// 지도를 표시할 div와  지도 옵션으로  지도를 생성합니다
			var map = new daum.maps.Map(mapContainer, mapOption); 
      
			if(XPoint.length>0 && YPoint.length>0){

				// 마커의 이미지정보를 가지고 있는 마커이미지를 생성합니다
				var markerPosition = new daum.maps.LatLng(YPoint, XPoint); // 마커가 표시될 위치입니다

				// 마커를 생성합니다
				var marker = new daum.maps.Marker({
					position: markerPosition
				});

				// 마커가 지도 위에 표시되도록 설정합니다
				marker.setMap(map);  

			}else{

				// 주소-좌표 변환 객체를 생성합니다
				var geocoder = new daum.maps.services.Geocoder();

				// 주소로 좌표를 검색합니다
				geocoder.addressSearch(ShopAddr, function(result, status) {

					// 정상적으로 검색이 완료됐으면 
					if (status == daum.maps.services.Status.OK) {
						var coords = new daum.maps.LatLng(result[0].y, result[0].x);

						// 결과값으로 받은 위치를 마커로 표시합니다
						var marker = new daum.maps.Marker({
							map: map,
							position: coords
						});

						// 지도의 중심을 결과값으로 받은 위치로 이동시킵니다
						map.setCenter(coords);
					} 
				});   
			}

		}
	</script>
<!-- #include virtual="/INC/PopupTop.asp" -->

    <!-- PopUp -->
        <div class="area-dim"></div>

        <div class="area-pop">
            <div class="full">
                <div class="tit-pop">
                    <p class="tit" id="setShopNm"></p>
                    <button onclick="APP_PopupHistoryBack();" class="btn-hide-pop">닫기</button>
                </div>
                <div class="container-pop">
                    <div class="contents no-padding-top">
                        <div class="pop-customer pop-customer-store" style="height:100%;">
                            <div class="tit-area">
                                <p id="setShopAddr"></p>
                                <span id="setShopTel"></span>
                            </div>
                            <div class="map" id="KakaoMap" style="height:75%;">
                                지도 영역
                            </div>
                            <!--<button type="button" class="button-ty2 is-expand ty-bd-gray"><span>매장정보 공유하기</span></button>-->
                        </div>
                    </div>
                    <div class="btns">
                        <button type="button" onclick="APP_PopupHistoryBack();" class="button ty-red">닫기</button>
                    </div>
                </div>
            </div>
        </div>
    <!-- // PopUp -->

		<form name="StoreView" id="StoreView">
			<input type="hidden" name="ShopNM"	value="<%=ShopNM%>"	 />
			<input type="hidden" name="ADDR"	value="<%=Addr%>"	 />
			<input type="hidden" name="TEL"		value="<%=Tel%>"	 />
			<input type="hidden" name="XPoint"	value="<%=XPoint%>"	 />
			<input type="hidden" name="YPoint"	value="<%=YPoint%>"	 />
		</form>


		<script type="text/javascript">
			$(function () {
				var ShopNM	= $("form[name=StoreView] input[name=ShopNM]").val();
				var ADDR	= $("form[name=StoreView] input[name=ADDR]").val();
				var TEL		= $("form[name=StoreView] input[name=TEL]").val();
				var XPoint	= $("form[name=StoreView] input[name=XPoint]").val();
				var YPoint	= $("form[name=StoreView] input[name=YPoint]").val();
				$("#setShopNm").html(ShopNM);
				$("#setShopAddr").html(ADDR);
				$("#setShopTel").html(TEL);
				kakaoMap(XPoint, YPoint, ShopNM, ADDR);
			});
		</script>

<!-- #include virtual="/INC/FooterNoBNB.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>