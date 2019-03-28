<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'index.asp - 마이페이지
'Date		: 2018.12.17
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
<!-- #include virtual="/Common/CheckID_Ajax.asp" -->

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


DIM mType						'# 회원정보 수정타입
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

mType			 = sqlFilter(request("mType"))

SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>


<!-- #include virtual="/INC/Header.asp" -->
<script type="text/javascript">
	/* 찜한상품 리스트 */
	function productPickList() {
		$.ajax({
			type		 : "post",
			url			 : "/ASP/Mypage/Ajax/Product_Pick_List.asp",
			async		 : false,
			data		 : "",
			dataType	 : "text",
			success		 : function (data) {
							var splitData	 = data.split("|||||");
							var result		 = splitData[0];
							var cont		 = splitData[1];

							if (result == "OK") {
								$("#productPickList").html(cont);
								return;
							}
							else if (result == "LOGIN") {
								common_msgPopOpen("", cont, "", "location.href='/ASP/Member/Login.asp'");
								return;
							}
							else {
								common_msgPopOpen("", cont);
								return;
							}
			},
			error		 : function (data) {
								alert(data.responseText)
								common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.");
			}
		});
	}

	/* 찜한상품 삭제 */
	function productPickDel(ProductCode) {
		//location.href = "/ASP/Mypage/Ajax/PwdModifyOk.asp?" + $("#chgPwdForm").serialize();
		//return;

		if (ProductCode == "") {
			common_msgPopOpen("", "해당 상품 정보가 없습니다.");
			return;
		}

		common_msgPopOpen("", "해당 상품을 삭제 하시겠습니까?", "refundAccountkDelOk('"+ ProductCode +"');", "", "C");
	}

	function productPickDelOk(ProductCode) {
		$.ajax({
			type		 : "post",
			url			 : "/ASP/Mypage/Ajax/Product_Pick_Delete.asp",
			async		 : false,
			data		 : "ProductCode="+ProductCode,
			dataType	 : "text",
			success		 : function (data) {
							var splitData	 = data.split("|||||");
							var result		 = splitData[0];
							var cont		 = splitData[1];


							if (result == "OK") {
								productPickList();
								return;
							}
							else if (result == "LOGIN") {
								common_msgPopOpen("", cont, "", "location.href='/ASP/Member/Login.asp'");
								return;
							}
							else {
								common_msgPopOpen("", cont);
								return;
							}
			},
			error		 : function (data) {
								//alert(data.responseText)
								common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.");
			}
		});
	}

	/* 찜한상품 전체 삭제 */
	function sel_productPickAllDel() {
		common_msgPopOpen("", "전체 상품을 삭제 하시겠습니까?", "sel_productPickAllDelOk();", "", "C");
	}

	function sel_productPickAllDelOk() {
		$.ajax({
			type		 : "post",
			url			 : "/ASP/Mypage/Ajax/Product_Pick_All_Delete.asp",
			async		 : false,
			data		 : "",
			dataType	 : "text",
			success		 : function (data) {
							var splitData	 = data.split("|||||");
							var result		 = splitData[0];
							var cont		 = splitData[1];


							if (result == "OK") {
								productPickList();
								return;
							}
							else if (result == "LOGIN") {
								common_msgPopOpen("", cont, "", "location.href='/ASP/Member/Login.asp'");
								return;
							}
							else {
								common_msgPopOpen("", cont);
								return;
							}
			},
			error		 : function (data) {
								//alert(data.responseText)
								common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.");
			}
		});
	}







	/* 최근 본 상품 리스트 */
	function productLatestList() {
		$.ajax({
			type		 : "post",
			url			 : "/ASP/Mypage/Ajax/Product_Latest_List.asp",
			async		 : false,
			data		 : "",
			dataType	 : "text",
			success		 : function (data) {
							var splitData	 = data.split("|||||");
							var result		 = splitData[0];
							var cont		 = splitData[1];


							if (result == "OK") {
								$("#productLatestList").html(cont);
								return;
							}
							else if (result == "LOGIN") {
								common_msgPopOpen("", cont, "", "location.href='/ASP/Member/Login.asp'");
								return;
							}
							else {
								common_msgPopOpen("", cont);
								return;
							}
			},
			error		 : function (data) {
								//alert(data.responseText)
								common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.");
			}
		});
	}

	/* 최근 본 상품 전체 삭제 */
	function sel_productLatestAllDel() {
		common_msgPopOpen("", "전체 상품을 삭제 하시겠습니까?", "sel_productLatestAllDelOk();", "", "C");
	}

	function sel_productLatestAllDelOk() {
		$.ajax({
			type		 : "post",
			url			 : "/ASP/Mypage/Ajax/Product_Latest_All_Delete.asp",
			async		 : false,
			data		 : "ProductCode="+checkArr,
			dataType	 : "text",
			success		 : function (data) {
							var splitData	 = data.split("|||||");
							var result		 = splitData[0];
							var cont		 = splitData[1];


							if (result == "OK") {
								productLatestList();
								return;
							}
							else if (result == "LOGIN") {
								common_msgPopOpen("", cont, "", "location.href='/ASP/Member/Login.asp'");
								return;
							}
							else {
								common_msgPopOpen("", cont);
								return;
							}
			},
			error		 : function (data) {
								//alert(data.responseText)
								common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.");
			}
		});
	}

	/* 최근 본 상품 찜하기 추가처리 */
	function chgPickAddOk(productCode) {
		if (productCode.length == 0) {
			common_msgPopOpen("", "찜한상품으로 추가 할 상품 정보가 없습니다.");
			return;
		}

		$.ajax({
			type		 : "post",
			url			 : "/ASP/Mypage/Ajax/Product_Pick_Insert.asp",
			async		 : false,
			data		 : "ProductCode="+productCode,
			dataType	 : "text",
			success		 : function (data) {
							var splitData	 = data.split("|||||");
							var result		 = splitData[0];
							var cont		 = splitData[1];


							if (result == "OK") {
								common_msgPopOpen("", "찜한상품에 추가 되었습니다.");
								productPickList();
								return;
							}
							else if (result == "LOGIN") {
								common_msgPopOpen("", cont, "", "location.href='/ASP/Member/Login.asp'");
								return;
							}
							else {
								common_msgPopOpen("", cont);
								return;
							}
			},
			error		 : function (data) {
								//alert(data.responseText)
								common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.");
			}
		});		
	}



	/* 브랜드 리스트 */
	function brandPickList() {
		$.ajax({
			type		 : "post",
			url			 : "/ASP/Mypage/Ajax/Brand_Pick_List.asp",
			async		 : false,
			data		 : "",
			dataType	 : "text",
			success		 : function (data) {
							var splitData	 = data.split("|||||");
							var result		 = splitData[0];
							var cont		 = splitData[1];


							if (result == "OK") {
								$("#brandPickList").html(cont);
								return;
							}
							else if (result == "LOGIN") {
								common_msgPopOpen("", cont, "", "location.href='/ASP/Member/Login.asp'");
								return;
							}
							else {
								common_msgPopOpen("", cont);
								return;
							}
			},
			error		 : function (data) {
								//alert(data.responseText)
								common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.");
			}
		});
	}

	/* 브랜드 전체 삭제 */
	function sel_brandPickAllDel() {
		common_msgPopOpen("", "전체 브랜드를 삭제 하시겠습니까?", "sel_brandPickAllDelOk();", "", "C");
	}

	function sel_brandPickAllDelOk() {
		$.ajax({
			type		 : "post",
			url			 : "/ASP/Mypage/Ajax/Brand_Pick_All_Delete.asp",
			async		 : false,
			data		 : "",
			dataType	 : "text",
			success		 : function (data) {
							var splitData	 = data.split("|||||");
							var result		 = splitData[0];
							var cont		 = splitData[1];


							if (result == "OK") {
								brandPickList();
								return;
							}
							else if (result == "LOGIN") {
								common_msgPopOpen("", cont, "", "location.href='/ASP/Member/Login.asp'");
								return;
							}
							else {
								common_msgPopOpen("", cont);
								return;
							}
			},
			error		 : function (data) {
								//alert(data.responseText)
								common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.");
			}
		});
	}

	/* 브랜드 추가팝업 */
	function brandPickAdd() {
		common_PopOpen("DimDepth1","BrandAdd");
	}

	/* 브랜드 추가처리 */
	function brandPickAddOk() {
		var bInt = $("form[name=BrandPickAddForm] input[name=add_idx]:checked").length;
		if (bInt == 0) {
			common_msgPopOpen("", "추가 할 브랜드 정보가 없습니다.");
			return;
		}

		var checkArr = [];
		$("form[name=BrandPickAddForm] input[name=add_idx]:checked").each(function () {
				checkArr.push($(this).val());
			}
		)

		$.ajax({
			type		 : "post",
			url			 : "/ASP/Mypage/Ajax/Brand_Pick_insert.asp",
			async		 : false,
			data		 : "BrandCode="+checkArr,
			dataType	 : "text",
			success		 : function (data) {
							var splitData	 = data.split("|||||");
							var result		 = splitData[0];
							var cont		 = splitData[1];


							if (result == "OK") {
								common_msgPopOpen("", "브랜드가 추가되었습니다.");
								common_PopClose('DimDepth1');
								brandPickList();
								return;
							}
							else if (result == "LOGIN") {
								common_msgPopOpen("", cont, "", "location.href='/ASP/Member/Login.asp'");
								return;
							}
							else {
								common_msgPopOpen("", cont);
								return;
							}
			},
			error		 : function (data) {
								//alert(data.responseText)
								common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.");
			}
		});		
	}
</script>
<!-- #include virtual="/INC/TopMain.asp" -->


    <!-- Main -->
    <main id="container" class="container">
        <div class="content">

            <div class="wrap-mypage">
				<div style="height:8px"></div>


				
                        <div id="OrderMenu" class="ly-title accordion">
                            <div class="selector">
	                            <button type="button" class="btn-list clickEvt" data-target="OrderMenu">MY&hearts;</button>
							</div>
							<div class="option my-recode">
								<ul>
									<li><a href="/ASP/Mypage/MyShoemarker.asp">MY&hearts;</a></li>
									<li><a href="#">재입고알림</a></li>
									<li><a href="#">상품평</a></li>
									<li><a href="#">상품문의</a></li>
								</ul>
							</div>
                        </div>


                <div class="mypage-my">
                    <div id="shoppingList_1" class="accord-mypage">

                        <div class="ly-content1">
                            <div id="tabs" class="tab">
                                <div class="tab-mypage">
                                    <ul class="tab-selector">
                                        <li class="part-3 "><a href="javascript:;" data-target="tabs-col1">찜한 상품</a></li>
                                        <!-- 탭메뉴 갯수에 따른 클래스명 지정 : part-탭메뉴 갯수 예) 탭메뉴가 3개일 때, part-3 -->
                                        <li class="part-3"><a href="javascript:;" data-target="tabs-col2">최근 본 상품</a></li>
                                        <li class="part-3"><a href="javascript:;" data-target="tabs-col3">MY브랜드</a></li>
                                    </ul>

                                    <!-- 찜한 상품 -->
                                    <div id="tabs-col1" class="tab-panel">
										<div id="productPickList"></div>
                                    </div>
                                    <!-- // 찜한 상품 -->

                                    <!-- 최근 본 상품 -->
                                    <div id="tabs-col2" class="tab-panel">
                                        <div class="h-line">
                                            <h2 class="h-level4">최근 본 상품 전체</h2>
                                            <span class="h-num">3건</span>
                                            <span class="h-date is-right">
                                                <button type="button" class="button-ty3 ty-bd-black">
                                                    <span>전체 삭제</span>
                                            </button>
                                            </span>
                                        </div>
                                        <ul>
                                            <li class="informItem">
                                                <a href="#">
                                                    <span class="cont">
                                                        <span class="thumbNail">
                                                            <span class="img">
                                                                <img src="/Images/tmp/@img_100_100_1.png" alt="상품 이미지">
                                                            </span>
                                                        </span>

                                                        <span class="detail">
                                                            <span class="brand">
                                                                <span class="name">FILA</span>
                                                                <span class="item-code">FS3SCA5331X-BLK FS3SCA5331X-BLK</span>
                                                            </span>
                                                            <span class="product-name"><em>FILA 앵클삭스 3족 세트 말줄임 처리 됩니다. </em></span>

                                                            <span class="product-more">
                                                                <span class="price"><strong>63,000</strong>원</span>
                                                                <span class="optional-info">
                                                                    <span class="icon ico-fav">140</span>
                                                                    <span class="icon ico-cmt">34</span>
                                                                </span>
                                                            </span>
                                                        </span>
                                                    </span>
                                                </a>

                                                <div class="buttongroup">
                                                    <button type="button" class="button-ty2 is-expand ty-bd-gray">상품 상세정보 보기</button>
                                                </div>
                                                <div class="right-circle">
                                                    <button type="button" class="closebtn">
                                                        <span class="hidden">닫기</span>
                                                    </button>
                                                </div>
                                            </li>
                                            <li class="informItem">
                                                <a href="#">
                                                    <span class="cont  ty-2">
                                                        <span class="thumbNail">
                                                            <span class="img">
                                                                <img src="/Images/tmp/@img_100_100_1.png" alt="상품 이미지">
                                                            </span>
                                                        </span>

                                                        <span class="detail">
                                                            <span class="brand">
                                                                <span class="name">FILA</span>
                                                                <span class="item-code">FS3SCA5331X-BLK FS3SCA5331X-BLK</span>
                                                            </span>
                                                            <span class="product-name"><em>FILA 앵클삭스 3족 세트 말줄임 처리 됩니다. </em></span>

                                                            <span class="product-more">
                                                                <span class="price"><strong>63,000</strong>원</span>
                                                                <span class="optional-info">
                                                                    <span class="icon ico-fav">140</span>
                                                                    <span class="icon ico-cmt">34</span>
                                                                </span>
                                                            </span>
                                                        </span>
                                                    </span>
                                                </a>

                                                <div class="buttongroup">
                                                    <button type="button" class="button-ty2 is-expand ty-bd-gray">상품 상세정보 보기</button>
                                                </div>
                                                <div class="right-circle">
                                                    <button type="button" class="closebtn">
                                                        <span class="hidden">닫기</span>
                                                    </button>
                                                </div>
                                            </li>
                                        </ul>
                                        <div class="inf-type1">
                                            <p class="tit">알려드립니다.</p>
                                            <ul>
                                                <li class="bullet-ty1">최근 본 상품을 기준으로 최대 30개까지 저장됩니다.</li>
                                            </ul>
                                        </div>
                                    </div>
                                    <!-- // 최근 본 상품 -->

                                    <!-- MY브랜드 -->
                                    <div id="tabs-col3" class="tab-panel">
                                        <div class="h-line">
                                            <h2 class="h-level4">찜한 브랜드 전체</h2>
                                            <span class="h-num">4건</span>
                                            <span class="h-date is-right">
                                                <button type="button" class="button-ty3 ty-bd-black">
                                                    <span>전체 삭제</span>
                                            </button>
                                            </span>
                                        </div>
                                        <div class="like-brand">
                                            <ul>
                                                <li class="brand-list">
                                                    <a href="">
                                                        <p>나이키</p>
                                                        <p href="" class="shortcut">브랜드샵 바로가기</p>
                                                    </a>
                                                    <div class="right-circle">
                                                        <button type="button" class="closebtn">
                                                            <span class="hidden">닫기</span>
                                                        </button>
                                                    </div>
                                                </li>
                                                <li class="brand-list">
                                                    <a href="">
                                                        <p>아디다스</p>
                                                        <p href="" class="shortcut">브랜드샵 바로가기</p>
                                                    </a>
                                                    <div class="right-circle">
                                                        <button type="button" class="closebtn">
                                                            <span class="hidden">닫기</span>
                                                        </button>
                                                    </div>
                                                </li>
                                                <li class="brand-list">
                                                    <a href="">
                                                        <p>마이애미스탠스</p>
                                                        <p href="" class="shortcut">브랜드샵 바로가기</p>
                                                    </a>
                                                    <div class="right-circle">
                                                        <button type="button" class="closebtn">
                                                            <span class="hidden">닫기</span>
                                                        </button>
                                                    </div>
                                                </li>
                                                <li class="brand-list">
                                                    <a href="">
                                                        <p>라코스테</p>
                                                        <p href="" class="shortcut">브랜드샵 바로가기</p>
                                                    </a>
                                                    <div class="right-circle">
                                                        <button type="button" class="closebtn">
                                                            <span class="hidden">닫기</span>
                                                        </button>
                                                    </div>
                                                </li>
                                            </ul>
                                        </div>
                                    </div>
                                    <!-- //MY브랜드 -->
                                </div>
                            </div>
                        </div>
                    </div>

                </div>
            </div>
        </div>
    </main>


<!-- #include virtual="/INC/Footer.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->
 <!-- script 공통 -->
 <script>
	/* 찜한 상품 */
	productPickList();
	/* 최근 본 상품 */
	//productLatestList();
	/* 브랜드 */
	//brandPickList();
 </script>


<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>