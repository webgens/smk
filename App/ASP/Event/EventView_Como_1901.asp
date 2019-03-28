<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'EventView_Como_1901.asp - 별도 추가페이지 이벤트 상세
'Date		: 2019.01.30
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
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'
EventIDX = sqlFilter(Request("EventIDX"))
CategoryIDX = sqlFilter(Request("CategoryIDX"))


If EventIDX = "" Then
	Call AlertMessage2("잘못된 경로 입니다.", "history.back();")
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
	
	Call AlertMessage2("정보가 없습니다.", "history.back();")
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
End If
oRs.Close

If EType <> "E" Then
	Set oRs = Nothing
	oConn.Close
	Set oConn = Nothing
	
	Call AlertMessage2("정보가 없습니다.", "history.back();")
	Response.End
End If
	
If SDate > U_DATE&Left(U_Time, 4) Or EDate <= U_DATE&Left(U_Time, 4) Or DelFlag = "Y" Then
	Set oRs = Nothing
	oConn.Close
	Set oConn = Nothing
	
	Call AlertMessage2("종료된 이벤트 입니다.", "history.back();")
	Response.End
End If

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
				type		 : "post",
				url			 : "/ASP/Event/Ajax/EventProductList.asp",
				async		 : false,
				data		 : "EventIDX="+eventIDX+"&CategoryIDX="+categoryIDX,
				success		 : function (data) {
								arrData	 = data.split("|||||");
								CateName = arrData[0];
								Data	 = arrData[1];

								$("#CateName").html(CateName);
								$("#ProductList").html(Data);
				},
				error		 : function (data) {
								alert(data.responseText);
								openAlertLayer("alert", "처리중 오류가 발생하였습니다.", "closePop('alertPop', '')", "");
				}
			});

		}
	</script>


<%TopSubMenuTitle = "이벤트"%>
<!-- #include virtual="/INC/TopSub.asp" -->

<%
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
%>
    <!-- Main -->
    <main id="container" class="container">
        <div class="sub_content">
            <section class="evt-como">
                <div class="evt-intro-img">
                    <img src="/images/tmp/@evt-intro-como.png" alt="꼬모 2행시, 휠라 학용품 키드 선착순 증성 이벤트">
                </div>

                <div class="evt-cont">
                    <div class="upto-link">
                        <a href="javascript:openExternal('https://www.facebook.com/153443871399191/posts/2063416120401947/');" target="_blank" class="facebook">
							<img src="/images/tmp/@como-facebook.png" alt="페이스 북으로 이행시 하러가기">
						</a>
                        <a href="<%IF U_MFLAG="Y" THEN%>javascript:openEvent();<%ELSE%>javascript:Loginchk();<%END IF%>" class="photo">
							<img src="/images/tmp/@como-img-upload.png" alt="인증샷 업로드">
						</a>
                    </div>

					<div id="AnotherEventList">
					</div>

                </div>

                <div class="inf-type2">
                    <p class="tit">&#8251; 주의사항</p>
                    <ul>
                        <li>인증샷 이벤트는 ID당 1회만 참여 가능합니다. 수정은 불가능하며, 삭제 후 재업로드만 가능합니다.</li>
                        <li>당첨자는 발표 후 개별 연락 드립니다. 광고 수신 동의를 확인해 주세요.</li>
                        <li>파우치 구성품은 모두 동일합니다.</li>
                        <li>꼬모 구매 인증 후 이벤트 참여시 당첨확률 업!</li>
                    </ul>
                </div>

				<div>

<% 
If oRs.EOF Then
Else
%>
                    <div class="event-item-list">
                        <div class="inner-cont">
                            <div class="sort-ty2">
							<%
								i = 1
								If CategoryIDX = "" Then CategoryIDX = oRs("IDX")
								Do While Not oRs.EOF
							%>
								<button type="button" class="length3" id="Category<%=i%>" onclick="CategoryView(<%=EventIDX%>, <%=oRs("IDX")%>, <%=i%>, <%=CategoryCount%>);"><%=oRs("CateName")%></button>
							<%
									If Trim(CategoryIDX) = Trim(oRs("IDX")) Then
										nCategoryView = i
									End If

									i = i + 1
									oRs.MoveNext
								Loop
							%>
                            </div>
                        </div>

                        <div class="cont">
                            <div class="tit" id="CateName"></div>

                            <div class="item-list">
                                <ul class="listview" id="ProductList">

                                </ul>
                            </div>

                        </div>
                    </div>

					<script type="text/javascript">
						CategoryView(<%=EventIDX%>, <%=CategoryIDX%>, <%=nCategoryView%>, <%=CategoryCount%>);
					</script>
<%
End If
oRs.Close
%>
				</div>
            </section>
        </div>
    </main>



	<form name="form" method="get">
		<input type="hidden" name="Page" />
		<input type="hidden" name="ISTopN" />
	</form>


	<script type="text/javascript">
		/* 이벤트 리스트 */
		function get_EventList(page) {
			//location.href = "/ASP/Event/Ajax/EventView_Como_1901_List.asp?EventIdx=<%=EventIdx%>&page="+page;
			//return;
			$.ajax({
				type: "post",
				url: "/ASP/Event/Ajax/EventView_Como_1901_List.asp",
				async: false,
				data: "EventIdx=<%=EventIdx%>&page="+page,
				dataType: "text",
				success: function (data) {
					var splitData = data.split("|||||");
					var result = splitData[0];
					var cont = splitData[1];

					if (result == "OK") {
						$("#AnotherEventList").html(cont);
						return;
					}
					else {
						alert(cont);
						//PageReload();
						return;
					}
				},
				error: function (data) {
					//alert(data.responseText);
					common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.");
				}
			});

		}

		/* 이벤트 등록창 열기 */
		function openEvent() {
			$.ajax({
				type: "post",
				url: "/ASP/Event/Ajax/EventWrite.asp",
				async: false,
				data: "EventIdx=<%=EventIdx%>",
				dataType: "text",
				success: function (data) {
					$("#eventView").html(data);
					$("#eventView").show();

					//Pop Up 높이 값
					var _windowHeight = $(window).height();
					var _maxHeight = _windowHeight - 100;
					//_this.css('height', _maxHeight);

					// Pop Up 호출 시 전체 스크롤 제거
					$("body").css("overflow", "hidden");

				},
				error: function (data) {
					alert(data.responseText);
					common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.");
				}
			});
		}

		/* 이벤트 등록창 닫기 */
		function closeEvent() {
			$("#eventView").hide();
			$("body").css("overflow", "auto");
		}

		/* 이벤트 등록 처리 */
		function eventWrite() {

			var comment = $("form[name='EventForm'] textarea[name='Comment']").val().length;
			if (comment > 50 || comment < 1) {
				common_msgPopOpen("", "코멘트는 50자 이내로 입력하세요.");
				$("form[name='EventForm'] textarea[name='Comment']").focus();
				return;
			}
			var filename = $("form[name='EventForm'] input[name='FileName']").val();
			if (filename == "") {
				common_msgPopOpen("", "인증샷 이미지를 선택하세요.");
				return;
			}

			common_msgPopOpen("", "현재 내용으로 이벤트에 참여 하시겠습니까?", "eventWriteOk();", "", "C");
		}

		function eventWriteOk() {
			$.ajax({
				type: "post",
				url: "/ASP/Event/Ajax/EventWriteOK.asp",
				async: false,
				data: $("form[name='EventForm']").serialize(),
				dataType: "text",
				success: function (data) {
					var splitData = data.split("|||||");
					var result = splitData[0];
					var cont = splitData[1];

					if (result == "OK") {
							common_msgPopOpen("", "이벤트 신청이 완료되었습니다.", "closeEvent();get_EventList(1)");
							//closeEvent();
							return;
					}
					else if (result == "LOGIN") {
						PareReload();
						return;
					}
					else {
						common_msgPopOpen("", cont);
						return;
					}
				},
				error: function (data) {
					alert(data.responseText);
					common_msgPopOpen("", "처리 중 오류가 발생하였습니다.[02]");
				}
			});

			
		}

		/* 이벤트 댓글 삭제 */
		function commentDel(idx) {
			common_msgPopOpen("", "댓글을 삭제하시겠습니까?", "commentDelOk('" + idx + "');", "", "C");
		}

		function commentDelOk(idx) {
			$.ajax({
				type: "post",
				url: "/ASP/Event/Ajax/EventDelOK.asp",
				async: false,
				data: "EventIdx=<%=EventIdx%>&idx="+idx,
				dataType: "text",
				success: function (data) {
					var splitData = data.split("|||||");
					var result = splitData[0];
					var cont = splitData[1];

					if (result == "OK") {
						common_msgPopOpen("", "댓글 삭제가 완료되었습니다.", "closeEvent();get_EventList(1)");
						return;
					}
					else if (result == "LOGIN") {
						PageReload();
						return;
					}
					else {
						common_msgPopOpen("", cont, "PageReload();");
						return;
					}
				},
				error: function (data) {
					//alert(data.responseText);
					common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.");
				}
			});
		}

		/* 이벤트 이미지 선택창 열기 */
		function openEventImageSearch() {
			var ImgCount = parseInt($("form[name='EventForm'] input[name='UploadFilesCount']").val());
			if (ImgCount >= 1) {
				common_msgPopOpen("", '첨부 이미지는 1개만 첨부 가능합니다.')
			}
			else {
				$("form[name='EventForm'] input[name='FileName']").trigger('click');
			}
		}

		/* 이벤트 이미지 추가 */
		function eventImageAdd() {
			var ImgCount = parseInt($("form[name='EventForm'] input[name='UploadFilesCount']").val());
			if (ImgCount >= 1) {
				common_msgPopOpen("", '첨부 이미지는 1개만 첨부 가능합니다.')
				return;
			}

			var img = $("form[name='EventForm'] input[name='FileName']").val().trim();
			if (img.length > 0) {
				lng = img.length;
				ext = img.substring(lng - 4, lng);
				ext = ext.toLowerCase();
				if (!(ext == ".jpg" || ext == ".gif" || ext == ".png" || ext == "jpeg")) {
					common_msgPopOpen("", "이미지는 gif, jpg, png, jpeg만 업도르 가능합니다.");
					return;
				}

				var formData = new FormData($("form[name='EventForm']")[0]);
				$.ajax({
					type: "post",
					url: "/ASP/Event/Ajax/EventImageTempUpload.asp",
					data: formData,
					async: false,
					contentType: false,
					cache: false,
					processData: false,
					dataType: "text",
					success: function (data) {
						var splitData = data.split("|||||");
						var result = splitData[0];
						var cont = splitData[1];

						if (result == "OK") {
							var splitData2 = cont.split("^^^^^");
							var imagePath = splitData2[0];
							var imageName = splitData2[1];
							eventImagePreView(imagePath, imageName);
						}
						else {
							common_msgPopOpen("", cont);
							return;
						}
					},
					error: function (data) {
						//alert(data.responseText);
						common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.");
					}
				});
			}
		}

		/* 이벤트 이미지 삭제 */
		function eventImageDelete(index) {
			// 삭제할 임시파일 경로
			var delFileName = $("form[name='EventForm'] .photo-list .img").eq(index - 1).find("img").attr("src");
			var splitFileName = delFileName.split("/");
			var filepath = delFileName.replace(splitFileName[splitFileName.length - 1], "");
			delFileName = splitFileName[splitFileName.length - 1];

			// 임시 이미지 삭제처리
			$.ajax({
				type: "post",
				url: "/ASP/Event/Ajax/EventImageTempDelete.asp",
				async: true,
				data: "FileName=" + delFileName,
				dataType: "text",
				success: function (data) {
					var splitData = data.split("|||||");
					var result = splitData[0];
					var cont = splitData[1];

					if (result == "OK") {
						// 미리보기 이미지 삭제
						$("form[name='EventForm'] .event-photo").html("");
						$("form[name='EventForm'] input[name='UploadFiles']").val("");
						$("form[name='EventForm'] input[name='UploadFilesCount']").val(0);
						//$("#file-name").val("인증샷 이미지를 등록하세요.");
					}
				},
				error: function (data) {
					//alert(data.responseText);
					common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.");
				}
			});
		}

		/* 이벤트 선택 이미지 미리보기 */
		function eventImagePreView(imagePath, imageName) {
			var i = parseInt($("form[name='EventForm'] input[name='UploadFilesCount']").val()) + 1;

			var html = "<li class=\"photo-list\">";
			html = html + "<button type=\"button\" onclick=\"eventImageDelete(" + i + ")\">삭제</button>";
			html = html + "<div class=\"img\">";
			html = html + "<img src=\"" + imagePath + imageName + "\" alt=\"후기 이미지\">";
			html = html + "</div>";
			html = html + "</div>";
			html = html + "</li>";
			$("form[name='EventForm'] .event-photo").append(html);

			$("form[name='EventForm'] input[name='UploadFiles']").val(imageName);
			$("form[name='EventForm'] input[name='UploadFilesCount']").val(1);
		}

		function Loginchk() {
			common_msgPopOpen("", "정회원만 이용가능한 이벤트 입니다.<br />로그인 또는 회원전환 후 이용하세요.");
			return;
		}


		/* 이벤트 이미지 보기 */
		function openImageView(filename) {
			$.ajax({
				type: "post",
				url: "/ASP/Event/Ajax/ImageView.asp",
				async: false,
				data: "filename="+filename,
				dataType: "text",
				success: function (data) {

					$("#eventImageView").html(data);
					$("#eventImageView").show();
					
					//Pop Up 높이 값
					var _this = $('.zoom-pop');
					var _imgHeight = $('.zoom-pop img').height();
					var _windowHeight = $(window).height();
					var _maxHeight = _windowHeight - 100;
					_this.css('max-height', '70%');
					if (parseInt(_imgHeight) > parseInt(_maxHeight)) {
						_this.css('height', _maxHeight);
					} else {
						_this.css('height', 'auto');
					}
					
					$("body").css("overflow", "hidden");
				},
				error: function (data) {
					//alert(data.responseText);
					alert("처리 도중 오류가 발생하였습니다.");
				}
			});

		}

		/* 이벤트 이미지 닫기 */
		function closeImageView() {
			$("#eventImageView").hide();
			$("body").css("overflow", "auto");
		}


		get_EventList(1);
	</script>



<!-- #include virtual="/INC/Footer.asp" -->

<!-- 이벤트 참여하기 POP -->
<section class="wrap-pop" id="eventView"></section>
<!-- //이벤트 참여하기 POP -->
<!-- 이벤트 이미지 뷰 POP -->
<section class="wrap-pop" id="eventImageView"></section>
<!-- 이벤트 이미지 뷰 POP -->
<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>