/* 스크롤 이동 */
function moveScroll() {
	/*
	var sTop = $(".t-type4-inform").offset().top - $(".header").height() - 2;
	//$("html, body").animate({ scrollTop: sTop }, 300);
	$("html, body").scrollTop(sTop);
	*/
}

/* 주문취소/반품/교환 리스트 */
function getOrderList(page, cancelType, moveFlag) {
	$("#Page").val(page);
	$("#SCancelType").val(cancelType);

	//location.href = "/ASP/Mypage/Ajax/OrderCRXList.asp?" + $("#form").serialize();
	//return;

	$.ajax({
		type: "post",
		url: "/ASP/Mypage/Ajax/OrderCRXList.asp",
		async: false,
		data: $("#form").serialize(),
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				$("#OrderList").html(cont);
				// 스크롤 이동
				if (moveFlag == "MOVE") {
					moveScroll();
				}
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
			alert(data.responseText);
			common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
		}
	});
}

/* 주문상세내역 */
function getOrderDetail(opIdx) {
	$("#SOPIdx").val(opIdx);
	//location.href = "/ASP/Mypage/Ajax/OrderCRXDetail.asp?" + $("#form").serialize();
	//return;
	$.ajax({
		type: "post",
		url: "/ASP/Mypage/Ajax/OrderCRXDetail.asp",
		async: false,
		data: $("#form").serialize(),
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				$("#DimDepth1").html(cont);
				openPop('DimDepth1');
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
			alert(data.responseText);
			common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
		}
	});
}

/* 교환/반품 불가 안내 팝업창 열기 */
function openOrderChangeReturnGuide() {
	$.ajax({
		type: "post",
		url: "/ASP/MyPage/Ajax/OrderChangeReturnGuide.asp",
		async: false,
		data: "",
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				$("#DimDepth2").html(cont);
				openPop('DimDepth2');
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
