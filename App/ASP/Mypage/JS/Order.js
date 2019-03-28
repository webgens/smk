/* 스크롤 이동 */
function moveScroll() {
	/*
	var sTop = $("#OrderList").offset().top;	// - $(".header").height() - 2;
	//$("html, body").animate({ scrollTop: sTop }, 300);
	$("html, body").scrollTop(sTop);
	*/
}

/* 주문배송조회/영수증발급 탭 전환시 */
function chgTab(listType) {
	$("#ListType").val(listType);

	if ($("#ListType").val() == "Receipt") {
		$("#tabs .tab-selector li").eq(0).removeClass("active");
		$("#tabs .tab-selector li").eq(1).addClass("active");
	} else {
		$("#tabs .tab-selector li").eq(0).addClass("active");
		$("#tabs .tab-selector li").eq(1).removeClass("active");
	}

	getOrderList(1, "", "");
	//document.form.submit();
}

/* 주문 리스트를 다시 불러온다 */
function orderListReload() {
	if ($("#OrderList").length > 0) {
		closePop('msgPopup');
		closePop('DimDepth1');

		var page = $("#Page").val();
		var orderState = $("#SOrderState").val();

		getOrderList(page, orderState, "MOVE");
	} else {
		PageReload();
	}
}
/* 주문 상세를 다시 불러온다 */
function orderDetailReload() {
	closePop('msgPopup');
	closePop('DimDepth1');

	var opIdx = $("#SOPIdx").val();

	getOrderDetail(opIdx);
}

/* 주문/영수증발급 리스트 */
function getOrderList(page, orderState, moveFlag) {
	$("#Page").val(page);
	$("#SOrderState").val(orderState);

	var url = "";
	if ($("#ListType").val() == "Receipt") {
		$("#tabs .tab-selector li").eq(0).removeClass("active");
		$("#tabs .tab-selector li").eq(1).addClass("active");
		url = "/ASP/Mypage/Ajax/ReceiptList.asp";
		//location.href = url + "?" + $("#form").serialize();
		//return;
	} else {
		$("#tabs .tab-selector li").eq(0).addClass("active");
		$("#tabs .tab-selector li").eq(1).removeClass("active");
		url = "/ASP/Mypage/Ajax/OrderList.asp";
	}


	$.ajax({
		type: "post",
		url: url,
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
	//location.href = "/ASP/Mypage/Ajax/OrderDetail.asp?" + $("#form").serialize();
	//return;
	$.ajax({
		type: "post",
		url: "/ASP/Mypage/Ajax/OrderDetail.asp",
		async: false,
		data: $("#form").serialize(),
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				//$("#OrderList").html(cont);
				// 스크롤 이동
				//moveScroll();

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

/* 주문상세리스트 */
function getOrderDetailList(orderCode) {
	//location.href = "/ASP/Mypage/Ajax/OrderDetailList.asp?" + "OrderCode=" + orderCode;
	//return;
	$.ajax({
		type: "post",
		url: "/ASP/Mypage/Ajax/OrderDetailList.asp",
		async: false,
		data: "OrderCode=" + orderCode,
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

/* 입금전 주문취소 팝업 띄우기 */
function openNonDepositOrderCancel(orderCode) {
	//location.href = "/ASP/Mypage/Ajax/NonDepositOrderCancel.asp?" + "OrderCode=" + orderCode;
	//return;
	$.ajax({
		type: "post",
		url: "/ASP/Mypage/Ajax/NonDepositOrderCancel.asp",
		async: false,
		data: "OrderCode=" + orderCode,
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				$("#DimDepth1").html(cont);
				openPop('DimDepth1');
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
/* 입금전 주문취소 */
function nonDepositOrderCancel(orderCode) {
	common_msgPopOpen("", "주문을 취소 하시겠습니까?", "nonDepositOrderCancel2('" + orderCode + "')", "msgPopup", "C");
}
function nonDepositOrderCancel2(orderCode) {
	//location.href = "/ASP/Mypage/Ajax/NonDepositOrderCancelOk.asp?" + "OrderCode=" + orderCode;
	//return;
	$.ajax({
		type: "post",
		url: "/ASP/Mypage/Ajax/NonDepositOrderCancelOk.asp",
		async: false,
		data: "OrderCode=" + orderCode,
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				common_msgPopOpen("", "주문이 취소 되었습니다.", "", "msgPopup", "N");
				closePop('DimDepth1');

				var page = $("#Page").val();
				var orderState = $("#SOrderState").val();
				getOrderList(page, orderState, "MOVE");
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

/* 주문취소 팝업 띄우기 */
function openOrderCancel(cancelType, orderCode, opIdx) {
	var url = "";
	// 주문취소
	if (cancelType == "C") {
		url = "/ASP/Mypage/Ajax/OrderCancel.asp";
	// 주문취소요청
	} else if (cancelType == "R") {
		url = "/ASP/Mypage/Ajax/OrderCancelRequest.asp";
	}

	if (url == "") {
		common_msgPopOpen("", "취소구분이 없습니다.", "", "msgPopup", "N");
		return;
	}

	//location.href = url + "?" + "OrderCode=" + orderCode + "&OPIdx=" + opIdx;
	//return;
	$.ajax({
		type: "post",
		url: url,
		async: false,
		data: "OrderCode=" + orderCode + "&OPIdx=" + opIdx,
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				$("#DimDepth1").html(cont);
				getRefundPrice(cancelType);
				openPop('DimDepth1');
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

/* 결제완료건 주문취소시 환불예상금액 계산 */
function getRefundPrice(cancelType) {
	//location.href = "/ASP/Mypage/Ajax/OrderCancel.asp?" + "OrderCode=" + orderCode;
	//return;
	$.ajax({
		type: "post",
		url: "/ASP/Mypage/Ajax/OrderCancelRefundInfo.asp",
		async: true,
		data: $("form[name='OrderCancelForm']").serialize(),
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				$("#RefundInfo").html(cont);

				// 취소신청일 경우 환불계좌 표시체크
				if (cancelType == "R") {
					chkRefundAccount("OrderCancel");
				}
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

/* 가상계좌 또는 에스크로 부분취소요청시 환불계좌정보를 표시한다 */
function chkRefundAccount(popID) {
	var payType				= $("#" + popID + " input[name='PayType']").val();
	var escrowFlag			= $("#" + popID + " input[name='EscrowFlag']").val();
	var totalSettlePrice	= $("#" + popID + " input[name='TotalSettlePrice']").val();
	var refundPrice			= $("#" + popID + " input[name='RefundPrice']").val();

	if (Number(refundPrice) <= 0) {
		$("#" + popID + " .refundaccount-info").hide();
	}
	else if (payType == "V") {
		$("#" + popID + " .refundaccount-info").show();
	}
	else if (escrowFlag == "Y" && Number(refundPrice) != Number(totalSettlePrice)) {
		$("#" + popID + " .refundaccount-info").show();
	}
	else {
		$("#" + popID + " .refundaccount-info").hide();
	}
}

/* 주문 취소/취소신청 */
function orderCancel(cancelType) {
	var msg = "";
	var url = "";
	// 주문취소
	if (cancelType == "C") {
		msg = "취소";
		url = "/ASP/Mypage/Ajax/OrderCancelOk.asp";
		// 주문취소요청
	} else if (cancelType == "R") {
		msg = "취소신청";
		url = "/ASP/Mypage/Ajax/OrderCancelRequestOk.asp";
	}

	if (url == "") {
		common_msgPopOpen("", "취소구분이 없습니다.", "", "msgPopup", "N");
		return;
	}

	var payType = $("form[name='OrderCancelForm'] input[name='PayType']").val();
	var escrowFlag = $("form[name='OrderCancelForm'] input[name='EscrowFlag']").val();
	var totalOrderCnt = $("form[name='OrderCancelForm'] input[name='TotalOrderCnt']").val();
	var checkOrderCnt = $("form[name='OrderCancelForm'] input[name='OPIdx']:checked").length;
	var totalSettlePrice = $("form[name='OrderCancelForm'] input[name='TotalSettlePrice']").val();
	var refundPrice = $("form[name='OrderCancelForm'] input[name='RefundPrice']").val();

	if (Number(checkOrderCnt) == 0) {
		common_msgPopOpen("", msg + "할 상품을 선택해 주십시오.", "", "msgPopup", "N");
		return;
	}

	if (Number(refundPrice) < 0) {
		common_msgPopOpen("", "환불금액이 부족하여 " + msg + " 하실 수 없습니다.\n고객센터에 문의해 주십시오.", "", "msgPopup", "N");
		return;
	}

	// 취소일 경우 에스크로 주문은 부분취소 불가
	if (cancelType == "C") {
		if (escrowFlag == "Y" && Number(checkOrderCnt) < Number(totalOrderCnt)) {
			common_msgPopOpen("", "에스크로 적용 주문은 주문상품 전체" + msg + "만 가능합니다.", "", "msgPopup", "N");
			return;
		}
	}

	// 취소신청일 경우 신청사유 체크
	if (cancelType == "R") {
		var reasonType = alltrim($("form[name='OrderCancelForm'] select[name='ReasonType'] option:selected").val());
		if (reasonType.length == 0) {
			common_msgPopOpen("", "취소사유를 선택해 주십시오.", "", "msgPopup", "N");
			$("form[name='OrderCancelForm'] select[name='ReasonType']").focus();
			return;
		}
	}


	// 가상계좌 or 에스크로 부분취소시 환불계좌정보 체크
	if (payType == "V" || (escrowFlag == "Y" && Number(refundPrice) != Number(totalSettlePrice))) {
		var refundBankCode = alltrim($("form[name='OrderCancelForm'] select[name='RefundBankCode'] option:selected").val());
		if (refundBankCode.length == 0) {
			common_msgPopOpen("", "환불은행을 선택해 주십시오.", "", "msgPopup", "N");
			$("form[name='OrderCancelForm'] select[name='RefundBankCode']").focus();
			return;
		}

		var refundAccountNum = alltrim($("form[name='OrderCancelForm'] input[name='RefundAccountNum']").val());
		if (refundAccountNum.length == 0) {
			common_msgPopOpen("", "환불계좌번호를 입력해 주십시오.", "", "msgPopup", "N");
			$("form[name='OrderCancelForm'] input[name='RefundAccountNum']").focus();
			return;
		}

		var refundAccountName = alltrim($("form[name='OrderCancelForm'] input[name='RefundAccountName']").val());
		if (refundAccountName.length == 0) {
			common_msgPopOpen("", "환불계좌 예금주명을 입력해 주십시오.", "", "msgPopup", "N");
			$("form[name='OrderCancelForm'] input[name='RefundAccountName']").focus();
			return;
		}

		var refundPhone1 = alltrim($("form[name='OrderCancelForm'] select[name='RefundPhone1'] option:selected").val());
		if (refundPhone1.length == 0) {
			common_msgPopOpen("", "연락처를 선택해 주십시오.", "", "msgPopup", "N");
			$("form[name='OrderCancelForm'] select[name='RefundPhone1']").focus();
			return;
		}

		var refundPhone23 = alltrim($("form[name='OrderCancelForm'] input[name='RefundPhone23']").val());
		if (refundPhone23.length == 0) {
			common_msgPopOpen("", "연락처를 입력해 주십시오.", "", "msgPopup", "N");
			$("form[name='OrderCancelForm'] input[name='RefundPhone23']").focus();
			return;
		}
	}

	common_msgPopOpen("", "주문을 " + msg + " 하시겠습니까?", "orderCancel2('" + cancelType + "')", "msgPopup", "C");
}
function orderCancel2(cancelType) {
	var msg = "";
	var url = "";
	// 주문취소
	if (cancelType == "C") {
		msg = "취소";
		url = "/ASP/Mypage/Ajax/OrderCancelOk.asp";
		// 주문취소요청
	} else if (cancelType == "R") {
		msg = "취소신청";
		url = "/ASP/Mypage/Ajax/OrderCancelRequestOk.asp";
	}

	if (url == "") {
		common_msgPopOpen("", "취소구분이 없습니다.", "", "msgPopup", "N");
		return;
	}

	//location.href = "/ASP/Mypage/Ajax/OrderCancelOk.asp?" + "OrderCode=" + orderCode;
	//return;
	$.ajax({
		type: "post",
		url: url,
		async: false,
		data: $("form[name='OrderCancelForm']").serialize(),
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				common_msgPopOpen("", "주문이 " + msg + " 되었습니다.", "orderListReload()", "msgPopup", "N");
				/*
				closePop('DimDepth1');

				var page = $("#Page").val();
				var orderState = $("#SOrderState").val();
				getOrderList(page, orderState, "MOVE");
				*/
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

/* 주문 교환/반품 신청 팝업 띄우기 */
function openOrderChangeReturn(cancelType, orderCode, opIdx) {
	var url = "";
	// 주문교환신청
	if (cancelType == "X") {
		url = "/ASP/Mypage/Ajax/OrderChangeRequest.asp";
	// 주문반품신청
	} else if (cancelType == "R") {
		url = "/ASP/Mypage/Ajax/OrderReturnRequest.asp";
	}

	if (url == "") {
		common_msgPopOpen("", "신청구분이 없습니다.", "", "msgPopup", "N");
		return;
	}

	//location.href = url + "?" + "OrderCode=" + orderCode + "&OPIdx=" + opIdx;
	//return;
	$.ajax({
		type: "post",
		url: url,
		async: false,
		data: "OrderCode=" + orderCode + "&OPIdx=" + opIdx,
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				$("#DimDepth1").html(cont);
				openPop('DimDepth1');
				// 반품신청일 경우 환불금액 계산
				if (cancelType == "R") {
					getReturnRefundPrice();
				}
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

/* 교환/반품 사유 변경시 */
function chgReasonType() {
	var reasonType = $("form[name='OrderChangeReturnForm'] select[name='ReasonType'] option:selected").val();
	// 제품불량(X02,R02), 오배송(X03,R03)일 경우만 배송비 슈마커부담 보이기
	if (reasonType == "X02" || reasonType == "X03" || reasonType == "R02" || reasonType == "R03") {
		$("form[name='OrderChangeReturnForm'] input:radio[name='DelvFeeType']").parent().hide();
		$("form[name='OrderChangeReturnForm'] #DelvFeeType_1").parent().show();
		$("form[name='OrderChangeReturnForm'] #DelvFeeType_1").prop("checked", true);
	} else {
		$("form[name='OrderChangeReturnForm'] input:radio[name='DelvFeeType']").parent().show();
		$("form[name='OrderChangeReturnForm'] #DelvFeeType_1").parent().hide();
		$("form[name='OrderChangeReturnForm'] input:radio[name='DelvFeeType']").eq(0).prop("checked", true);
	}
	getReturnRefundPrice();
}

/* 주문반품신청시 환불예상금액 계산 */
function getReturnRefundPrice() {
	//location.href = "/ASP/Mypage/Ajax/OrderReturnRefundInfo.asp?" + $("form[name='OrderChangeReturnForm']").serialize();
	//return;
	$.ajax({
		type: "post",
		url: "/ASP/Mypage/Ajax/OrderReturnRefundInfo.asp",
		async: true,
		data: $("form[name='OrderChangeReturnForm']").serialize(),
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				$("#RefundInfo").html(cont);

				chkRefundAccount("OrderReturn");
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

// 교환/반품지 주소 변경 팝업 열기
function openZipCodeSearch(addrType) {
	var name = "";
	var phone = "";
	var zipcode = "";
	var addr1 = "";
	var addr2 = "";

	if (addrType == "Return") {
		name	= $("form[name='OrderChangeReturnForm'] input[name='ReturnName']").val();
		phone	= $("form[name='OrderChangeReturnForm'] input[name='ReturnHp']").val();
		zipcode = $("form[name='OrderChangeReturnForm'] input[name='ReturnZipCode']").val();
		addr1	= $("form[name='OrderChangeReturnForm'] input[name='ReturnAddr1']").val();
		addr2	= $("form[name='OrderChangeReturnForm'] input[name='ReturnAddr2']").val();
	}
	else if (addrType == "Receive") {
		name	= $("form[name='OrderChangeReturnForm'] input[name='ReceiveName']").val();
		phone	= $("form[name='OrderChangeReturnForm'] input[name='ReceiveHp']").val();
		zipcode = $("form[name='OrderChangeReturnForm'] input[name='ReceiveZipCode']").val();
		addr1	= $("form[name='OrderChangeReturnForm'] input[name='ReceiveAddr1']").val();
		addr2	= $("form[name='OrderChangeReturnForm'] input[name='ReceiveAddr2']").val();
	}
	else {
		common_msgPopOpen("", "주소지 구분정보가 없습니다.", "", "msgPopup", "N");
		return;
	}

	var param = "AddrType=" + addrType;
	param += "&Name="		+ escape(name);
	param += "&Phone="		+ phone;
	param += "&ZipCode="	+ zipcode;
	param += "&Addr1="		+ escape(addr1);
	param += "&Addr2="		+ escape(addr2);

	//location.href = "/ASP/MyPage/Ajax/ChangeReturnZipCD.asp?" + param;
	//return;

	$.ajax({
		type: "post",
		url: "/ASP/MyPage/Ajax/ChangeReturnZipCD.asp",
		async: false,
		data: param,
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

/* 교환/반품 주소지 변경 셋팅 */
function setChangeReturnAddress(addrType) {
	var name = $("#ChangeReturnZipCD input[name='Name']").val();
	var zipcode = $("#ChangeReturnZipCD input[name='ZipCode']").val();
	var addr1 = $("#ChangeReturnZipCD input[name='Addr1']").val();
	var addr2 = $("#ChangeReturnZipCD input[name='Addr2']").val();
	var phone1 = alltrim($("#ChangeReturnZipCD select[name='Phone1'] option:selected").val());
	var phone23 = alltrim($("#ChangeReturnZipCD input[name='Phone23']").val());
	var phone2 = "";
	var phone3 = "";
	if (alltrim(phone23).length == 8) {
		phone2 = alltrim(phone23).substring(0, 4);
		phone3 = alltrim(phone23).substring(4);
	}
	else if (alltrim(phone23).length == 7) {
		phone2 = alltrim(phone23).substring(0, 3);
		phone3 = alltrim(phone23).substring(3);
	}
	else if (alltrim(phone23).length == 6) {
		phone2 = alltrim(phone23).substring(0, 2);
		phone3 = alltrim(phone23).substring(2);
	}
	var phone = phone1 + "-" + phone2 + "-" + phone3;

	if (alltrim(name).length == 0) {
		common_msgPopOpen("", "이름을 입력해 주십시오.", "", "msgPopup", "N");
		$("#ChangeReturnZipCD input[name='Name']").focus();
		return;
	}
	if (alltrim(zipcode).length == 0) {
		common_msgPopOpen("", "우편번호를 선택해 주십시오.", "", "msgPopup", "N");
		return;
	}
	if (alltrim(addr1).length == 0) {
		common_msgPopOpen("", "주소를 선택해 주십시오.", "", "msgPopup", "N");
		return;
	}
	if (alltrim(addr2).length == 0) {
		common_msgPopOpen("", "상세주소를 입력해 주십시오.", "", "msgPopup", "N");
		$("#ChangeReturnZipCD input[name='Addr2']").focus();
		return;
	}
	if (alltrim(phone1).length == 0) {
		common_msgPopOpen("", "휴대전화 앞번호를 선택해 주십시오.", "", "msgPopup", "N");
		$("#ChangeReturnZipCD select[name='Phone1']").focus();
		return;
	}
	if (alltrim(phone23).length == 0) {
		common_msgPopOpen("", "휴대전화 뒷번호를 입력해 주십시오.", "", "msgPopup", "N");
		$("#ChangeReturnZipCD input[name='Phone23']").focus();
		return;
	}

	if (addrType == "Return") {
		$("form[name='OrderChangeReturnForm'] input[name='ReturnName']").val(name);
		$("form[name='OrderChangeReturnForm'] input[name='ReturnHP']").val(phone);
		$("form[name='OrderChangeReturnForm'] input[name='ReturnZipCode']").val(zipcode);
		$("form[name='OrderChangeReturnForm'] input[name='ReturnAddr1']").val(addr1);
		$("form[name='OrderChangeReturnForm'] input[name='ReturnAddr2']").val(addr2);
		$("form[name='OrderChangeReturnForm'] .ReturnName").html("<span>" + name + "</span><span>" + phone + "</span>");
		$("form[name='OrderChangeReturnForm'] .ReturnAddr").html("[" + zipcode + "] " + addr1 + " " + addr2);
	}
	else if (addrType == "Receive") {
		$("form[name='OrderChangeReturnForm'] input[name='ReceiveName']").val(name);
		$("form[name='OrderChangeReturnForm'] input[name='ReceiveHP']").val(phone);
		$("form[name='OrderChangeReturnForm'] input[name='ReceiveZipCode']").val(zipcode);
		$("form[name='OrderChangeReturnForm'] input[name='ReceiveAddr1']").val(addr1);
		$("form[name='OrderChangeReturnForm'] input[name='ReceiveAddr2']").val(addr2);
		$("form[name='OrderChangeReturnForm'] .ReceiveName").html("<span>" + name + "</span><span>" + phone + "</span>");
		$("form[name='OrderChangeReturnForm'] .ReceiveAddr").html("[" + zipcode + "] " + addr1 + " " + addr2);
	}

	closePop("DimDepth2");
}

/* 주문 교환/반품 신청처리 */
function orderChangeReturnCheck(cancelType) {
	var msg = "";
	if (cancelType == "X") {
		msg = "교환";
	} else if (cancelType == "R") {
		msg = "반품";
	}

	if (msg == "") {
		common_msgPopOpen("", "신청구분이 없습니다.", "", "msgPopup", "N");
		return;
	}

	var reasonType = $("form[name='OrderChangeReturnForm'] select[name='ReasonType'] option:selected").val();
	if (reasonType.length == 0) {
		common_msgPopOpen("", msg + " 사유를 선택해 주십시오.", "", "msgPopup", "N");
		$("form[name='OrderChangeReturnForm'] select[name='ReasonType']").focus();
		return;
	}

	// 교환신청일 경우 신청사유 체크
	if (cancelType == "X") {
		var sizeCD = $("form[name='OrderChangeReturnForm'] input[name='SizeCD']").val();
		var chgSizeCD = alltrim($("form[name='OrderChangeReturnForm'] select[name='ChgSizeCD'] option:selected").val());

		if (chgSizeCD == "") {
			common_msgPopOpen("", "사이즈를 선택 하여 주세요.", "", "msgPopup", "N");
			return;
		}

		if (sizeCD == chgSizeCD) {
			if (reasonType == "X01") {
				common_msgPopOpen("", "동일사이즈 교환시는 사유를 사이즈교환으로 선택하실 수 없습니다.\n다른 교환사유를 선택해 주십시오.", "", "msgPopup", "N");
				$("form[name='OrderChangeReturnForm'] select[name='ReasonType']").focus();
				return;
			}
		} else {
			if (reasonType != "X01") {
				common_msgPopOpen("", "다른 사이즈 교환시는 사유를 사이즈교환만 선택하실 수 있습니다.", "", "msgPopup", "N");
				$("form[name='OrderChangeReturnForm'] select[name='ReasonType']").focus();
				return;
			}
		}
	}
	else if (cancelType == "R") {
		var payType = $("form[name='OrderChangeReturnForm'] input[name='PayType']").val();
		var escrowFlag = $("form[name='OrderChangeReturnForm'] input[name='EscrowFlag']").val();
		var totalSettlePrice = $("form[name='OrderChangeReturnForm'] input[name='TotalSettlePrice']").val();
		var refundPrice = $("form[name='OrderChangeReturnForm'] input[name='RefundPrice']").val();

		// 가상계좌 or 에스크로 부분취소시 환불계좌정보 체크
		if (payType == "V" || (escrowFlag == "Y" && Number(refundPrice) != Number(totalSettlePrice))) {
			var refundBankCode = alltrim($("form[name='OrderChangeReturnForm'] select[name='RefundBankCode'] option:selected").val());
			if (refundBankCode.length == 0) {
				common_msgPopOpen("", "환불은행을 선택해 주십시오.", "", "msgPopup", "N");
				$("form[name='OrderChangeReturnForm'] select[name='RefundBankCode']").focus();
				return;
			}

			var refundAccountNum = alltrim($("form[name='OrderChangeReturnForm'] input[name='RefundAccountNum']").val());
			if (refundAccountNum.length == 0) {
				common_msgPopOpen("", "환불계좌번호를 입력해 주십시오.", "", "msgPopup", "N");
				$("form[name='OrderChangeReturnForm'] input[name='RefundAccountNum']").focus();
				return;
			}

			var refundAccountName = alltrim($("form[name='OrderChangeReturnForm'] input[name='RefundAccountName']").val());
			if (refundAccountName.length == 0) {
				common_msgPopOpen("", "환불계좌 예금주명을 입력해 주십시오.", "", "msgPopup", "N");
				$("form[name='OrderChangeReturnForm'] input[name='RefundAccountName']").focus();
				return;
			}

			var refundPhone1 = alltrim($("form[name='OrderChangeReturnForm'] select[name='RefundPhone1'] option:selected").val());
			if (refundPhone1.length == 0) {
				common_msgPopOpen("", "연락처를 선택해 주십시오.", "", "msgPopup", "N");
				$("form[name='OrderChangeReturnForm'] select[name='RefundPhone1']").focus();
				return;
			}

			var refundPhone23 = alltrim($("form[name='OrderChangeReturnForm'] input[name='RefundPhone23']").val());
			if (refundPhone23.length == 0) {
				common_msgPopOpen("", "연락처를 입력해 주십시오.", "", "msgPopup", "N");
				$("form[name='OrderChangeReturnForm'] input[name='RefundPhone23']").focus();
				return;
			}
		}

	}


	var delvFeeType = $("form[name='OrderChangeReturnForm'] input[name='DelvFeeType']:checked").val();
	// 쿠폰사용시 쿠폰리스트 팝업 띄우기
	if (delvFeeType == "7") {
		openDeliveryCouponList();
	} else {
		orderChangeReturn();
	}
}

/* 배송비 쿠폰 리스트 팝업창 열기 */
function openDeliveryCouponList() {
	$("form[name='OrderChangeReturnForm'] input[name='DeliveryCouponIdx']").val("");

	$.ajax({
		type: "post",
		url: "/ASP/MyPage/Ajax/DeliveryCouponList.asp",
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

/* 교환/반품 신청 처리 */
function orderChangeReturn() {
	var cancelType = $("form[name='OrderChangeReturnForm'] input[name='CancelType']").val();
	var msg = "";
	if (cancelType == "X") {
		msg = "교환";
	} else if (cancelType == "R") {
		msg = "반품";
	} else {
		common_msgPopOpen("", "신청구분이 없습니다.", "", "msgPopup", "N");
		return;
	}

	var delvFeeType = $("form[name='OrderChangeReturnForm'] input[name='DelvFeeType']:checked").val();

	// 쿠폰사용일 경우 쿠폰선택여부 체크
	var deliveryCouponIdx = "";
	if (delvFeeType == "7") {
		if ($("#DeliveryCouponList input[name='DeliveryCouponIdx']:checked").length == 0) {
			common_msgPopOpen("", "쿠폰을 선택해 주십시오.", "", "msgPopup", "N");
			return;
		}
		deliveryCouponIdx = $("#DeliveryCouponList input[name='DeliveryCouponIdx']:checked").val()
	}
	$("form[name='OrderChangeReturnForm'] input[name='DeliveryCouponIdx']").val(deliveryCouponIdx);

	common_msgPopOpen("", msg + "신청 하시겠습니까?", "orderChangeReturn2('" + delvFeeType + "')", "msgPopup", "C");
}
function orderChangeReturn2(delvFeeType) {
	// 신용카드, 계좌이체일 경우만 PG 결제창 보이기
	// if (delvFeeType == "6" || delvFeeType == "3") {
	// 	$("#OrderDiv").show();
	// }
	document.OrderChangeReturnForm.submit();
}
/* 결제완료시 */
function completePay() {
	//$("#OrderDiv").hide();
	orderDetailReload();
}
/* 결제취소시 */
function cancelPay() {
	//$("#OrderDiv").hide();
}

/* 교환/반품 불가 안내 팝업창 열기 */
/*
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
*/
/* 구매확정 */
function openOrderConfirm(orderCode, idx) {
	$.ajax({
		type: "post",
		url: "/ASP/MyPage/Ajax/OrderConfirm.asp",
		async: false,
		data: "OrderCode=" + orderCode + "&Idx=" + idx,
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				$("#DimDepth1").html(cont);
				openPop('DimDepth1');
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

/* 구매확정 처리 */
function orderConfirm(orderCode, idx) {
	$.ajax({
		type: "post",
		url: "/ASP/MyPage/Ajax/OrderConfirmOK.asp",
		async: false,
		data: "OrderCode=" + orderCode + "&Idx=" + idx,
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				//orderDetailReload();
				openOrderConfirmComplete(orderCode, idx);
			}
			else if (result == "LOGIN") {
				PageReload();
			}
			else {
				common_msgPopOpen("", cont, "orderListReload();", "msgPopup", "N");
				return;
			}
		},
		error: function (data) {
			alert(data.responseText);
			common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
		}
	});
}

/* 구매확정 완료 팝업 */
function openOrderConfirmComplete(orderCode, idx) {
	$.ajax({
		type: "post",
		url: "/ASP/MyPage/Ajax/OrderConfirmComplete.asp",
		async: false,
		data: "OrderCode=" + orderCode + "&Idx=" + idx,
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				$("#DimDepth1").html(cont);
				openPop('DimDepth1');
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

/* 구매후기 작성 팝업 */
function openReviewWrite(orderCode, idx) {
	$.ajax({
		type: "post",
		url: "/ASP/MyPage/Ajax/ReviewWrite.asp",
		async: false,
		data: "OrderCode=" + orderCode + "&Idx=" + idx,
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				$("#DimDepth1").html(cont);
				openPop('DimDepth1');
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

/* 구매후기 등록 처리 */
function reviewWrite() {
	var reviewType = "B";
	if (alltrim($("form[name='ReviewWriteForm'] input[name='UploadFiles']").val()).length > 0) {
		reviewType = "P";
	}

	var contents = alltrim($("form[name='ReviewWriteForm'] textarea[name='Contents']").val());
	if (contents.length == 0) {
		common_msgPopOpen("", "구매후기를 입력해 주십시오.", "", "msgPopup", "N");
		$("form[name='ReviewWriteForm'] textarea[name='Contents']").focus();
		return;
	}

	$.ajax({
		type: "post",
		url: "/ASP/MyPage/Ajax/ReviewWriteOK.asp",
		async: false,
		data: $("form[name='ReviewWriteForm']").serialize(),
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				//orderListReload();
				openReviewWriteComplete(reviewType);
			}
			else if (result == "LOGIN") {
				PageReload();
			}
			else {
				common_msgPopOpen("", cont, "", "msgPopup", "N");
				orderListReload();
				return;
			}
		},
		error: function (data) {
			alert(data.responseText);
			common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
		}
	});
}

/* 구매후기 작성완료 팝업 */
function openReviewWriteComplete(reviewType) {
	$.ajax({
		type: "post",
		url: "/ASP/MyPage/Ajax/ReviewWriteComplete.asp",
		async: false,
		data: "ReviewType=" + reviewType,
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				$("#DimDepth1").html(cont);
				openPop('DimDepth1');
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

/* 구매후기 이미지 선택창 열기 */
function openReviewImageSearch() {
	var ImgCount = parseInt($("form[name='ReviewWriteForm'] input[name='UploadFilesCount']").val());
	if (ImgCount >= 5) {
		common_msgPopOpen("", "첨부 이미지는 5개까지만 가능합니다.", "", "msgPopup", "N");
	}
	else {
		$("form[name='ReviewWriteForm'] input[name='FileName']").trigger('click');
	}
}

/* 구매후기 이미지 추가 */
function reviewImageAdd() {
	var ImgCount = parseInt($("form[name='ReviewWriteForm'] input[name='UploadFilesCount']").val());

	if (ImgCount >= 5) {
		common_msgPopOpen("", "첨부 이미지는 5개까지만 가능합니다.", "", "msgPopup", "N");
		return;
	}

	var img = $("form[name='ReviewWriteForm'] input[name='FileName']").val().trim();
	if (img.length > 0) {
		lng = img.length;
		ext = img.substring(lng - 4, lng);
		ext = ext.toLowerCase();
		if (!(ext == ".jpg" || ext == ".gif" || ext == ".png" || ext == "jpeg")) {
			common_msgPopOpen("", "이미지는 gif, jpg, png, jpeg만 업도르 가능합니다.", "", "msgPopup", "N");
			return;
		}

		var formData = new FormData($("form[name='ReviewWriteForm']")[0]);
		$.ajax({
			type: "post",
			url: "/ASP/Mypage/Ajax/ReviewImageTempUpload.asp",
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
					reviewImagePreView(imagePath, imageName);
				}
				else {
					common_msgPopOpen("", cont, "", "msgPopup", "N");
				}
			},
			error: function (data) {
				alert(data.responseText);
				common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
			}
		});
	}
}

/* 구매후기 이미지 삭제 */
function reviewImageDelete(index) {
	// 삭제할 임시파일 경로
	var delFileName = $("form[name='ReviewWriteForm'] .review-photo .photo-list").eq(index - 1).find("img").attr("src");
	var splitFileName = delFileName.split("/");
	var filepath = delFileName.replace(splitFileName[splitFileName.length - 1], "");
	delFileName = splitFileName[splitFileName.length - 1];

	// 미리보기 이미지 삭제
	$("form[name='ReviewWriteForm'] .review-photo .photo-list").eq(index - 1).remove();

	// 업로드할 이미지 리스트 작성
	var uploadFiles = "";
	$("form[name='ReviewWriteForm'] .review-photo .photo-list").each(function () {
		var imageUrl = $(this).find("img").attr("src");
		var splitImageUrl = imageUrl.split("/");
		if (uploadFiles == "") {
			uploadFiles = splitImageUrl[splitImageUrl.length - 1];
		}
		else {
			uploadFiles = uploadFiles + "|||||" + splitImageUrl[splitImageUrl.length - 1];
		}
	});

	$("form[name='ReviewWriteForm'] input[name='UploadFiles']").val(uploadFiles);

	// 업로드할 이미지 수
	var i = parseInt($("form[name='ReviewWriteForm'] input[name='UploadFilesCount']").val()) - 1;
	$("form[name='ReviewWriteForm'] input[name='UploadFilesCount']").val(i);

	// 임시 이미지 삭제처리
	$.ajax({
		type: "post",
		url: "/ASP/Mypage/Ajax/ReviewImageTempDelete.asp",
		async: true,
		data: "FileName=" + delFileName,
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				reReviewImagePreView(filepath);
			}
		},
		error: function (data) {
			//alert(data.responseText);
			//common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
		}
	});
}

/* 구매후기 선택 이미지 미리보기 재배치 */
function reReviewImagePreView(imagePath) {
	$("form[name='ReviewWriteForm'] .review-photo").html("");
	var i = parseInt($("form[name='ReviewWriteForm'] input[name='UploadFilesCount']").val());
	var UploadFiles = $("form[name='ReviewWriteForm'] input[name='UploadFiles']").val();

	if (i == 1) {
		var html = "<li class=\"photo-list\">";
		html = html + "<button type=\"button\" onclick=\"reviewImageDelete(" + i + ")\" >삭제</button>";
		html = html + "<div class=\"img\">";
		html = html + "<img src=\"" + imagePath + UploadFiles + "\" alt=\"후기 이미지\">";
		html = html + "</div>";
		html = html + "</li>";
		$("form[name='ReviewWriteForm'] .review-photo").append(html);
	} else {
		var UploadFilesArr = UploadFiles.split("|||||");
		for (var k = 1; k <= i; k++) {
			var html = "<li class=\"photo-list\">";
			html = html + "<button type=\"button\" onclick=\"reviewImageDelete(" + i + ")\" >삭제</button>";
			html = html + "<div class=\"img\">";
			html = html + "<img src=\"" + imagePath + UploadFilesArr[k - 1] + "\" alt=\"후기 이미지\">";
			html = html + "</div>";
			html = html + "</li>";
			$("form[name='ReviewWriteForm'] .review-photo").append(html);
		}
	}
}

/* 구매후기 선택 이미지 미리보기 */
function reviewImagePreView(imagePath, imageName) {
	var i = parseInt($("form[name='ReviewWriteForm'] input[name='UploadFilesCount']").val()) + 1;
	var html = "<li class=\"photo-list\">";
	html = html + "<button type=\"button\" onclick=\"reviewImageDelete(" + i + ")\" >삭제</button>";
	html = html + "<div class=\"img\">";
	html = html + "<img src=\"" + imagePath + imageName + "\" alt=\"후기 이미지\">";
	html = html + "</div>";
	html = html + "</li>";

	$("form[name='ReviewWriteForm'] .review-photo").append(html);

	var uploadFiles = $("form[name='ReviewWriteForm'] input[name='UploadFiles']").val().trim();

	if (uploadFiles == "") {
		uploadFiles = imageName;
	}
	else {
		uploadFiles = uploadFiles + "|||||" + imageName;
	}

	$("form[name='ReviewWriteForm'] input[name='UploadFiles']").val(uploadFiles);
	$("form[name='ReviewWriteForm'] input[name='UploadFilesCount']").val(i);
}

/* 작성한 구매후기 보기*/
function openReview(reviewIdx) {
	$.ajax({
		type: "post",
		url: "/ASP/MyPage/Ajax/ReviewView.asp",
		async: false,
		data: "Idx=" + reviewIdx,
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				$("#DimDepth1").html(cont);
				openPop('DimDepth1');
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


/* A/S 작성창 열기 */
function openAfterService(ocode, oidx) {
	$.ajax({
		type: "post",
		url: "/ASP/Mypage/Ajax/AfterServiceWrite.asp",
		async: false,
		data: "OrderCode=" + ocode + "&Order_Product_IDX=" + oidx,
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				$("#afterServiceView").html(data);
				$("#afterServiceView").show();

				//Pop Up 높이 값
				var _windowHeight = $(window).height();
				var _maxHeight = _windowHeight - 100;
				$("body").css("overflow", "hidden");
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
			//alert(data.responseText);
			common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
		}
	});
}

/* A/S 작성창 닫기 */
function closeAfterService() {
	$("#afterServiceView").hide();
	$("body").css("overflow", "auto");
}


/* A/S 등록 처리 */
function asWrite() {
	var requestCode = $("form[name='ASForm'] input[name='RequestCode']:checked").length;
	if (requestCode == 0) {
		common_msgPopOpen("", "신청구분을 선택하세요.", "", "msgPopup", "N");
		return;
	}
	var contents = alltrim($("form[name='ASForm'] textarea[name='Contents']").val());
	if (contents.length == 0) {
		common_msgPopOpen("", "신청내용을 입력해 주십시오.", "", "msgPopup", "N");
		$("form[name='ASForm'] textarea[name='Contents']").focus();
		return;
	}

	$.ajax({
		type: "post",
		url: "/ASP/MyPage/Ajax/AfterServiceWriteOK.asp",
		async: false,
		data: $("form[name='ASForm']").serialize(),
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				common_msgPopOpen("", "A/S 신청이 완료되었습니다.", "PageReload();");
				return;
			}
			else if (result == "LOGIN") {
				PageReload();
			}
			else {
				common_msgPopOpen("", cont, "orderListReload();", "msgPopup", "N");
				return;
			}
		},
		error: function (data) {
			//alert(data.responseText);
			common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
		}
	});
}

/* A/S 이미지 선택창 열기 */
function openAsImageSearch() {
	var ImgCount = parseInt($("form[name='ASForm'] input[name='UploadFilesCount']").val());
	if (ImgCount >= 5) {
		common_msgPopOpen("", "첨부 이미지는 5개까지만 가능합니다.", "", "msgPopup", "N");
	}
	else {
		$("form[name='ASForm'] input[name='FileName']").trigger('click');
	}
}

/* A/S 이미지 추가 */
function asImageAdd() {
	var ImgCount = parseInt($("form[name='ASForm'] input[name='UploadFilesCount']").val());
	if (ImgCount >= 5) {
		common_msgPopOpen("", "첨부 이미지는 5개까지만 가능합니다.", "", "msgPopup", "N");
		return;
	}

	var img = $("form[name='ASForm'] input[name='FileName']").val().trim();
	if (img.length > 0) {
		lng = img.length;
		ext = img.substring(lng - 4, lng);
		ext = ext.toLowerCase();
		if (!(ext == ".jpg" || ext == ".gif" || ext == ".png" || ext == "jpeg")) {
			common_msgPopOpen("", "이미지는 gif, jpg, png, jpeg만 업도르 가능합니다.", "", "msgPopup", "N");
			return;
		}

		var formData = new FormData($("form[name='ASForm']")[0]);
		$.ajax({
			type: "post",
			url: "/ASP/Mypage/Ajax/AsImageTempUpload.asp",
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
					asImagePreView(imagePath, imageName);
				}
				else {
					common_msgPopOpen("", cont, "", "msgPopup", "N");
				}
			},
			error: function (data) {
				//alert(data.responseText);
				common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
			}
		});
	}
}

/* A/S 이미지 삭제 */
function asImageDelete(index) {
	// 삭제할 임시파일 경로
	var delFileName = $("form[name='ASForm'] .photo-list .img").eq(index - 1).find("img").attr("src");
	var splitFileName = delFileName.split("/");
	var filepath = delFileName.replace(splitFileName[splitFileName.length - 1], "");
	delFileName = splitFileName[splitFileName.length - 1];

	// 임시 이미지 삭제처리
	$.ajax({
		type: "post",
		url: "/ASP/Mypage/Ajax/AsImageTempDelete.asp",
		async: true,
		data: "FileName=" + delFileName,
		dataType: "text",
		success: function (data) {
			var splitData = data.split("|||||");
			var result = splitData[0];
			var cont = splitData[1];

			if (result == "OK") {
				// 미리보기 이미지 삭제
				$("form[name='ASForm'] .photo-list .img").eq(index - 1).remove();

				// 업로드할 이미지 리스트 작성
				var uploadFiles = "";
				$("form[name='ASForm'] .photo-list .img").each(function () {
					var imageUrl = $(this).find("img").attr("src");
					var splitImageUrl = imageUrl.split("/");
					if (uploadFiles == "") {
						uploadFiles = splitImageUrl[splitImageUrl.length - 1];
					}
					else {
						uploadFiles = uploadFiles + "|||||" + splitImageUrl[splitImageUrl.length - 1];
					}
				});

				$("form[name='ASForm'] input[name='UploadFiles']").val(uploadFiles);


				// 업로드할 이미지 수
				var i = parseInt($("form[name='ASForm'] input[name='UploadFilesCount']").val()) - 1;
				$("form[name='ASForm'] input[name='UploadFilesCount']").val(i);

				reAsImagePreView(filepath);
			}
		},
		error: function (data) {
			//alert(data.responseText);
			//common_msgPopOpen("", "처리 도중 오류가 발생하였습니다.", "", "msgPopup", "N");
		}
	});
}

/* A/S 선택 이미지 미리보기 */
function asImagePreView(imagePath, imageName) {
	var i = parseInt($("form[name='ASForm'] input[name='UploadFilesCount']").val()) + 1;
	var html = "<li class=\"photo-list\">";
	html = html + "<button type=\"button\" onclick=\"asImageDelete(" + i + ")\" >삭제</button>";
	html = html + "<div class=\"img\">";
	html = html + "<img src=\"" + imagePath + imageName + "\" alt=\"후기 이미지\">";
	html = html + "</div>";
	html = html + "</li>";

	$("form[name='ASForm'] .as-photo").append(html);

	var uploadFiles = $("form[name='ASForm'] input[name='UploadFiles']").val().trim();

	if (uploadFiles == "") {
		uploadFiles = imageName;
	}
	else {
		uploadFiles = uploadFiles + "|||||" + imageName;
	}

	$("form[name='ASForm'] input[name='UploadFiles']").val(uploadFiles);
	$("form[name='ASForm'] input[name='UploadFilesCount']").val(i);
}

/* A/S 선택 이미지 미리보기 재배치 */
function reAsImagePreView(filepath) {
	$("form[name='ASForm'] .as-photo").html("");
	var i = parseInt($("form[name='ASForm'] input[name='UploadFilesCount']").val());
	var UploadFiles = $("form[name='ASForm'] input[name='UploadFiles']").val();
	if (i == 1) {
		var html = "<li class=\"photo-list\">";
		html = html + "<button type=\"button\" onclick=\"asImageDelete(" + i + ")\" >삭제</button>";
		html = html + "<div class=\"img\">";
		html = html + "<img src=\"" + filepath + UploadFiles + "\" alt=\"후기 이미지\">";
		html = html + "</div>";
		html = html + "</li>";
		$("form[name='ASForm'] .as-photo").append(html);
	} else {
		var UploadFilesArr = UploadFiles.split("|||||");
		for (var k = 1; k <= i; k++) {
			var html = "<li class=\"photo-list\">";
			html = html + "<button type=\"button\" onclick=\"asImageDelete(" + k + ")\" >삭제</button>";
			html = html + "<div class=\"img\">";
			html = html + "<img src=\"" + filepath + UploadFilesArr[k - 1] + "\" alt=\"후기 이미지\">";
			html = html + "</div>";
			html = html + "</li>";
			$("form[name='ASForm'] .as-photo").append(html);
		}
	}
}

