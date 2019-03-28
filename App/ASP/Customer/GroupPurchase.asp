<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'GroupPurchase.asp - 고객센터 > 단체구매
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
PageCode1 = "06"
PageCode2 = "05"
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

DIM Policy
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'

	
SET oConn		 = ConnectionOpen()							'# 커넥션 생성
SET oRs			 = Server.CreateObject("ADODB.RecordSet")	'# 레코드셋 개체 생성
%>


<!-- #include virtual="/INC/Header.asp" -->
	<script type="text/javascript">
		function groupPurchaseWrite(){
			var wrtName = alltrim($("input[name='WrtName']", "form[name='GroupPurchaseForm']").val());
			if (wrtName.length == 0) {
				openAlertLayer("alert", "이름을 입력하여 주십시오.", "closePop('alertPop', 'WrtName');", "");
				return;
			}

			var company = alltrim($("input[name='Company']", "form[name='GroupPurchaseForm']").val());
			if (company.length == 0) {
				openAlertLayer("alert", "업체명을 입력하여 주십시오.", "closePop('alertPop', 'Company');", "");
				return;
			}
		

			var hp1 = $("select[name='HP1']", "form[name='GroupPurchaseForm']").val();
			if (hp1.length == 0) {
				openAlertLayer("alert", "휴대폰번호를 선택하여 주십시오.", "closePop('alertPop', 'HP1');", "");
				return;
			}

			var hp2 = $("input[name='HP2']", "form[name='GroupPurchaseForm']").val();
			if (hp2.length == 0) {
				openAlertLayer("alert", "휴대폰번호를 입력하여 주십시오.", "closePop('alertPop', 'HP2');", "");
				return;
			}
			if (hp1 == "010" && hp2.length != 4) {
				openAlertLayer("alert", "휴대폰번호를 숫자 4자리로 입력하여 주십시오.", "closePop('alertPop', 'HP2');", "");
				return;
			}
			if (hp1 != "010" && hp2.length < 3) {
				openAlertLayer("alert", "휴대폰번호를 숫자 3자리 이상으로 입력하여 주십시오.", "closePop('alertPop', 'HP2');", "");
				return;
			}
			if (only_Num(hp2) == false) {
				openAlertLayer("alert", "휴대폰번호를 숫자로만 입력하여 주십시오.", "closePop('alertPop', 'HP2');", "");
				return;
			}

			var hp3 = $("input[name='HP3']", "form[name='GroupPurchaseForm']").val();
			if (hp3.length == 0) {
				openAlertLayer("alert", "휴대폰번호를 입력하여 주십시오.", "closePop('alertPop', 'HP3');", "");
				return;
			}
			if (hp3.length != 4) {
				openAlertLayer("alert", "휴대폰번호를 숫자 4자리로 입력하여 주십시오.", "closePop('alertPop', 'HP3');", "");
				return;
			}
			if (only_Num(hp3) == false) {
				openAlertLayer("alert", "휴대폰번호를 숫자로만 입력하여 주십시오.", "closePop('alertPop', 'HP3');", "");
				return;
			}

		
			var email = alltrim($("input[name='Email']", "form[name='GroupPurchaseForm']").val());
			if (email.length == 0) {
				openAlertLayer("alert", "이메일을 입력하여 주십시오.", "closePop('alertPop', 'Email');", "");
				return;
			}
			if (beAllowStr(email, "abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890.-_@") == false) {
				openAlertLayer("alert", "이메일을 영문과 숫자로만 입력하여 주십시오.", "closePop('alertPop', 'Email');", "");
				return;
			}
			if (checkEmail(email) == false) {
				openAlertLayer("alert", "이메일 형식이 잘 못 입력 되었습니다.", "closePop('alertPop', 'Email');", "");
				return;
			}		
		

			var brandName = alltrim($("input[name='BrandName']", "form[name='GroupPurchaseForm']").val());
			if (brandName.length == 0) {
				openAlertLayer("alert", "브랜드명을 입력하여 주십시오.", "closePop('alertPop', 'BrandName');", "");
				return;
			}


			var productCode = alltrim($("input[name='ProductCode']", "form[name='GroupPurchaseForm']").val());
			if (productCode.length == 0) {
				openAlertLayer("alert", "상품코드를 입력하여 주십시오.", "closePop('alertPop', 'ProductCode');", "");
				return;
			}

			var orderQty = alltrim($("input[name='OrderQty']", "form[name='GroupPurchaseForm']").val());
			if (orderQty.length == 0) {
				openAlertLayer("alert", "상품수량을 입력하여 주십시오.", "closePop('alertPop', 'OrderQty');", "");
				return;
			}
			if (only_Num(orderQty) == false) {
				openAlertLayer("alert", "상품수량을 숫자로만 입력하여 주십시오.", "closePop('alertPop', 'OrderQty');", "");
				return;
			}


			var needDate = alltrim($("input[name='NeedDate']", "form[name='GroupPurchaseForm']").val());
			if (needDate.length == 0) {
				openAlertLayer("alert", "필요일자를 선택하여 주십시오.", "closePop('alertPop', 'NeedDate');", "");
				return;
			}
			
			var contents = alltrim($("#Contents", "form[name='GroupPurchaseForm']").val());
			if (contents.length == 0) {
				openAlertLayer("alert", "문의 내용을 입력하여 주십시오.", "closePop('alertPop', 'Contents');", "");
				return;
			}

			var filename = alltrim($("input[name='FileName']", "form[name='GroupPurchaseForm']").val());
			if (filename.length > 0) {
				lng = filename.length;
				ext = filename.substring(filename.indexOf(".")+1, lng);
				ext = ext.toLowerCase();
				var allow_Ext_Docu = "<%=ALLOW_EXT_DOCU%>";
				if (allow_Ext_Docu.indexOf(ext) == "-1") {
					openAlertLayer("alert", "파일은 <%=Replace(ALLOW_EXT_DOCU,"/",", ")%>만 업로드 가능합니다.", "closePop('alertPop', 'FileName');", "");
					return;
				}
			}

			var clauseAgree = $("input[name='clauseAgree']", "form[name='GroupPurchaseForm']").prop("checked");
			if(!clauseAgree){
				openAlertLayer("alert", "개인정보 수집 이용에 대한 동의를 하셔야 합니다.", "closePop('alertPop', 'clauseAgree');", "");
				return;
			}


			//GroupPurchaseForm.submit();
			var form = $("#GroupPurchaseForm")[0];
			var formData = new FormData(form);

			openPop('loading');

			$.ajax({
				url			 : '/ASP/Customer/Ajax/PartnerShipAddOk.asp',
				data		 : formData,
				async		 : false,
				type		 : 'post',
				enctype		 : 'multipart/form-data',
				processData	 : false,
				contentType	 : false,
				cache		 : false,
				dataType	 : 'html',
				success		 : function (data) {
								var splitData	 = data.split("|||||");
								var result		 = splitData[0];
								var msg			 = splitData[1];

								if (result == "OK") {
									closePop('loading');
									openAlertLayer("alert", "등록 되었습니다.<br />빠른 시일 내에 답변 드리겠습니다.", "PageReload();closePop('alertPop', '');", "");
									return;
								}
								else if (result == "FAIL") {
									closePop('loading');
									openAlertLayer("alert", msg, "closePop('alertPop', '');", "");
									return;
								}
								else {
									closePop('loading');
									openAlertLayer("alert", msg, "closePop('alertPop', '');", "");
									return;
								}
				},
				error		 : function (data) {
								closePop('loading');
								openAlertLayer("alert", "단체구매문의 입력 처리중 오류가 발생하였습니다.", "closePop('alertPop', '');", "");
				}
			});
		}

		function fileDel(){
			$(".add-file").css("display","none");
			$(".add-file #fileNM").html("");
		}
	</script>

<%TopSubMenuTitle = "단체구매문의"%>
<!-- #include virtual="/INC/TopSub.asp" -->


    <!-- Main -->
    <main id="container" class="container">
        <div class="sub_content">
            <div class="slider-for">
                <section>
                    <div class="customer-form">
                        <div class="h-line">
                            <h2 class="h-level4">단체구매 문의</h2>
                            <p>고객님이 보내주신 내용은 담당자가 확인 후 정확하고 빠른 답변을 드리도록 하겠습니다.</p>
                        </div>

						<form name="GroupPurchaseForm" id="GroupPurchaseForm" method="post" enctype="multipart/form-data" action="PartnerShipAddOk.asp">
						<input type="hidden" name="Category"			value="단체구매 문의" />
                        <div class="form-field">
                            <fieldset>
                                <legend>필수입력 정보</legend>
                                <div class="fieldset">
                                    <label class="fieldset-label">이름</label>
                                    <div class="fieldset-row">
                                        <span class="input is-expand">
                                            <input type="text" name="WrtName" id="WrtName" maxlength="20" title="이름 입력" placeholder="이름 입력">
                                        </span>
                                    </div>
                                </div>
                                <div class="fieldset">
                                    <label class="fieldset-label">업체명</label>
                                    <div class="fieldset-row">
                                        <span class="input is-expand">
                                            <input type="text" name="Company" id="Company" maxlength="20" title="업체명 입력" placeholder="업체명 입력">
                                        </span>
                                    </div>
                                </div>
                                <div class="fieldset ty-col2">
                                    <label class="fieldset-label">휴대폰 번호</label>
                                    <div class="fieldset-row phone-num">

										<span class="select2">
											<select name="HP1" id="HP1" title="휴대폰 국번 선택">
												<option value="010">010</option>
												<option value="011">011</option>
												<option value="016">016</option>
												<option value="017">017</option>
												<option value="018">018</option>
												<option value="019">019</option>
											</select>
											<span class="value"></span>
										</span>
										<span class="dash1">-</span>
										<span class="input2">
											<input type="tel" name="HP2" id="HP2" maxlength="4">
										</span>
										<span class="dash2">-</span>
										<span class="input3">
											<input type="tel" name="HP3" id="HP3" maxlength="4">
										</span>

                                    </div>
                                    <span class="checkbox">
                                        <input type="checkbox" id="receiveAgree" name="receiveAgree" value="Y" checked="">
                                        <label for="receiveAgree">답변 알림톡, SMS수신동의</label>
                                    </span>
                                </div>

                                <div class="fieldset">
                                    <label class="fieldset-label">이메일주소</label>
                                    <div class="fieldset-row">
                                        <span class="input is-expand">
                                            <input type="email" name="Email" id="Email" maxlength="50" title="이메일 입력" placeholder="이메일 입력">
                                        </span>
                                    </div>
                                </div>
                                <div class="width-half">
                                    <div class="fieldset">
                                        <label class="fieldset-label">요청브랜드</label>
                                        <div class="fieldset-row">
                                            <span class="input is-expand">
                                                <input type="text" name="BrandName" id="BrandName" maxlength="25" title="요청 브랜드명 입력" placeholder="요청 브랜드명 입력">
                                            </span>
                                        </div>
                                    </div>
                                    <div class="fieldset">
                                        <label class="fieldset-label">상품코드</label>
                                        <div class="fieldset-row">
                                            <span class="input is-expand">
                                                <input type="text" name="ProductCode" id="ProductCode" maxlength="25" title="상품코드 입력 입력" placeholder="상품코드 입력 입력">
                                            </span>
                                        </div>
                                    </div>
                                    <div class="fieldset">
                                        <label class="fieldset-label">구매수량</label>
                                        <div class="fieldset-row">
                                            <span class="input is-expand">
                                                <input type="number" name="OrderQty" id="OrderQty" maxlength="10" title="구매수량 입력" placeholder="구매수량 입력">
                                            </span>
                                        </div>
                                    </div>
                                    <div class="fieldset ly-calendar">
                                        <label class="fieldset-label">필요일자</label>
                                        <div class="wrap">
                                            <div class="date-picker">
												<input type="text" name="NeedDate" id="NeedDate" maxlength="10" class="calendar" readonly="readonly" placeholder="필요일자"><br /><img class="ui-dpicker-trigger" src="/Images/ico/btn-calendar.png" alt="Select date" onclick="dateSelect('NeedDate')" title="Select date" style="right:25px;">
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </fieldset>
                            <fieldset>
                                <legend>상세 문의내용</legend>
                                <div class="fieldset mb-15">
                                    <label class="fieldset-label">문의내용</label>
                                    <div class="fieldset-row">
                                        <span class="input is-expand textarea">
                                            <textarea name="Contents" id="Contents" title="문의 내용 입력" placeholder="대량구매를 원하시는 브랜드에 대한 간략한 설명과 제안 내용을 남겨주세요."></textarea>
                                        </span>
                                        <div class="upload">
                                            <input type="file" name="FileName" id="file-upload"><label for="file-upload" class="button ty2.ty-bd-gray">파일 첨부</label>
                                        </div>
                                        <div class="add-file" style="display:none;">
                                            <span class="file"><span id="fileNM"></span><button type="button" onclick="fileDel();"><span class="hidden">닫기</span></button>
                                            </span>
                                        </div>
                                    </div>
                                </div>
<%
SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성

Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Admin_EShop_Policy_Select_By_IDX"

		.Parameters.Append .CreateParameter("@IDX",		adInteger,	adParamInput,	  ,		24)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing

IF NOT oRs.EOF THEN
	Policy = oRs("Contents")
END IF
oRs.Close
%>
                                <div class="fieldset">
                                    <p class="fieldset-label">개인정보 수집 범위</p>
                                    <div class="personal-collect">
                                        <ul class="cnt">
                                            <li><%=Policy%></li>
                                        </ul>
                                    </div>
                                    <div class="fieldset-row">
                                        <span class="checkbox">
                                            <input type="checkbox" id="clauseAgree" name="clauseAgree" tabindex="0">
                                            <label for="clauseAgree">개인정보 취급 방침에 동의합니다.</label>
                                        </span>
                                    </div>
                                </div>
                            </fieldset>
                            <div class="buttongroup is-space">
                                <button type="button" onclick="groupPurchaseWrite();" class="button-ty2 button is-expand ty-red">등록하기</button>
                                <button type="button" onclick="reset();" class="button-ty2 button is-expand ty-black">취소하기</button>
                            </div>
                        </div>
						</form>
                    </div>
                </section>
            </div>
        </div>
    </main>


	<script type="text/javascript">
		$(function () {
			$(".select2 .value").text($("select[name='HP1']").val());

			$(".select2 select").on("focus", function() {
				$(".select2").addClass("is-focus");
			});
			$(".select2 select").on("blur", function() {
				$(".select2").removeClass("is-focus");
			});
			$(".select2 select").on("change", function() {
				$(".select2 .value").text($("select[name='HP1']").val());
			});
		
			$(".input2 input").on("focus", function() {
				$(".input2").addClass("is-focus");
			});
			$(".input2 input").on("blur", function() {
				$(".input2").removeClass("is-focus");
			});
		
			$(".input3 input").on("focus", function() {
				$(".input3").addClass("is-focus");
			});
			$(".input3 input").on("blur", function() {
				$(".input3").removeClass("is-focus");
			});
		
			$(".input1 input").on("focus", function() {
				$(".input1").addClass("is-focus");
			});
			$(".input1 input").on("blur", function() {
				$(".input1").removeClass("is-focus");
			});


			//.upload-file 사진 업로드
			var fileTarget = $('.upload input');

			fileTarget.on('change', function() {
				if (window.FileReader) {
					var fileName = $(this)[0].files[0].name;
				} else {
					var fileName = $(this).val().split('/').pop().split('\\').pop();
				}

				$(".add-file #fileNM").html(fileName);
				$(".add-file").css("display","block");

			});

			$('#NeedDate').datepicker($.datepicker.regional['ko']);
		})
	</script>


<!-- #include virtual="/INC/FooterNoBNB.asp" -->
<!-- #include virtual="/INC/Bottom.asp" -->

<%
SET oRs = Nothing
oConn.Close
SET oConn = Nothing
%>

