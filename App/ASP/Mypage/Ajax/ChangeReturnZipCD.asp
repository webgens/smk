<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'ChangeReturnZipCD.asp - 교환/반품지 주소 변경 폼 페이지
'Date		: 2019.01.02
'Update	: 
'/****************************************************************************************'

'//페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'//---------------------------------------------------------------------------------------'
%>

<!-- #include Virtual = "/ADO/ADODBCommon.asp" -->
<!-- #include Virtual = "/Common/Common.asp" -->
<!-- #include Virtual = "/Common/SetInfo.asp" -->

<%
'/****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM oConn											'# ADODB Connection 개체
DIM oRs												'# ADODB Recordset 개체
DIM oCmd											'# ADODB Command 개체

DIM i
DIM j
DiM x
DIM y

DIM wQuery											'# WHERE 절
DIM sQuery											'# SORT 절

DIM AddrType
DIM AddrTypeNM
DIM Name
DIM Phone
DIM Phone1
DIM Phone2
DIM Phone3
DIM ZipCode
DIM Addr1
DIM Addr2

DIM arrHP1
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


AddrType		= sqlFilter(Request("AddrType"))
Name			= sqlFilter(Request("Name"))
Phone			= sqlFilter(Request("Phone"))
ZipCode			= sqlFilter(Request("ZipCode"))
Addr1			= sqlFilter(Request("Addr1"))
Addr2			= sqlFilter(Request("Addr2"))


IF Phone <> "" THEN
		IF UBound(SPLIT(Phone,"-")) = 2 THEN
				Phone1	 = SPLIT(Phone, "-")(0)
				Phone2	 = SPLIT(Phone, "-")(1)
				Phone3	 = SPLIT(Phone, "-")(2)
		ELSEIF UBound(SPLIT(Phone,"-")) = 1 THEN
				Phone1	 = SPLIT(Phone, "-")(0)
				Phone2	 = SPLIT(Phone, "-")(1)
				Phone3	 = ""
		ELSEIF UBound(SPLIT(Phone,"-")) = 0 THEN
				Phone1	 = SPLIT(Phone, "-")(0)
				Phone2	 = ""
				Phone3	 = ""
		ELSE
				Phone1	 = Phone
				Phone2	 = ""
				Phone3	 = ""
		END IF
END IF


IF AddrType = "" THEN
		Response.Write "FAIL|||||변경할 주소구분 정보가 없습니다."
		Response.End
END IF


IF AddrType = "Return" THEN
		AddrTypeNM	= "반품상품 수거지"
ELSEIF AddrType = "Receive" THEN
		AddrTypeNM	= "교환상품 배송지"
END IF


arrHP1		= ARRAY("010", "011", "016", "017", "018", "019")


SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성


Response.Write "OK|||||"
%>					
        <div class="area-pop" id="ChangeReturnZipCD">
            <div class="full">
                <div class="tit-pop">
                    <p class="tit"><%=AddrTypeNM%> 정보변경</p>
                    <button type="button" onclick="closePop('DimDepth2')" class="btn-hide-pop">닫기</button>
                </div>

                <div class="container-pop">
                    <div class="contents delivery-info" style="border-bottom:none;">
						<div class="formfield">
							<div class="fieldset">
								<label for="Name_<%=AddrType%>" class="fieldset-label"><%IF AddrType = "Return" THEN%>신청인<%ELSE%>수취인<%END IF%></label>
								<div class="fieldset-row">
									<span class="input is-expand">
										<input type="text" name="Name" id="Name_<%=AddrType%>" value="<%=Name%>" placeholder="<%IF AddrType = "Return" THEN%>신청인<%ELSE%>수취인<%END IF%> 입력">
									</span>
								</div>
							</div>
							<div class="fieldset">
								<label for="Addr2_<%=AddrType%>" class="fieldset-label">주소</label>
								<div class="postnum">
									<button class="search-postnum" type="button" onclick="execDaumPostcode('ZipCode_<%=AddrType%>','Addr1_<%=AddrType%>','Addr2_<%=AddrType%>')"><span>우편번호 검색</span></button>
									<div class="fieldset-row delivery-num">
										<span class="input is-expand">
											<input type="text" name="ZipCode" id="ZipCode_<%=AddrType%>" value="<%=ZipCode%>" placeholder="우편번호" readonly="readonly" />
										</span>
									</div>
								</div>
								<div class="fieldset-row">
									<span class="input is-expand">
										<input type="text" name="Addr1" id="Addr1_<%=AddrType%>" value="<%=Addr1%>" placeholder="주소 입력" readonly="readonly" />
									</span>
								</div>
								<div class="fieldset-row">
									<span class="input is-expand">
										<input type="text" name="Addr2" id="Addr2_<%=AddrType%>" value="<%=Addr2%>" placeholder="나머지 주소를 입력해주세요.">
									</span>
								</div>
							</div>
							<div class="fieldset">
								<div class="fieldset ty-col2 pt0">
									<label for="Phone_<%=AddrType%>" class="fieldset-label">휴대폰번호</label>
									<div class="fieldset-row">
										<span class="select">
											<select name="Phone1" title="휴대폰 국번 선택">
                                                <option value="">선택</option>
												<%FOR i = 0 TO UBOUND(arrHP1)%>
                                                <option value="<%=arrHP1(i)%>"<%IF arrHP1(i) = Phone1 THEN%> selected="selected"<%END IF%>><%=arrHP1(i)%></option>
												<%NEXT%>
											</select>
											<span class="value"><%=Phone1%></span>
										</span>
										<span class="input">
											<input type="text" name="Phone23" id="Phone_<%=AddrType%>" value="<%=Phone2 & Phone3%>" placeholder="휴대폰의 앞 번호와 뒷 번호 입력">
										</span>
									</div>
								</div>
							</div>
						</div>
					</div>

                    <div class="btns">
                        <button type="button" onclick="setChangeReturnAddress('<%=AddrType%>')" class="button ty-red">확인</button>
                    </div>
				</div>
			</div>
		</div>

		<script type="text/javascript">
			$(function () {
				$("#ChangeReturnZipCD select").on("change", function () {
					$(this).parent().find('.value').text($('option:selected', $(this)).text());
				});
			})
		</script>

<%
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>