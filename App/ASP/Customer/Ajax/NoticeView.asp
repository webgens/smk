<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'/****************************************************************************************'
'NoticeList.asp - 고객센터 > 공지사항 뷰
'Date		: 2019.01.06
'Update	: 
'/****************************************************************************************'

'//페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'//---------------------------------------------------------------------------------------'
%>

<!-- #include virtual="/ADO/ADODBCommon.asp" -->
<!-- #include virtual="/Common/Common.asp" -->
<!-- #include virtual="/Common/SetInfo.asp" -->
<!-- #include virtual="/Common/ProgID1.asp" -->

<%
'/****************************************************************************************'
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


DIM IDX
DIM TopFlag
DIM Title
DIM Contents
DIM CreateDT
DIM ViewCount
DIM FileName
'-----------------------------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------------------------'


IDX				 = sqlFilter(Request("IDX"))
IF IDX = "" THEN
	Response.Write "FAIL|||||공지사항 정보가 없습니다."
	Response.End
END IF

SET oConn = ConnectionOpen()	'//커넥션 생성
SET oRs = Server.CreateObject("ADODB.RecordSet")	'//레코드셋 개체 생성

wQuery = "WHERE A.DelFlag = 'N' AND A.IDX = '"& IDX &"' "
sQuery = "ORDER BY A.TopFlag DESC, A.CreateDT DESC "
Set oCmd = Server.CreateObject("ADODB.Command")
WITH oCmd
		.ActiveConnection = oConn
		.CommandType = adCmdStoredProc
		.CommandText = "USP_Front_EShop_Notice_Select_By_IDX"

		.Parameters.Append .CreateParameter("@IDX",		adInteger,	adParamInput,	  ,		IDX)
END WITH
oRs.CursorLocation = adUseClient
oRs.Open oCmd, , adOpenStatic, adLockReadOnly
Set oCmd = Nothing

IF oRs.EOF THEN
	Response.Write "FAIL|||||정보가 일치하지 않습니다."
ELSE
	TopFlag			= oRs("TopFlag")
	Title			= oRs("Title")
	Contents		= oRs("Contents")
	FileName		= oRs("FileName")
	CreateDT		= LEFT(oRs("CreateDT"), 10)
	ViewCount		= oRs("ViewCount")
END IF
oRs.Close

Response.Write "OK|||||"
%>

    <!-- PopUp -->
        <div class="area-dim"></div>

        <div class="area-pop">
            <div class="full">
                <div class="tit-pop">
                    <p class="tit">슈마커 소식</p>
                    <button onclick="common_PopClose('DimDepth1')" class="btn-hide-pop">닫기</button>
                </div>
                <div class="container-pop">
                    <div class="contents no-padding-top">
                        <div class="pop-customer">
                            <div class="tit-area">
                                <p><%=Title%></p>
                                <span><span>(<%=CreateDT%></span><span>조회 <%=FormatNumber(ViewCount,0)%>)</span></span>
                            </div>
                            <div class="cnt">
                                <p><%=Contents%></p>
								<%IF FileName <> "" THEN%>
									<%IF Right(FileName,3) = "gif" Or Right(FileName,3) = "PNG" Or Right(FileName,3) = "JPG" THEN%>
                                    <img src="<%=FileName%>" alt="<%=Title%>">
									<%ELSE%>
									첨부파일 : <a href="/Common/Down.asp?file=<%=MID(FileName, INSTRREV(FileName,"/")+1)%>&dtype=Notice"><%=MID(FileName, INSTRREV(FileName,"/")+1)%></a>
									<%END IF%>
								<%END IF%>
                            </div>
                        </div>
                    </div>
                    <div class="btns">
                        <button type="button" onclick="common_PopClose('DimDepth1')" class="button ty-red">닫 기</button>
                    </div>
                </div>
            </div>
        </div>
    <!-- // PopUp -->


<%
Set oRs = Nothing
oConn.Close
Set oConn = Nothing
%>