<%@ Language=VBScript CodePage="65001" %>
<%Option Explicit%>
<%
'*****************************************************************************************'
'messagePopup.asp - 공통팝업
'Date		: 2018.12.13
'Update		:
'*****************************************************************************************'

'# 페이지 응답헤더 설정-------------------------------------------------------------------'
Response.Expires = -1
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma","no-cache"
Response.CharSet = "utf-8"
'-----------------------------------------------------------------------------------------'

'*****************************************************************************************'
'변수 선언 START
'-----------------------------------------------------------------------------------------'
DIM Title
DIM Msg
DIM Script
DIM tempScript
DIM Focus
DIM pStyle
'-----------------------------------------------------------------------------------------'
'변수 선언 END
'-----------------------------------------------------------------------------------------'
Title		 = Trim(Request("title"))
Msg			 = Trim(Request("msg"))
Script		 = Trim(Request("Script"))
Focus		 = Trim(Request("Focus"))
pStyle		 = Trim(Request("pStyle"))


IF Title = "" THEN
	Title = "SHOEMARKER"
END IF

IF Script = "undefined" THEN
	Script = ""
END IF

IF Script <> "" THEN
	tempScript = MID(Script, LEN(Script)-1, 1)
	IF tempScript <> ";" THEN
		tempScript = ";"
	END IF
END IF

IF Focus = "undefined" THEN
	Focus = ""
END IF

IF pStyle = "undefined" THEN
	pStyle = "N"
END IF

Response.Write "OK|||||"
%>
        <div class="area-dim" style="z-index:101"></div>

        <div class="area-pop">
            <div class="alert">
                <div class="tit-pop">
                    <p class="tit"><%=Title%></p>
                    <!--<button class="btn-hide-pop" onclick="common_msgPopClose('<%=Focus%>');">닫기</button>-->
                </div>

                <div class="container-pop">
                    <div class="contents">
                        <div class="ly-cont">
                            <p class="t-level4"><%=Msg%></p>
                        </div>
                    </div>
                    <div class="btns">
                        <%IF pStyle="C" THEN%><button type="button"  onclick="common_msgPopClose('<%=Focus%>');" class="button ty-black">취소</button><%END IF%>
                        <button type="button"  onclick="<%Response.Write "common_msgPopClose('"& Focus &"');" & Script & tempScript%>" class="button ty-red">확인</button>
                    </div>
                </div>
            </div>
        </div>

