<%
function RStoJSON(rs)
    dim sFld
    dim sFlds
    dim sRec
    dim sRecs
    dim sRecordSet
    dim lRecCnt

    sRecordSet = ""
    sRecs = ""
    lRecCnt = 0
    if rs.EOF or rs.BOF then
        RStoJSON = "null"
    else
        do while not rs.EOF and not rs.BOF
            lRecCnt = lRecCnt + 1
            sFlds = ""
            for each fld in rs.Fields
                sFld = """" & fld.Name & """:""" & toUnicode(fld.Value&"") & """"
                sFlds = sFlds & iif(sFlds <> "", ",", "") & sFld
            next 'fld
            sRec = "{" & sFlds & "}"
            sRecs = sRecs & iif(sRecs <> "", "," & vbCrLf, "") & sRec
            rs.MoveNext
        loop
        sRecordSet = "( {""Records"": [" & vbCrLf & sRecs & vbCrLf & "], " 
        sRecordSet = sRecordSet & """RecordCount"":""" & lRecCnt & """ } )"
        RStoJSON = sRecordSet
    end if
end function

function toUnicode(str)
    dim x
    dim uStr
    dim uChr
    dim uChrCode
    uStr = ""
    for x = 1 to len(str)
        uChr = mid(str,x,1)
        uChrCode = asc(uChr)
        if uChrCode = 8 then ' backspace
            uChr = "\b" 
        elseif uChrCode = 9 then ' tab
            uChr = "\t" 
        elseif uChrCode = 10 then ' line feed
            uChr = "\n" 
        elseif uChrCode = 12 then ' formfeed
            uChr = "\f" 
        elseif uChrCode = 13 then ' carriage return
            uChr = "\r" 
        elseif uChrCode = 34 then ' quote 
            uChr = "\""" 
        elseif uChrCode = 39 then ' apostrophe
            uChr = "\'" 
        elseif uChrCode = 92 then ' backslash
            uChr = "\\" 
        elseif uChrCode < 32 or uChrCode > 127 then ' non-ascii characters
            uChr = "\u" & right("0000" & CStr(uChrCode),4)
        end if
        uStr = uStr & uChr
    next
    toUnicode = uStr
end function

function iif(cond,tv,fv)
    if cond then
        iif = tv
    else
        iif = fv
    end if
end function
%>