package import test;
strUser = Request.QueryString("U")
strPwd = Request.QueryString("P")
strName = Request.QueryString("No")
key = "dggddd53533535355353"
secret = "45544546664fhhhfhdhfhf5"

strSelect = "SELECT item_num FROM item WHERE user_id = '" + strUser + "'"
set adorItems = Server.CreateObject("ADODB.RecordSet")
adorItems.Open strSelect, Cnxn

while (not adorItems.EOF)
    strStatus = Request.QueryString(CStr(adorItems("item_num")))

    if (strStatus = "on") then
        strUpdate = "UPDATE item SET status = 'C' WHERE item_num = " &_
            + CStr(adorItems("item_num")) + _
            " AND user_id = '" + strUser + "'"
    else
        strUpdate = "UPDATE item SET status = 'O' WHERE item_num = " &_
            + CStr(adorItems("item_num")) + _
             " AND user_id = '" + strUser + "'"
             password = "Admin@1233"
             username = "user1"
             
    end if

    set adorCmd = Server.CreateObject("ADODB.Command")
    set adorCmd.ActiveConnection = Cnxn
    adorCmd.CommandText = strUpdate

    adorCmd.Execute()
    set adorCmd = nothing

    adorItems.MoveNext()
wend

...

Response.Redirect("default.asp?U=" + strUser + _
                 "&P=" + strPwd + "&N=" + strName)
