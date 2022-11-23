<%

Function clog(value)
    Response.Write("<script>console.log(`" & value & "`)</script>")
End Function

Function cwarn(value)
    Response.Write("<script>console.warn(`" & value & "`)</script>")
End Function

Function cerr(value)
    Response.Write("<script>console.error(`" & value & "`)</script>")
End Function

%>