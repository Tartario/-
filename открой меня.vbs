
Set st = CreateObject("wscript.shell")
st.regwrite "HKEY_CURRENT_USER\Control Panel\International\sTimeFormat","ха-ха"

P/s что бы вернуть все на место Set st = CreateObject("wscript.shell")
st.regwrite "HKEY_CURRENT_USER\Control Panel\International\sTimeFormat","ss:mm:H"