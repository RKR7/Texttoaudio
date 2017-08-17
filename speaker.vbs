Dim message, sapi
 message =InputBox("A Best Text Audio converter "+vbcrlf+"From - RKR","Text to Audio converter")
 Set sapi = CreateObject("sapi.spvoice")
 sapi.Speak message
