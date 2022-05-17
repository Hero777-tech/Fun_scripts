Dim msg, sapi 'created by Aditya-20
msg=InputBox("Enter your text for conversion                                                                                                                                                                                          Created by Shayan" ,"Simpe Text to Speech" ,"Type Here")
Set sapi=CreateObject("sapi.spvoice")
sapi.Speak msg
msgbox("Your Text has been Spoken: ") + msg
msgbox("Thank You for using, Consider visiting my Linkedin")
WScript.Sleep(5000) 'Counted in miliseconds.
Set shl = CreateObject("Wscript.shell")
shl.run "https://www.linkedin.com/in/aditya-nath-453341221/" 'Link of my Linkedin
