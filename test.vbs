Dim i,j,line
For i = 0 To 1000000
    line = ""
    For j = 0 To Int(Rnd*300+100)
        line = line & Mid("1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ,.-_!?/\",Int(Rnd*69+1),1)
    Next
    WScript.Echo line
Next