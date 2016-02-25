With CreateObject("ADODB.Stream")
    .Type = 2
    .Mode = adModeReadWrite
    .Open
    .Charset = "utf-8"
    .Position = .Size
    .WriteText = "abc text"
    .SaveToFile "aaaa.txt", 2
    .Close
End With