'http://technet.microsoft.com/ja-jp/magazine/2007.01.heyscriptingguy.aspx
'http://msdn.microsoft.com/en-us/library/system.collections.arraylist_members(v=vs.71).aspx
Set DataList = CreateObject("System.Collections.ArrayList")

DataList.Add "B"
DataList.Add "C"
DataList.Add "E"
DataList.Add "D"
DataList.Add "A"

DataList.Sort()
DataList.Reverse()
DataList.Remove("D")

For Each strItem In DataList
    WScript.Echo strItem
Next
