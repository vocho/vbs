Set dicFolderName = CreateObject("Scripting.Dictionary")
With dicFolderName
    .Add "key1", "abc"
    .Add "key2", "def"
End With

For Each strKey In dicFolderName.Keys
    WScript.Echo strKey
    WScript.Echo dicFolderName.Item(strKey)
Next
For Each strItem In dicFolderName.Items
    WScript.Echo strItem
Next

Set dicFolderName = Nothing
