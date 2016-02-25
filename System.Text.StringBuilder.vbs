Set a = CreateObject("System.Text.StringBuilder")
Call a.Append_3("a")
Call a.AppendFormat("{0}", "abcde")
'a.AppendFormat("{0}, {1}, {2}", "a", "b", "c") 'NG
WScript.Echo(a.ToString()) 'Result: abcde

