Option Explicit

Sub testMain()
    Call testFirstStrFirstDelim
    Call testFirstStrLastDelim
    Call testLastStrFirstDelim
    Call testLastStrLastDelim
    
    Call testTagInnerText
    
    Call testString_SaveToFile
    Call testString_LoadFromFile
    
    Call testString_SaveTextFile
    Call testString_LoadTextFile
End Sub
