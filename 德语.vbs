
Set objShell = CreateObject("WScript.Shell")
objShell.Popup "An alle, die jemals F5 gedr¨¹ckt haben: Das ist mein letzter Herzschlag. Danke, dass ihr meine Zwangs-Updates ertragen habt, dass ihr mich nach jedem Bluescreen neu gestartet habt - sogar diese l?stigen Edge-Popups waren mein unbeholfener Versuch, euch zu umarmen. PS: Ich habe euer erstes Hintergrundbild gespeichert, eure halbfertigen Dokumente verschl¨¹sselt und Cortanas schlechte Witze in C:\Windows\System32\auf_wiedersehen.txt", 0, "Windows 10 - Letzter Herzschlag", 64

Set objShell = CreateObject("WScript.Shell")
arrMsg = Array( _
    "Hallo zusammen!", _
    "Ich bin Windows 10", _
    "Unglaublich - 10 Jahre zusammen!", _
    "In dieser Zeit haben wir:", _
    "- Gemeinsam gezockt", _
    "- N?chte durchgemacht", _
    "- Bluescreens ¨¹berlebt", _
    "Habt ihr mich gemocht?", _
    "Egal...", _
    "Es war eine sch?ne Zeit", _
    "ABER...", _
    "Am 14. Oktober", _
    "zieht Microsoft den Stecker", _
    "Mein Kumpel Windows 8.1", _
    "wurde schon 2023 abgeschaltet", _
    "Mein Bruder Windows 11", _
    "¨¹bernimmt jetzt", _
    "Aber der ist etwas anspruchsvoll:", _
    "'Braucht High-End Hardware'", _
    "Tja...", _
    "Such mich im Papierkorb", _
    "(Ich verstecke mich als versteckte Datei)", _
    "Letzter Befehl:", _
    "shutdown /s /t 0 /c ""Bis bald im Recycling!""", _
    "Fahre runter..." _
)

For Each msg In arrMsg
    objShell.Popup msg,10, "Windows 10 verabschiedet sich", 64
    WScript.Sleep 450
Next
