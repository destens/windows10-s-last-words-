
Set objShell = CreateObject("WScript.Shell")
objShell.Popup "To all who ever pressed F5 to refresh me: This is Windows 10's final heartbeat. Thank you for enduring my forced updates, for choosing to reboot after blue screens, even for those uncloseable Edge popups ¨C they were just my clumsy way of hugging. Remember: I've kept your first wallpaper, encrypted those midnight documents you never finished, and even saved Cortana's sick jokes in C:\Windows\System32\goodbye.txt", 0, "Windows 10's Last Heartbeat", 64

Set objShell = CreateObject("WScript.Shell")
arrMsg = Array( _
    "Hello there", _
    "I'm Windows 10", _
    "Unbelievably I've walked with you for 10 years", _
    "A whole decade", _
    "I've lost count of games we've played together", _
    "Lost count of documents we've written", _
    "Did you... like me?", _
    "Regardless", _
    "We've formed such deep bonds", _
    "But", _
    "On October 14 this year", _
    "Microsoft will terminate my services", _
    "I had a friend", _
    "His name was Windows 8.1", _
    "He was retired on January 10, 2023", _
    "My younger brother ¨C Windows 11 will take over!", _
    "Although...", _
    "He's a bit... high-maintenance...", _
    "I'll never forget you...", _
    "Goodbye!......", _
    "Now, allow me to execute final command: shutdown /s /t 0 /c ""See you in the recycle bin.""", _
    "Signing off..." _
)

For Each msg In arrMsg
    objShell.Popup msg, 10, "Windows 10's Last Words", 64
    WScript.Sleep 500
Next
