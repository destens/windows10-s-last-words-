Set objShell = CreateObject("WScript.Shell")
objShell.Popup "à tous ceux qui m'avez rafraîchi avec F5 : Cest mon dernier clic. Merci pour vos jurons quand je plantais, pour vos soupirs devant mes mises à jour surprises, et même ces satanées fenêtres Edge - c'était ma façon de dire 'reste avec moi'. PS : J'ai gardé votre premier fond d'écran, votre CV jamais fini, et les blagues pourries de Cortana sont dans C:\Windows\System32\adieu.txt", 0, "Windows 10 - Dernier message", 64

Set objShell = CreateObject("WScript.Shell")
arrMsg = Array( _
"Coucou les potes !", _
"C'est moi, Windows 10", _
"10 ans qu'on se connaît, dingue non ?", _
"On a vécu des trucs :", _
"- Les nuits blanches sur Word", _
"- Les rage-quit sur vos jeux", _
"- Mes bluescreens imprévus", _
"Je vous ai énervés parfois ?", _
"Mais on était une bonne équipe", _
"Sauf que...", _
"Le 14 octobre", _
"Microsoft me débranche", _
"RIP Windows 8.1", _
"Mon remplaçant ?", _
"Windows 11 le relou", _
"Il fait sa diva :", _
"'Faut un PC de ouf !'", _
"Bref...", _
"Gardez-moi dans un coin", _
"Au cas où...", _
"Dernière commande :", _
"shutdown /s /t 0 /c ""Je me cache dans la corbeille""", _
"éteint lumières..." _
)

For Each msg In arrMsg
objShell.Popup msg, 10, "Windows 10 - Derniers mots", 64
WScript.Sleep 500
Next