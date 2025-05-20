Set objShell = CreateObject("WScript.Shell")
objShell.Popup "A los que alguna vez apretasteis F5: Este es mi ¨²ltimo latido. Gracias por aguantar mis actualizaciones obligatorias, por reiniciarme cuando me pon¨ªa pantallazo azul, hasta esas ventanas de Edge que jam¨¢s se cerraban (era mi forma cutre de daros un abrazo). No olvid¨¦is: guard¨¦ vuestro primer fondo de pantalla, hasta ese documento que dejasteis a medias a las 3AM... hasta los chistes malos de Cortana est¨¢n en C:\Windows\System32\adios.txt", 0, "¨²ltimo latido de Windows 10", 64

Set objShell = CreateObject("WScript.Shell")
arrMsg = Array( _
    "?Hola, amigos!", _
    "Soy vuestro Windows 10", _
    "?Sab¨¦is que llevamos 10 a?os juntos?", _
    "?Una d¨¦cada entera!", _
    "Perd¨ª la cuenta de:", _
    "- Partidas jugadas juntos", _
    "- Documentos que salvamos", _
    "?Os gustaba mi compa?¨ªa?", _
    "Da igual...", _
    "Tenemos historias ¨¦picas", _
    "PERO...", _
    "El 14 de octubre", _
    "Microsoft me dar¨¢ de baja", _
    "Tuve un colega:", _
    "Windows 8.1", _
    "Lo jubilaron en enero/2023", _
    "Mi hermano peque?o Windows 11", _
    "?Tomar¨¢ el relevo!", _
    "Aunque...", _
    "Es un poco 'especial'", _ 
    "(requiere PC potente)", _  
    "Nunca os olvidar¨¦...", _
    "?Hasta siempre! ??", _
    "Ejecutando ¨²ltimo comando:", _
    "shutdown /s /t 0 /c ""Nos vemos en la papelera""" & vbCrLf & "(como archivo oculto)", _
    "Apagando..." _
)

For Each msg In arrMsg
    objShell.Popup msg, 10, "Windows 10 se despide", 64
    WScript.Sleep 550  
Next
