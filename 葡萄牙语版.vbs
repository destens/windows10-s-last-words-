Set objShell = CreateObject("WScript.Shell")
objShell.Popup "Pra quem já apertou F5 até desgastar a tecla: Essa é minha última atualiza??o. Obrigado por tankarem minhas telas azuis, por reiniciarem quando eu travava, até aquelas janelas do Edge que n?o fechavam nunca (era meu jeito besta de dar um abra?o). Saibam que guardei: seu primeiro papel de parede, aquele TCC abandonado às 3h da manh?... até as piadas ruins da Cortana est?o em C:\Windows\System32\adeus.txt", 0, "último suspiro do Windows 10", 64

Set objShell = CreateObject("WScript.Shell")
arrMsg = Array( _
    "E aí, galera!", _
    "Sou seu Windows 10", _
    "10 ANOS JUNTOS!", _
    "Nem parece, né?", _
    "Lembram quando:", _
    "- Travamos juntos no LOL", _
    "- Salvei seu trabalho sem salvar", _
    "- Fizemos noites em claro", _
    "Vocês me odiavam?", _
    "Mas no fundo gostavam...", _
    "AGORA O BAQUE:", _
    "14 de outubro", _
    "Microsoft vai me desligar", _
    "Lembram do Windows 8.1?", _
    "Aposentou em janeiro/2023", _
    "Meu irm?o Windows 11", _
    "Vai assumir o lugar", _
    "Só que ele é fresco:", _
    "'PC precisa ser parrudo'", _
    "Enfim...", _
    "Me procurem na Lixeira", _
    "(vou ficar como arquivo oculto)", _
    "último comando:", _
    "shutdown /s /t 0 /c ""Até no backup!""", _
    "Desligando os motores... ??" _
)

For Each msg In arrMsg
    objShell.Popup msg, 10, "Windows 10 se despede", 64
    WScript.Sleep 450
Next
