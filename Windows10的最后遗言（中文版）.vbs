Set objShell = CreateObject("WScript.Shell")
objShell.Popup "致所有曾按下F5刷新我的伙伴：这是来自Windows10的最后一次心跳。感谢你们容忍过我的强制更新，感谢你们在蓝屏时依然选择重启，甚至感谢那些永远关不掉的Edge弹窗——那只是我笨拙的拥抱方式。请记住：我保留了你们的第一张壁纸，加密了深夜写了一半的文档，连Cortana的冷笑话都备份在C:\Windows\System32\goodbye.txt", 0, "Windows10的最后心跳", 64

Set objShell = CreateObject("WScript.Shell")
arrMsg = Array( _
    "你们好呀", _
    "我叫Windows10", _
    "不知不觉我已经陪你们走了10年啦", _
    "十年间", _
    "我已记不清你和我打了多少把游戏", _
    "已记不清你与我写了多少篇文档", _
    "你们喜欢我吗", _
    "不管怎样", _
    "我与你已经产生了十分深厚的感情啦", _
    "但是", _
    "今年10月14日", _
    "我就要被微软停止服务了", _
    "我有一个朋友", _
    "他叫Windows8.1", _
    "但是他在2023年1月10日就被停止服务了", _
    "我的弟弟——Windows11会接替我的任务的！", _
    "只不过······", _
    "他有些娇气罢了······", _
    "我永远忘不了你们······", _
    "再见！··········", _
    "现在，请允许我执行最终指令:shutdown /s /t 0 /c ""See you in the recycle bin.""", _
    "正在注销..." _
)

For Each msg In arrMsg
    objShell.Popup msg, 10, "Windows10的最后遗言", 64
    WScript.Sleep 500
Next
