import win32com.client
import easygui as g
import winreg
#import subprocess

scope=win32com.client.Dispatch("LeCroy.ActiveDSOCtrl.1")

def get_desktop():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                          r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders',)
    return winreg.QueryValueEx(key, "Desktop")[0]

print("DESKTOP PATH: ",get_desktop())
desktop = get_desktop()
path = g.diropenbox(msg="请选择保存截图的路径", title=None)
if path != None :
    path = path +"\\"
else :
    path = desktop
print("CHOOSE PATH: ",path)
msg = "请填写如下内容"
title = "Lecroy PNG Saver By: 御坂初琴"
fieldNames = ["示波器IP","保存路径","文件名"]
fieldValues = ["", path, ""] #初始化结束
while True:
    lastip = fieldValues[0]
    lastpath = fieldValues[1]
    lastname = fieldValues[2]
    print("LASTIP: ",lastip)
    print("LASTPATH: ",lastpath)
    print("LASTNAME: ",lastname)

    fieldValues = g.multenterbox(msg,title,fieldNames,[lastip,lastpath,lastname])
    print("VALUES: ",fieldValues)
    #判断循环
    while True:
        if fieldValues == None :
            break
        errmsg = ""
        if fieldValues[0] == "" :
            errmsg = "IP值不能为空"
        if fieldValues[1] == "" :
            errmsg = "路径不能为空"
        if fieldValues[2] == "" :
            errmsg = "文件名不能为空"
        if errmsg == "":
            break
        fieldValues = g.multenterbox(errmsg,title,fieldNames,fieldValues)
    #判断循环
    if fieldValues == None :
        break
    if fieldValues[0] != lastip :
        scope.Disconnect()
        scope.MakeConnection("IP:"+fieldValues[0]) #IP地址
    scope.StoreHardcopyToFile("PNG", "", fieldValues[1] + fieldValues[2] + ".png")
    #subprocess.call('cscript ./msg.vbs')
    g.msgbox('已保存图片,路径为:' + fieldValues[1] + fieldValues[2] + ".png", title)

scope.Disconnect()