
#SendKeysを使うため、System.Windows.Forms名前空間読込
Add-Type -AssemblyName System.Windows.Forms

#テキスト読み上げ用
Add-Type -AssemblyName System.Speech

$ps_speak = New-Object System.Speech.Synthesis.SpeechSynthesizer
$ps_speak.SelectVoice("Microsoft Haruka Desktop")


#読み出し時間(秒)
$ReadSecond = 60
#何ページ分繰り返すか
$CountLength = 30
#読み出す本の最大ページ数
$BookPages = 16


#音声読み上げ用ファイルを読み込み
$arr = Get-Content "C:\hogehoge\aiueo-Text1.txt"
#echo $arr


# .NET Frameworkの宣言
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")
[void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")

# Windows APIの宣言
$signature=@'
[DllImport("user32.dll",CharSet=CharSet.Auto,CallingConvention=CallingConvention.StdCall)]
public static extern void mouse_event(long dwFlags, long dx, long dy, long cButtons, long dwExtraInfo);
'@
$SendMouseClick = Add-Type -memberDefinition $signature -name "Win32MouseEventNew" -namespace Win32Functions -passThru

echo "ランダム読み出し開始"
#この間にトップ画面をkindleにする
Start-Sleep -s 60
echo "Start"


for($i=0; $i -lt $CountLength; $i++){

# 変数設定
#---A---
$x = 1888
$y = 140


# マウスカーソル移動
[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x, $y)

Start-sleep -s 1

# 左クリック
$SendMouseClick::mouse_event(0x0002, 0, 0, 0, 0);

Start-sleep -s 1

$SendMouseClick::mouse_event(0x0004, 0, 0, 0, 0);


Start-Sleep -s 5

# 変数設定
#---B---
$x = 1821
$y = 276

# マウスカーソル移動
[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x, $y)

Start-sleep -s 1

# 左クリック
$SendMouseClick::mouse_event(0x0002, 0, 0, 0, 0);

Start-sleep -s 1

$SendMouseClick::mouse_event(0x0004, 0, 0, 0, 0);

Start-Sleep -s 2
$page = Get-Random -Maximum $BookPages -Minimum 1
[System.Windows.Forms.SendKeys]::SendWait($page)
Start-Sleep -s 2

# 変数設定
#---C---
$x = 1064
$y = 632

# マウスカーソル移動
[System.Windows.Forms.Cursor]::Position = New-Object System.Drawing.Point($x, $y)

Start-sleep -s 1

# 左クリック
$SendMouseClick::mouse_event(0x0002, 0, 0, 0, 0);

Start-sleep -s 1

$SendMouseClick::mouse_event(0x0004, 0, 0, 0, 0);

Start-Sleep -s 10

$ps_speak.Speak($arr[($page-1)])


Start-Sleep -s $ReadSecond


}


