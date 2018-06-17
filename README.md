PowerShellで動かす電子オルゴールを制作してみました。  
MIDI関連のWin32APIをC#で利用するクラスを作成し、PowerShellからinlineで呼び出すことにより、楽譜データを演奏させてみます。  
  
MIDI関連のWin32APIを呼び出すC#のクラスはこんな感じです。  
###myMIDI.cs
```
using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;

using HANDLE = System.IntPtr;
using MMRESULT = System.UInt32;
using UINT = System.UInt32;
using DWORD_PTR = System.IntPtr;
using DWORD = System.UInt64;
using LPHMIDIOUT = System.IntPtr;
using HMIDIOUT = System.IntPtr;

public class myMIDI
{
    //const UInt32 MIDI_MAPPER = -1;
    const UInt32 MIDI_MAPPER = UInt32.MaxValue;
    HMIDIOUT h;

    [DllImport("Winmm.dll", EntryPoint = "midiOutOpen")]
    private static extern MMRESULT midiOutOpen(
       ref LPHMIDIOUT lphmo,
       UINT uDeviceID,
       DWORD_PTR dwCallback,
       DWORD_PTR dwCallbackInstance,
       DWORD dwFlags
    );

    [DllImport("Winmm.dll", EntryPoint = "midiOutShortMsg")]
    private static extern MMRESULT midiOutShortMsg(
        HMIDIOUT hmo,
        DWORD dwMsg
    );

    [DllImport("Winmm.dll", EntryPoint = "midiOutReset")]
    private static extern MMRESULT midiOutReset(
          HMIDIOUT hmo
    );

    [DllImport("Winmm.dll", EntryPoint = "midiOutClose")]
    private static extern MMRESULT midiOutClose(
          HMIDIOUT hmo
    );

    public void Init(UInt32 timbre)
    {
        midiOutOpen(ref h, MIDI_MAPPER, (System.IntPtr)0, (System.IntPtr)0, 0);
        midiOutShortMsg(h, timbre);     // 音色を定義
    }

    public void OutOnly(UInt32 outData)
    {
        midiOutShortMsg(h, outData);    // 鍵盤を押す
    }

    public void Out(UInt32 outData, Int32 len)
    {
        midiOutShortMsg(h, outData);         // 鍵盤を押す
        System.Threading.Thread.Sleep(len);  // 一定時間音を鳴らし続ける
    }

    public void Close()
    {
        midiOutReset(h);
        midiOutClose(h);
    }

    public HMIDIOUT getHvalue()
    {
        return (h);
    }

    public void Sleep(Int32 len)
    {
        System.Threading.Thread.Sleep(len);
    }
}
```

次に、PowerShell側のメインプログラムです。

###PS_PlayBox.ps1
```
param( $musicFile, $timbre )

##--------------------------------------------------------##
## 音階定義ファイルの読み込み
##--------------------------------------------------------##
function loadDefFile([string]$defFileName)
{
    $f = (Get-Content $defFileName) -as [string[]]
    $lines = @()

    # 音階名と周波数に分け、配列に格納する
    foreach ($currentLine in $f) {

        # コメント開始位置の検出
        $commentStartPostion = $currentLine.IndexOf("//")

        if ($commentStartPostion -eq 0) {
            continue
        }
        elseif ($commentStartPostion -gt 0) {
            $currentLine = $currentLine.Substring(0, $commentStartPostion)
        }

        # スペースの削除
        $currentLine = $currentLine.Replace(" ","")

        # TABの削除
        $currentLine = $currentLine.Replace("`t", "")

        # "="で区切り、音階と周波数に分けて格納する
        if ($currentLine -ne "") {
            $scale, $note = $currentLine -split "="
            $lines += New-Object PSObject -Property @{scale=$scale; note=$note}
        }
    }
    
    return($lines)
}

##--------------------------------------------------------##
## 音譜ファイルの読み込み
##--------------------------------------------------------##
function loadPlayFile([string]$musicFile)
{
    $f = (Get-Content $musicFile) -as [string[]]
    $lines = @()

    foreach ($currentLine in $f) {

        # コメント開始位置の検出
        $commentStartPostion = $currentLine.IndexOf("//")

        if ($commentStartPostion -eq 0) {
            continue
        }
        elseif ($commentStartPostion -gt 0) {
            $currentLine = $currentLine.Substring(0, $commentStartPostion)
        }

        # スペースの削除
        $currentLine = $currentLine.Replace(" ","")

        # TABの削除
        $currentLine = $currentLine.Replace("`t", "")

        # "="で区切り、音階と長さを配列に格納する
        if ($currentLine -ne "") {
            $scale, $tlen = $currentLine -split "="
            $lines += New-Object PSObject -Property @{scale=$scale; note=""; tlen=$tlen}
        }
    }

    return($lines)
}

##--------------------------------------------------------##
## 音階文字列を検索し、周波数をセットする
##--------------------------------------------------------##
function replaceScalt_to_Freq([array]$defs, [array]$playData)
{
    for ($i = 0; $i -lt $playData.Length; $i++) {

        $scale = $playData[$i].scale -split ","

        foreach ($temp in $scale) {

            for ($j = 0; $j -lt $defs.Length; $j++) {

                if ($temp -eq $defs[$j].scale) {

                    if ($playData[$i].note -eq "") {
                       $playData[$i].note = $defs[$j].note
                    }
                    else {
                       $playData[$i].note += "," + $defs[$j].note
                    }

                    break
                }
            }
        }
    }
    
    return($playData)
}

<#
##--------------------------------------------------------##
## 16進数文字列を数値に変換する
##--------------------------------------------------------##
function toHex([string]tempStr)
{
	$ret = [Convert]::ToInt32($tempStr, 16)
	return($ret)
}
#>

##--------------------------------------------------------##
## Main
##--------------------------------------------------------##

if (-Not($musicFile)) {
    Write-Host "Usage : PlayMIDI.ps1 musicDataFile <timbre>"
    exit
}

if (-Not($timbre)) {
    #$timbre = 42     # 0x2axx   ビオラ？チェロ？
    $timbre = 1     # ピアノ
}

$defs = @()
$defs = loadDefFile "note-number.dat" 

$playData = @()
$playData = loadPlayFile $musicFile

# 音階をMIDIノート番号に変換する
$playData = replaceScalt_to_Freq $defs $playData


##-----------------------------------------##
## CSharpライブラリの読み込み(Win32API参照)
##-----------------------------------------##
add-type -path .\myMIDI.cs -passThru
$pm = New-Object myMIDI

$initData = [UINT32]$timbre*256 + [Convert]::ToInt32("c0", 16)
$pm.INIT($initData)    # MIDI初期化

Write-Host
Write-Host "Load Done. Play Start!!"

for ($i = 0; $i -lt $playData.Length; $i++) {

    if ($playData[$i].note -ne "") {
        Write-Host "[$i] = "$playData[$i].scale"("$playData[$i].note"), "$playData[$i].tlen"[ms]"

        $cnote = $playData[$i].note -split ","

        foreach ($data in $cnote) {

            # MIDIに出力する
            $note_on = "7f" + $data + "90"
            $play_on = [Convert]::ToInt32($note_on, 16)

            # 鍵盤を押す
            $pm.OutOnly($play_on)
        }

        # 一定時間鳴らし続ける
        $pm.Sleep($playData[$i].tlen)

        foreach ($data in $cnote) {
            $note_off = "7f" + $data + "80"
            $play_off = [Convert]::ToInt32($note_off, 16)

            # 鍵盤を離す
            $pm.OutOnly($play_off)
        }
    } 
    else {
        Write-Host "[$i] = rest ("$playData[$i].note"), "$playData[$i].tlen"[ms]"

        # 休符
        $pm.Sleep($playData[$i].tlen)
    }
}

$pm.Close()
```


楽譜データを入力する際に、ドレミ...もしくはdo,re,mi,...というように  
音階で入力できるように、音階データとコードの変換表を作成しておきます。
###note-number.dat
```
// 全角カタカナ表記
ド4=3c
ド#4=3d
レ4=3e
レ#4=3f
ミ4=40
フ4=41
フ#4=42
ソ4=43
ソ#4=44
ラ4=45
ラ#4=46
...
(長いので途中省略）
...
...
re#8=6f
mi8=70
fa8=71
fa#8=72
so8=73
so#8=74
ra8=75
ra#8=76
si8=77

```

次に、演奏させたい曲の楽譜データを下記の書式で並べていきます。

```
音階およびオクターブ番号=演奏時間(ms)
```

『こぎつねこんこん』だとこんな感じになります。  
ちなみに、同じ長さの音符であれば、和音も鳴らすことができます。
###PB_kitune.txt
```
ド4,ド5 = 250
レ4 = 250
ミ4 = 250
フ4 = 250
ソ4,ソ5 = 500
ソ4,ソ5 = 500

ラ4 = 250
フ4 = 250
ド5 = 250
ラ4 = 250
ソ4,ソ5 = 1000

ラ4 = 250
フ4 = 250
ド5 = 250
ラ4 = 250
ソ4,ソ5 = 1000

ソ4 = 250
フ4 = 250
フ4 = 250
フ4 = 250

フ4 = 250
ミ4 = 250
ミ4 = 250
ミ4 = 250
ミ4 = 250
レ4 = 250
ミ4 = 250
レ4 = 250
ド4 = 250
ミ4 = 250
ソ4,ソ5 = 500

ソ4 = 250
フ4 = 250
フ4 = 250
フ4 = 250

フ4 = 250
ミ4 = 250
ミ4 = 250
ミ4 = 250

ミ4 = 250
レ4 = 250
レ4 = 250
ミ4 = 250

ド4,ド5 = 1000
```

では早速、演奏してみましょう。  
PowerShellを起動し、下記の書式でPS_PlayBox.ps1を実行します。

```
PS D:\PS_PlayBox> .\PS_PlayBox.ps1
Usage : PlayMIDI.ps1 musicDataFile <timbre>
```

tibreはMIDIの音色（楽器）の指定です。0～128の数値で指定してください。  
以下は、さきほどの『ごぎつねこんこん』(PS_kitune.txt)をハーモニカ(23)で演奏する例です。

```
PS D:\PS_PlayBox> .\PS_PlayBox.ps1 .\PB_kitune.txt 23

IsPublic IsSerial Name                                     BaseType
-------- -------- ----                                     --------
True     False    myMIDI                                   System.Object

Load Done. Play Start!!
[0] =  ド4,ド5 ( 3c,48 ),  250 [ms]
[1] =  レ4 ( 3e ),  250 [ms]
[2] =  ミ4 ( 40 ),  250 [ms]
[3] =  フ4 ( 41 ),  250 [ms]
[4] =  ソ4,ソ5 ( 43,4f ),  500 [ms]
[5] =  ソ4,ソ5 ( 43,4f ),  500 [ms]
[6] =  ラ4 ( 45 ),  250 [ms]
[7] =  フ4 ( 41 ),  250 [ms]
[8] =  ド5 ( 48 ),  250 [ms]
[9] =  ラ4 ( 45 ),  250 [ms]
[10] =  ソ4,ソ5 ( 43,4f ),  1000 [ms]
[11] =  ラ4 ( 45 ),  250 [ms]
[12] =  フ4 ( 41 ),  250 [ms]
[13] =  ド5 ( 48 ),  250 [ms]
[14] =  ラ4 ( 45 ),  250 [ms]
[15] =  ソ4,ソ5 ( 43,4f ),  1000 [ms]
[16] =  ソ4 ( 43 ),  250 [ms]
[17] =  フ4 ( 41 ),  250 [ms]
[18] =  フ4 ( 41 ),  250 [ms]
[19] =  フ4 ( 41 ),  250 [ms]
[20] =  フ4 ( 41 ),  250 [ms]
[21] =  ミ4 ( 40 ),  250 [ms]
[22] =  ミ4 ( 40 ),  250 [ms]
[23] =  ミ4 ( 40 ),  250 [ms]
[24] =  ミ4 ( 40 ),  250 [ms]
[25] =  レ4 ( 3e ),  250 [ms]
[26] =  ミ4 ( 40 ),  250 [ms]
[27] =  レ4 ( 3e ),  250 [ms]
[28] =  ド4 ( 3c ),  250 [ms]
[29] =  ミ4 ( 40 ),  250 [ms]
[30] =  ソ4,ソ5 ( 43,4f ),  500 [ms]
[31] =  ソ4 ( 43 ),  250 [ms]
[32] =  フ4 ( 41 ),  250 [ms]
[33] =  フ4 ( 41 ),  250 [ms]
[34] =  フ4 ( 41 ),  250 [ms]
[35] =  フ4 ( 41 ),  250 [ms]
[36] =  ミ4 ( 40 ),  250 [ms]
[37] =  ミ4 ( 40 ),  250 [ms]
[38] =  ミ4 ( 40 ),  250 [ms]
[39] =  ミ4 ( 40 ),  250 [ms]
[40] =  レ4 ( 3e ),  250 [ms]
[41] =  レ4 ( 3e ),  250 [ms]
[42] =  ミ4 ( 40 ),  250 [ms]
[43] =  ド4,ド5 ( 3c,48 ),  1000 [ms]


PS D:\PS_PlayBox>
```
