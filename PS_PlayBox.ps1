param( $musicFile, $timbre )

##--------------------------------------------------------##
## 音階定義ファイルの読み込み
##--------------------------------------------------------##
function loadDefFile([string]$defFileName)
{
    $f = (Get-Content $defFileName) -as [string[]]
    $lines = @()

    # 音階名とノートナンバーに分け、配列に格納する
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

        # "="で区切り、音階とノートナンバーに分けて格納する
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
## 音階文字列を検索し、ノートナンバーをセットする
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
    $timbre = 1     # ピアノ
}

$defs = @()
$defs = loadDefFile "note-number.dat" 

$playData = @()
$playData = loadPlayFile $musicFile

# 音階をMIDIノートナンバーに変換する
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
