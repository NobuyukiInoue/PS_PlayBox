PowerShell�œ������d�q�I���S�[���𐧍삵�Ă݂܂����B  
MIDI�֘A��Win32API��C#�ŗ��p����N���X���쐬���APowerShell����inline�ŌĂяo�����Ƃɂ��A�y���f�[�^�����t�����Ă݂܂��B  
  
MIDI�֘A��Win32API���Ăяo��C#�̃N���X�͂���Ȋ����ł��B  
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
        midiOutShortMsg(h, timbre);     // ���F���`
    }

    public void OutOnly(UInt32 outData)
    {
        midiOutShortMsg(h, outData);    // ���Ղ�����
    }

    public void Out(UInt32 outData, Int32 len)
    {
        midiOutShortMsg(h, outData);         // ���Ղ�����
        System.Threading.Thread.Sleep(len);  // ��莞�ԉ���炵������
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

���ɁAPowerShell���̃��C���v���O�����ł��B

###PS_PlayBox.ps1
```
param( $musicFile, $timbre )

##--------------------------------------------------------##
## ���K��`�t�@�C���̓ǂݍ���
##--------------------------------------------------------##
function loadDefFile([string]$defFileName)
{
    $f = (Get-Content $defFileName) -as [string[]]
    $lines = @()

    # ���K���Ǝ��g���ɕ����A�z��Ɋi�[����
    foreach ($currentLine in $f) {

        # �R�����g�J�n�ʒu�̌��o
        $commentStartPostion = $currentLine.IndexOf("//")

        if ($commentStartPostion -eq 0) {
            continue
        }
        elseif ($commentStartPostion -gt 0) {
            $currentLine = $currentLine.Substring(0, $commentStartPostion)
        }

        # �X�y�[�X�̍폜
        $currentLine = $currentLine.Replace(" ","")

        # TAB�̍폜
        $currentLine = $currentLine.Replace("`t", "")

        # "="�ŋ�؂�A���K�Ǝ��g���ɕ����Ċi�[����
        if ($currentLine -ne "") {
            $scale, $note = $currentLine -split "="
            $lines += New-Object PSObject -Property @{scale=$scale; note=$note}
        }
    }
    
    return($lines)
}

##--------------------------------------------------------##
## �����t�@�C���̓ǂݍ���
##--------------------------------------------------------##
function loadPlayFile([string]$musicFile)
{
    $f = (Get-Content $musicFile) -as [string[]]
    $lines = @()

    foreach ($currentLine in $f) {

        # �R�����g�J�n�ʒu�̌��o
        $commentStartPostion = $currentLine.IndexOf("//")

        if ($commentStartPostion -eq 0) {
            continue
        }
        elseif ($commentStartPostion -gt 0) {
            $currentLine = $currentLine.Substring(0, $commentStartPostion)
        }

        # �X�y�[�X�̍폜
        $currentLine = $currentLine.Replace(" ","")

        # TAB�̍폜
        $currentLine = $currentLine.Replace("`t", "")

        # "="�ŋ�؂�A���K�ƒ�����z��Ɋi�[����
        if ($currentLine -ne "") {
            $scale, $tlen = $currentLine -split "="
            $lines += New-Object PSObject -Property @{scale=$scale; note=""; tlen=$tlen}
        }
    }

    return($lines)
}

##--------------------------------------------------------##
## ���K��������������A���g�����Z�b�g����
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
## 16�i��������𐔒l�ɕϊ�����
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
    #$timbre = 42     # 0x2axx   �r�I���H�`�F���H
    $timbre = 1     # �s�A�m
}

$defs = @()
$defs = loadDefFile "note-number.dat" 

$playData = @()
$playData = loadPlayFile $musicFile

# ���K��MIDI�m�[�g�ԍ��ɕϊ�����
$playData = replaceScalt_to_Freq $defs $playData


##-----------------------------------------##
## CSharp���C�u�����̓ǂݍ���(Win32API�Q��)
##-----------------------------------------##
add-type -path .\myMIDI.cs -passThru
$pm = New-Object myMIDI

$initData = [UINT32]$timbre*256 + [Convert]::ToInt32("c0", 16)
$pm.INIT($initData)    # MIDI������

Write-Host
Write-Host "Load Done. Play Start!!"

for ($i = 0; $i -lt $playData.Length; $i++) {

    if ($playData[$i].note -ne "") {
        Write-Host "[$i] = "$playData[$i].scale"("$playData[$i].note"), "$playData[$i].tlen"[ms]"

        $cnote = $playData[$i].note -split ","

        foreach ($data in $cnote) {

            # MIDI�ɏo�͂���
            $note_on = "7f" + $data + "90"
            $play_on = [Convert]::ToInt32($note_on, 16)

            # ���Ղ�����
            $pm.OutOnly($play_on)
        }

        # ��莞�Ԗ炵������
        $pm.Sleep($playData[$i].tlen)

        foreach ($data in $cnote) {
            $note_off = "7f" + $data + "80"
            $play_off = [Convert]::ToInt32($note_off, 16)

            # ���Ղ𗣂�
            $pm.OutOnly($play_off)
        }
    } 
    else {
        Write-Host "[$i] = rest ("$playData[$i].note"), "$playData[$i].tlen"[ms]"

        # �x��
        $pm.Sleep($playData[$i].tlen)
    }
}

$pm.Close()
```


�y���f�[�^����͂���ۂɁA�h���~...��������do,re,mi,...�Ƃ����悤��  
���K�œ��͂ł���悤�ɁA���K�f�[�^�ƃR�[�h�̕ϊ��\���쐬���Ă����܂��B
###note-number.dat
```
// �S�p�J�^�J�i�\�L
�h4=3c
�h#4=3d
��4=3e
��#4=3f
�~4=40
�t4=41
�t#4=42
�\4=43
�\#4=44
��4=45
��#4=46
...
(�����̂œr���ȗ��j
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

���ɁA���t���������Ȃ̊y���f�[�^�����L�̏����ŕ��ׂĂ����܂��B

```
���K����уI�N�^�[�u�ԍ�=���t����(ms)
```

�w�����˂��񂱂�x���Ƃ���Ȋ����ɂȂ�܂��B  
���Ȃ݂ɁA���������̉����ł���΁A�a�����炷���Ƃ��ł��܂��B
###PB_kitune.txt
```
�h4,�h5 = 250
��4 = 250
�~4 = 250
�t4 = 250
�\4,�\5 = 500
�\4,�\5 = 500

��4 = 250
�t4 = 250
�h5 = 250
��4 = 250
�\4,�\5 = 1000

��4 = 250
�t4 = 250
�h5 = 250
��4 = 250
�\4,�\5 = 1000

�\4 = 250
�t4 = 250
�t4 = 250
�t4 = 250

�t4 = 250
�~4 = 250
�~4 = 250
�~4 = 250
�~4 = 250
��4 = 250
�~4 = 250
��4 = 250
�h4 = 250
�~4 = 250
�\4,�\5 = 500

�\4 = 250
�t4 = 250
�t4 = 250
�t4 = 250

�t4 = 250
�~4 = 250
�~4 = 250
�~4 = 250

�~4 = 250
��4 = 250
��4 = 250
�~4 = 250

�h4,�h5 = 1000
```

�ł͑����A���t���Ă݂܂��傤�B  
PowerShell���N�����A���L�̏�����PS_PlayBox.ps1�����s���܂��B

```
PS D:\PS_PlayBox> .\PS_PlayBox.ps1
Usage : PlayMIDI.ps1 musicDataFile <timbre>
```

tibre��MIDI�̉��F�i�y��j�̎w��ł��B0�`128�̐��l�Ŏw�肵�Ă��������B  
�ȉ��́A�����قǂ́w�����˂��񂱂�x(PS_kitune.txt)���n�[���j�J(23)�ŉ��t�����ł��B

```
PS D:\PS_PlayBox> .\PS_PlayBox.ps1 .\PB_kitune.txt 23

IsPublic IsSerial Name                                     BaseType
-------- -------- ----                                     --------
True     False    myMIDI                                   System.Object

Load Done. Play Start!!
[0] =  �h4,�h5 ( 3c,48 ),  250 [ms]
[1] =  ��4 ( 3e ),  250 [ms]
[2] =  �~4 ( 40 ),  250 [ms]
[3] =  �t4 ( 41 ),  250 [ms]
[4] =  �\4,�\5 ( 43,4f ),  500 [ms]
[5] =  �\4,�\5 ( 43,4f ),  500 [ms]
[6] =  ��4 ( 45 ),  250 [ms]
[7] =  �t4 ( 41 ),  250 [ms]
[8] =  �h5 ( 48 ),  250 [ms]
[9] =  ��4 ( 45 ),  250 [ms]
[10] =  �\4,�\5 ( 43,4f ),  1000 [ms]
[11] =  ��4 ( 45 ),  250 [ms]
[12] =  �t4 ( 41 ),  250 [ms]
[13] =  �h5 ( 48 ),  250 [ms]
[14] =  ��4 ( 45 ),  250 [ms]
[15] =  �\4,�\5 ( 43,4f ),  1000 [ms]
[16] =  �\4 ( 43 ),  250 [ms]
[17] =  �t4 ( 41 ),  250 [ms]
[18] =  �t4 ( 41 ),  250 [ms]
[19] =  �t4 ( 41 ),  250 [ms]
[20] =  �t4 ( 41 ),  250 [ms]
[21] =  �~4 ( 40 ),  250 [ms]
[22] =  �~4 ( 40 ),  250 [ms]
[23] =  �~4 ( 40 ),  250 [ms]
[24] =  �~4 ( 40 ),  250 [ms]
[25] =  ��4 ( 3e ),  250 [ms]
[26] =  �~4 ( 40 ),  250 [ms]
[27] =  ��4 ( 3e ),  250 [ms]
[28] =  �h4 ( 3c ),  250 [ms]
[29] =  �~4 ( 40 ),  250 [ms]
[30] =  �\4,�\5 ( 43,4f ),  500 [ms]
[31] =  �\4 ( 43 ),  250 [ms]
[32] =  �t4 ( 41 ),  250 [ms]
[33] =  �t4 ( 41 ),  250 [ms]
[34] =  �t4 ( 41 ),  250 [ms]
[35] =  �t4 ( 41 ),  250 [ms]
[36] =  �~4 ( 40 ),  250 [ms]
[37] =  �~4 ( 40 ),  250 [ms]
[38] =  �~4 ( 40 ),  250 [ms]
[39] =  �~4 ( 40 ),  250 [ms]
[40] =  ��4 ( 3e ),  250 [ms]
[41] =  ��4 ( 3e ),  250 [ms]
[42] =  �~4 ( 40 ),  250 [ms]
[43] =  �h4,�h5 ( 3c,48 ),  1000 [ms]


PS D:\PS_PlayBox>
```
