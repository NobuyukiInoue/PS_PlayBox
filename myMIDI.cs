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
        midiOutShortMsg(h, timbre);     // âπêFÇíËã`
    }

    public void OutOnly(UInt32 outData)
    {
        midiOutShortMsg(h, outData);    // åÆî’ÇâüÇ∑
    }

    public void Out(UInt32 outData, Int32 len)
    {
        midiOutShortMsg(h, outData);         // åÆî’ÇâüÇ∑
        System.Threading.Thread.Sleep(len);  // àÍíËéûä‘âπÇñ¬ÇÁÇµë±ÇØÇÈ
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
