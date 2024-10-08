using System;
using System.IO;
using System.Runtime.InteropServices;
using Dalamud.Game.ClientState.Conditions;
using Dalamud.Logging;
using Dalamud.Plugin.Services;

namespace RankSSpawnHelper.Features;

public class LeaveDuty : IDisposable
{
    // taken from https://github.com/marimelon/LeaveDutyCmdPlugin/blob/master/LeaveDutyCmdPlugin/LeaveDutyCmdPlugin.cs
    private readonly LeaveDutyDelegate _leaveDungeon;
    private readonly CanLeaveDutyDelegate _canLeaveDutyDelegate;

    public LeaveDuty()
    {
        try
        {
            _leaveDungeon = Marshal.GetDelegateForFunctionPointer<LeaveDutyDelegate>(DalamudApi.SigScanner.ScanText("40 53 48 83 EC ?? 48 8B 05 ?? ?? ?? ?? 0F B6 D9 48 85 C0"));
            _canLeaveDutyDelegate = Marshal.GetDelegateForFunctionPointer<CanLeaveDutyDelegate>(DalamudApi.SigScanner.ScanText("48 83 EC ?? 48 8B 05 ?? ?? ?? ?? 48 85 C0 74 ?? 83 B8 ?? ?? ?? ?? ?? 75 ?? E8"));
            DalamudApi.Framework.Update += OnFrameworkUpdate;
        }
        catch (Exception e)
        {
            throw new InvalidDataException($"Failed to load leave duty. r:{e}");
        }
    }

    public void Dispose()
    {
        DalamudApi.Framework.Update -= OnFrameworkUpdate;
    }

    private void OnFrameworkUpdate(IFramework _)
    {
        if (!DalamudApi.Configuration._autoLeaveDuty)
            return;

        if (DalamudApi.PartyList.Length > 0)
            return;

        var local = DalamudApi.ClientState.LocalPlayer;
        if (local == null)
            return;

        if (local.ClassJob.Id != 36)
            return;

        if (DalamudApi.ClientState.TerritoryType != 1045)
            return;

        if (!_canLeaveDutyDelegate())
            return;

        _leaveDungeon((char)0);
    }

    public void Execute()
    {
        if (!DalamudApi.Configuration._autoLeaveDuty)
            return;

        if (DalamudApi.PartyList.Length > 0)
            return;

        if (DalamudApi.ClientState.LocalPlayer.ClassJob.Id != 36)
        {
            DalamudApi.ChatGui.PrintError("[自动退本] 职业只能是青魔");
            return;
        }

        if (DalamudApi.ClientState.TerritoryType != 1045)
        {
            DalamudApi.ChatGui.PrintError("[自动退本] 该功能只能在假火神(伊芙利特讨伐战)用");
            return;
        }

        _leaveDungeon((char)0);
    }

    private delegate void LeaveDutyDelegate(char isTimeout);
    private delegate bool CanLeaveDutyDelegate();
}