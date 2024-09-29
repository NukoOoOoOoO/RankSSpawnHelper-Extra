using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Dalamud.Game;
using Dalamud.Game.ClientState.Conditions;
using Dalamud.Game.ClientState.Objects.Enums;
using Dalamud.Logging;
using Dalamud.Plugin.Services;
using Dalamud.Utility.Signatures;
using FFXIVClientStructs.FFXIV.Client.Game;
using Lumina.Excel.GeneratedSheets;

namespace RankSSpawnHelper.Features;

internal class SummonMinion : IDisposable
{
    [Signature("48 8D 0D ?? ?? ?? ?? 0F B6 04 08 84 D0 75 10 B8 ?? ?? ?? ?? 48 8B 5C 24",
        ScanType = ScanType.StaticAddress)]
    private readonly IntPtr? _minionBitmask = null;

    private DateTime _lastUpDateTime;

    private readonly Dictionary<ushort, uint> _minionMap = new()
    {
        { 960, 423 },
        { 816, 303 },
        { 956, 434 },
        { 614, 215 },
        { 397, 148 }
    };

    private readonly List<Tuple<uint, string>> _unlockedMinions = new();

    public SummonMinion()
    {
        DalamudApi.GameInteropProvider.InitializeFromAttributes(this);
        DalamudApi.Framework.Update += FrameworkOnUpdate;

        unsafe bool IsMinionUnlocked(uint minionId)
        {
            if (_minionBitmask == null || _minionBitmask.Value == IntPtr.Zero) return false;

            return ((1 << ((int)minionId & 7)) & ((byte*)_minionBitmask.Value)[minionId >> 3]) > 0;
        }

        Task.Run(async () =>
        {
            while (DalamudApi.ClientState.LocalPlayer == null) await Task.Delay(1000);

            var unlockedCompanions = DalamudApi.DataManager.GetExcelSheet<Companion>().Where(i =>
                IsMinionUnlocked(i.RowId) && i.RowId is 434 or 423 or 215 or 303 or 148);

            foreach (var companion in unlockedCompanions)
            {
                DalamudApi.PluginLog.Debug($"{companion.Singular.RawString}");
                _unlockedMinions.Add(new(companion.RowId, companion.Singular.RawString));
            }
        });
    }

    public void Dispose()
    {
        DalamudApi.Framework.Update -= FrameworkOnUpdate;
    }

    private void FrameworkOnUpdate(IFramework framework)
    {
        if (DateTime.Now - _lastUpDateTime <= TimeSpan.FromSeconds(2))
        {
            return;
        }

        if (!DalamudApi.Configuration._summonMinion)
            goto end;

        if (!_minionMap.ContainsKey(DalamudApi.ClientState.TerritoryType)) goto end;

        if (!_minionMap.TryGetValue(DalamudApi.ClientState.TerritoryType, out var currnetMinionId)) goto end;

        var minion = _unlockedMinions.Find(i => i.Item1 == currnetMinionId);

        if (minion == null) goto end;

        if (DalamudApi.Condition[ConditionFlag.Mounted] || DalamudApi.Condition[ConditionFlag.Mounted2] ||
            DalamudApi.Condition[ConditionFlag.Unknown57] ||
            DalamudApi.Condition[ConditionFlag.Mounting] || DalamudApi.Condition[ConditionFlag.Mounting71])
            goto end;

        if (!CanUseAction(minion.Item1))
            goto end;

        if (DalamudApi.ObjectTable[1] == null && CanUseAction(minion.Item1))
        {
            UseAction(minion.Item1);
            goto end;
        }

        var obj = DalamudApi.ObjectTable[1];
        if (obj == null)
        {
            UseAction(minion.Item1);
            goto end;
        }

        if (obj.ObjectKind != ObjectKind.Companion)
        {
            UseAction(minion.Item1);
            goto end;
        }

        if (!CanUseAction(minion.Item1))
            goto end;
        if (obj.Name.TextValue == minion.Item2)
            goto end;

        UseAction(minion.Item1);

        end:
        _lastUpDateTime = DateTime.Now;
    }

    private static unsafe bool CanUseAction(uint id)
    {
        return ActionManager.Instance()->GetActionStatus(ActionType.Companion, id) == 0 &&
               !ActionManager.Instance()->IsRecastTimerActive(ActionType.Action, id);
    }
    
    private static unsafe void UseAction(uint id)
    {
        ActionManager.Instance()->UseAction(ActionType.Companion, id);
    }
}