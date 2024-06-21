using System;
using System.Collections.Generic;
using System.Linq;
using System.Numerics;
using System.Runtime.InteropServices;
using System.Text;
using Dalamud.Game.ClientState.Objects.Enums;
using Dalamud.Game.Text.SeStringHandling.Payloads;
using Dalamud.Logging;
using FFXIVClientStructs.FFXIV.Client.System.Framework;
using FFXIVClientStructs.FFXIV.Client.UI;
using FFXIVClientStructs.FFXIV.Component.GUI;
using Lumina.Excel;
using Lumina.Excel.GeneratedSheets;

namespace RankSSpawnHelper;

internal static class Utils
{
    internal const int MinutesOfHour = 60;
    internal const int HoursOfDay = 24;
    internal const int DaysOfMonth = 32;
    internal const int MonthsOfYear = 12;
    private const double TimeRate = 175.0;

    private static unsafe UIModule* _uiModule;
    private static ProcessChatBoxDelegate _processChatBox;
    private static PlaySound _playSound;

    private static ExcelSheet<TerritoryType> _terr;
    private static List<string> _rankSMonsterName;

    public static DateTime TargetEorzeaTime = new();

    public static unsafe void Initialize()
    {
        _terr = DalamudApi.DataManager.GetExcelSheet<TerritoryType>();
        var bNpcNames = DalamudApi.DataManager.GetExcelSheet<BNpcName>();

        var names = new List<string>();
        // 2.0
        {
            for (uint i = 2953; i < 2970; i++)
                names.Add(bNpcNames.GetRow(i).Singular.RawString);
        }

        // 3.0
        {
            for (uint i = 4374; i < 4381; i++)
                if (i != 4379)
                    names.Add(bNpcNames.GetRow(i).Singular.RawString);
        }

        // 4.0
        {
            for (uint i = 5984; i < 5990; i++)
                names.Add(bNpcNames.GetRow(i).Singular.RawString);
        }

        // 5.0
        {
            names.Add(bNpcNames.GetRow(8653).Singular.RawString); // 阿格拉俄珀
            for (uint i = 8895; i < 8916; i += 5)
                names.Add(bNpcNames.GetRow(i).Singular.RawString);
        }

        // 6.0
        {
            for (uint i = 10615; i < 10623; i++)
                if (i != 10616) // 克尔的侍从
                    names.Add(bNpcNames.GetRow(i).Singular.RawString);
        }

        _rankSMonsterName = names;

        var easierProcessChatBoxPtr = DalamudApi.SigScanner.ScanText("48 89 5C 24 ?? 57 48 83 EC 20 48 8B FA 48 8B D9 45 84 C9");
        _processChatBox = Marshal.GetDelegateForFunctionPointer<ProcessChatBoxDelegate>(easierProcessChatBoxPtr);

        var playSoundAddress = DalamudApi.SigScanner.ScanText("E8 ?? ?? ?? ?? 4D 39 BE");
        _playSound = Marshal.GetDelegateForFunctionPointer<PlaySound>(playSoundAddress);

        _uiModule = Framework.Instance()->GetUiModule();
    }

    public static unsafe void ExecuteCommand(string cmd)
    {
        try
        {
            var bytes = Encoding.UTF8.GetBytes(cmd);

            var mem1 = Marshal.AllocHGlobal(400);
            var mem2 = Marshal.AllocHGlobal(bytes.Length + 30);

            Marshal.Copy(bytes, 0, mem2, bytes.Length);
            Marshal.WriteByte(mem2 + bytes.Length, 0);
            Marshal.WriteInt64(mem1, mem2.ToInt64());
            Marshal.WriteInt64(mem1 + 8, 64);
            Marshal.WriteInt64(mem1 + 8 + 8, bytes.Length + 1);
            Marshal.WriteInt64(mem1 + 8 + 8 + 8, 0);

            _processChatBox!(_uiModule, mem1, IntPtr.Zero, 0);

            Marshal.FreeHGlobal(mem1);
            Marshal.FreeHGlobal(mem2);
        }
        catch (Exception err)
        {
            DalamudApi.ChatGui.PrintError(err.Message);
        }
    }
    
    public static void PlayChatSoundSound(uint effectId)
    {
        if (effectId is < 1 or 19 or 21)
        {
            DalamudApi.ChatGui.PrintError("Valid chat sfx values are 1 through 16.");
            return;
        }

        _playSound(effectId + 0x24u, 0, 0, 0);
    }

    public static unsafe DateTime LocalTimeToEorzeaTime(double minutes = 0, double seconds = 0)
    {
        try
        {
            const double eorzeaMultiplier = 3600D / 175D;

            var epochTicks = DateTimeOffset.FromUnixTimeSeconds(Framework.Instance()->ServerTime).AddMinutes(minutes).AddSeconds(seconds).ToUniversalTime().Ticks - new DateTime(1970, 1, 1).Ticks;

            var eorzeaTicks = (long)Math.Round(epochTicks * eorzeaMultiplier);

            return new(eorzeaTicks);

        }
        catch (Exception e)
        {
            PluginLog.Debug(e, "Exception happened when converting local time to eorzea time");
            return new(0);
        }
    }

    public static DateTime EorzaTimeToLocalTime(DateTime et)
    {
        var months = MonthsOfYear * (et.Year - 1) + (et.Month - 1);
        var days = DaysOfMonth * months + (et.Day - 1);
        var hours = HoursOfDay * days + et.Hour;
        var minutes = MinutesOfHour * hours + et.Minute;
        var seconds = (long)Math.Round(minutes * TimeRate / MinutesOfHour);

        var utc = DateTimeOffset.FromUnixTimeSeconds(seconds);
        return utc.DateTime;
    }

    private static float ToMapCoordinate(float val, float scale)
    {
        var c = scale / 100.0f;

        val *= c;

        return 41.0f / c * ((val + 1024.0f) / 2048.0f) + 1;
    }

    public static bool IsSRankMonster(string name)
    {
        return _rankSMonsterName.Contains(name);
    }

    public static MapLinkPayload CreateMapLinkPayload(Vector3 pos, ushort territoryId)
    {
        var territory = _terr.GetRow(territoryId);

        var x = ToMapCoordinate(pos.X, territory.Map.Value.SizeFactor);
        var y = ToMapCoordinate(pos.Z, territory.Map.Value.SizeFactor);
        var payload = new MapLinkPayload(territory.RowId, territory.Map.Row, x, y);

        return payload;
    }

    public static unsafe void PrintSetTimeMessage(bool preview = false, bool yell = false, string etSet = null)
    {
        if (etSet != null)
        {
            string message = "在<pos>发现了<t>，预定于ET" + etSet + "开打，请勿脸开抢开跟开";
            ExecuteCommand($"/sh {message}");
            if (yell) ExecuteCommand($"/yell {message}");
            return;
        }

        if (preview && DalamudApi.ClientState.LocalPlayer == null)
            return;

        var obj = DalamudApi.ObjectTable.Where(i => i.IsValid() && i.ObjectKind == ObjectKind.BattleNpc && _rankSMonsterName.Contains(i.Name.TextValue)).Select(i => i).ToList();
        if (obj.Count == 0 && !preview)
        {
            DalamudApi.ChatGui.PrintError("[ET喊话] 地图里没有S怪");
            return;
        }

        if (!preview && !IsSRankMonster(obj[0].Name.TextValue))
        {
            DalamudApi.ChatGui.PrintError("[ET喊话] 地图里没有S怪");
            return;
        }

        var msg = DalamudApi.Configuration._mainSetTimeMessage;
        if (msg.Contains("{tpos}")) msg = msg.Replace("{tpos}", "<flag>");

        if (msg.Contains("{tname}")) msg = msg.Replace("{tname}", preview ? DalamudApi.ClientState.LocalPlayer.Name.TextValue : obj[0].Name.TextValue);

        var currentEt = LocalTimeToEorzeaTime();
        if (msg.Contains("{etmsg}"))
        {
            var unsetMessage = DalamudApi.Configuration._etMessageUnset;
            var setMessage = DalamudApi.Configuration._etMessageSet;

            if (preview)
            {
                var backup = msg.Replace("{etmsg}", "");
                msg = "已定ET消息:" + backup + setMessage.Replace("{et}", $"{TargetEorzeaTime.Hour:D2}:{TargetEorzeaTime.Minute:D2}") + "\n未定ET消息:";
                msg += backup + unsetMessage;
            }
            else
            {
                msg = msg.Replace("{etmsg}", currentEt > TargetEorzeaTime ? unsetMessage : setMessage.Replace("{et}", $"{TargetEorzeaTime.Hour:D2}:{TargetEorzeaTime.Minute:D2}"));
            }
        }

        var mapLinkPayload = CreateMapLinkPayload(preview ? DalamudApi.ClientState.LocalPlayer.Position : obj[0].Position, DalamudApi.ClientState.TerritoryType);
        DalamudApi.GameGui.OpenMapWithMapLink(mapLinkPayload);

        var areaMapPtr = DalamudApi.GameGui.GetAddonByName("AreaMap", 1);
        if (areaMapPtr != IntPtr.Zero) ((AtkUnitBase*)areaMapPtr)->Hide2();

        if (preview)
        {
            ExecuteCommand($"/e {msg}");
            return;
        }

        ExecuteCommand($"/sh {msg}");
        if (yell)
            ExecuteCommand($"/yell {msg}");
    }

    private delegate bool PlaySound(uint effectId, long a2, long a3, byte a4);

    private unsafe delegate void ProcessChatBoxDelegate(UIModule* uiModule, IntPtr message, IntPtr unused, byte a4);

}