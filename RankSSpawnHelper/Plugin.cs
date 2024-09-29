using System;
using Dalamud.Interface.Windowing;
using Dalamud.IoC;
using Dalamud.Plugin;
using RankSSpawnHelper.Features;

namespace RankSSpawnHelper;

public class Plugin : IDalamudPlugin
{
    private readonly WindowSystem _windowSystem;

    public Plugin(IDalamudPluginInterface pluginInterface)
    {
        pluginInterface.Create<DalamudApi>();
        Utils.Initialize();

        DalamudApi.Configuration = pluginInterface.GetPluginConfig() as Configuration ?? new Configuration();

        DalamudApi.Commands = new();
        DalamudApi.ConfigWindow = new();
        DalamudApi.AutoDiscardItem = new();
        DalamudApi.LeaveDuty = new();
        DalamudApi.Journal = new();
        DalamudApi.SummonMinion = new();

        _windowSystem = new("RankSSpawnHelper-Extra");
        _windowSystem.AddWindow(DalamudApi.ConfigWindow);

        DalamudApi.Interface.UiBuilder.OpenConfigUi += OpenConfigUi;
        DalamudApi.Interface.UiBuilder.Draw += _windowSystem.Draw;
    }

    public string Name => "S怪触发小助手-Extra";

    public void Dispose()
    {
        DalamudApi.Commands.Dispose();
        GC.SuppressFinalize(this);
        DalamudApi.AutoDiscardItem.Dispose();
        DalamudApi.LeaveDuty.Dispose();
        DalamudApi.Journal.Dispose();
        DalamudApi.SummonMinion.Dispose();

        DalamudApi.Interface.UiBuilder.OpenConfigUi -= OpenConfigUi;
        DalamudApi.Interface.UiBuilder.Draw -= _windowSystem.Draw;
    }

    private static void OpenConfigUi()
    {
        DalamudApi.ConfigWindow.IsOpen = true;
    }
}