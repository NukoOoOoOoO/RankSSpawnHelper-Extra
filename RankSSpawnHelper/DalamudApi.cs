using Dalamud.Data;
using Dalamud.Game;
using Dalamud.Game.ClientState;
using Dalamud.Game.ClientState.Conditions;
using Dalamud.Game.ClientState.Objects;
using Dalamud.Game.ClientState.Party;
using Dalamud.Game.Command;
using Dalamud.Game.Gui;
using Dalamud.IoC;
using Dalamud.Plugin;
using Dalamud.Plugin.Services;
using RankSSpawnHelper.Features;

// ReSharper disable AutoPropertyCanBeMadeGetOnly.Local
namespace RankSSpawnHelper;

internal class DalamudApi
{
    internal static Commands Commands { get; set; } = null!;
    internal static Configuration Configuration { get; set; } = null!;
    internal static ConfigWindow ConfigWindow { get; set; } = null!;
    internal static AutoDiscardItem AutoDiscardItem { get; set; } = null!;
    internal static LeaveDuty LeaveDuty { get; set; } = null!;
    internal static JournalStuff Journal { get; set; } = null!;
    internal static SummonMinion SummonMinion { get; set; } = null!;

    [PluginService] internal static IDalamudPluginInterface Interface { get; private set; } = null!;

    [PluginService] internal static IChatGui ChatGui { get; private set; } = null!;

    [PluginService] internal static IClientState ClientState { get; private set; } = null!;

    [PluginService] internal static ICommandManager CommandManager { get; private set; } = null!;

    [PluginService] internal static ISigScanner SigScanner { get; private set; } = null!;

    [PluginService] internal static IDataManager DataManager { get; private set; } = null!;

    [PluginService] internal static IGameGui GameGui { get; set; } = null!;

    [PluginService] internal static ICondition Condition { get; set; } = null!;

    [PluginService] internal static IPartyList PartyList { get; set; } = null!;

    [PluginService] internal static IObjectTable ObjectTable { get; set; } = null!;

    [PluginService] internal static IGameInteropProvider GameInteropProvider { get; set; } = null!;

    [PluginService] internal static IFramework Framework { get; set; } = null!;
    [PluginService] internal static IAddonLifecycle AddonLifecycle { get; set; } = null!;
    [PluginService] internal static IPluginLog PluginLog { get; set; } = null!;
}