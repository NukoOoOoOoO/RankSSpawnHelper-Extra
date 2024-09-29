using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using Dalamud.Game.Addon.Lifecycle;
using Dalamud.Game.Addon.Lifecycle.AddonArgTypes;
using FFXIVClientStructs.FFXIV.Client.UI;
using FFXIVClientStructs.FFXIV.Component.GUI;

namespace RankSSpawnHelper.Features;

public class JournalStuff : IDisposable
{
    [StructLayout(LayoutKind.Explicit)]
    public struct AddonGuildLeveDifficulty
    {
        [FieldOffset(0)] public AtkUnitBase AtkUnitBase;
    }

    public JournalStuff()
    {
        DalamudApi.AddonLifecycle.RegisterListener(AddonEvent.PostSetup, "GuildLeveDifficulty", AddonGuildLeveDifficultyHandler);
        DalamudApi.AddonLifecycle.RegisterListener(AddonEvent.PostSetup, "JournalDetail", AddonJournalDetailHandler);
        DalamudApi.AddonLifecycle.RegisterListener(AddonEvent.PostDraw, "_Image", AddonImageHandler);
        DalamudApi.AddonLifecycle.RegisterListener(AddonEvent.PostSetup, "SelectYesno", AddonSelectYesnoHandler);
        /*unsafe
        {
            _addonSetupHook = DalamudApi.GameInteropProvider.HookFromSignature<AddonSetupDelegate>("E8 ?? ?? ?? ?? 8B 83 ?? ?? ?? ?? C1 E8 14", hk_AddonSetup);
            _addonSetupHook.Enable();

            _addonUpdate = DalamudApi.GameInteropProvider.HookFromSignature<AddonUpdateDelegate>("48 8B 81 ?? ?? ?? ?? 48 85 C0 74 ?? BA ?? ?? ?? ?? 66 85 90",
                hk_AddonUpdate);
            _addonUpdate.Enable();
        }*/
    }

    public void Dispose()
    {
        DalamudApi.AddonLifecycle.UnregisterListener(AddonEvent.PostSetup, "GuildLeveDifficulty", AddonGuildLeveDifficultyHandler);
        DalamudApi.AddonLifecycle.UnregisterListener(AddonEvent.PostSetup, "JournalDetail", AddonJournalDetailHandler);
        DalamudApi.AddonLifecycle.UnregisterListener(AddonEvent.PostDraw, "_Image", AddonImageHandler);
        DalamudApi.AddonLifecycle.UnregisterListener(AddonEvent.PostSetup, "SelectYesno", AddonSelectYesnoHandler);
    }

    private unsafe void AddonGuildLeveDifficultyHandler(AddonEvent type, AddonArgs args)
    {
        if (!DalamudApi.Configuration._autoJournal)
            return;
        if (DalamudApi.ClientState.TerritoryType != 152 && DalamudApi.ClientState.TerritoryType != 145)
            return;

        var addon = (AddonGuildLeveDifficulty*)args.Addon;
        ClickLib.ClickAddonButtonIndex(&addon->AtkUnitBase, 10, 0);
    }

    private unsafe void ProcessJornalDetail(nint addonPtr)
    {
        var addon = (AddonJournalDetail*)addonPtr;

        var initButton = addon->InitiateButton;
        var declineButton = addon->AbandonDeclineButton;
        if (declineButton == null)
            return;

        if (!declineButton->AtkComponentBase.AtkResNode->IsVisible())
            return;

        var text = Marshal.PtrToStringUTF8((nint)initButton->ButtonTextNode->GetText());
        if (text is "重新挑战" or "开始")
        {
            ClickLib.ClickAddonButtonIndex(&addon->AtkUnitBase, 6, 2);
            return;
        }

        text = Marshal.PtrToStringUTF8((nint)declineButton->ButtonTextNode->GetText());
        if (text is not ("中止" or "取消"))
            return;

        ClickLib.ClickAddonButton(&addon->AtkUnitBase, addon->AbandonDeclineButton, 3);
    }

    private void AddonJournalDetailHandler(AddonEvent type, AddonArgs args)
    {
        if (!DalamudApi.Configuration._autoJournal)
            return;
        if (DalamudApi.ClientState.TerritoryType != 152 && DalamudApi.ClientState.TerritoryType != 145)
            return;

        Task.Run(async () =>
        {
            await Task.Delay(100);
            ProcessJornalDetail(args.Addon);
        });
    }

    private unsafe void AddonImageHandler(AddonEvent type, AddonArgs args)
    {
        if (!DalamudApi.Configuration._autoJournal)
            return;
        if (DalamudApi.ClientState.TerritoryType != 152 && DalamudApi.ClientState.TerritoryType != 145)
            return;

        var addon = (AtkUnitBase*)args.Addon;

        if (!addon->IsVisible)
            return;

        var node = addon->GetNodeById(2);
        if (node->Type != NodeType.Image)
            return;

        var imageNode = (AtkImageNode*)node;
        var textureInfo = imageNode->PartsList->Parts[imageNode->PartId].UldAsset;
        var texType = textureInfo->AtkTexture.TextureType;
        if (texType != TextureType.Resource)
            return;

        var texFileNameStdString =
            &textureInfo->AtkTexture.Resource->TexFileResourceHandle->ResourceHandle.FileName;
        var texString = texFileNameStdString->Length < 16
            ? Marshal.PtrToStringAnsi((IntPtr)texFileNameStdString->Buffer)
            : Marshal.PtrToStringAnsi((IntPtr)texFileNameStdString->BufferPtr);

        if (texString != null && (texString.Contains("120033") || texString.Contains("120035")))
            addon->IsVisible = false;
    }

    private unsafe void AddonSelectYesnoHandler(AddonEvent type, AddonArgs args)
    {
        if (!DalamudApi.Configuration._autoJournal)
            return;
        if (DalamudApi.ClientState.TerritoryType != 152 && DalamudApi.ClientState.TerritoryType != 145)
            return;
        var addon = (AddonSelectYesno*)args.Addon;
        if (addon == null)
            return;
        var text = Marshal.PtrToStringUTF8((nint)addon->PromptText->GetText()) ?? "empty";
        if (!text.Contains("要开始理符任务吗？") && !text.Contains("要中止当前任务吗？") && !text.Contains("要重新挑战理符任务吗？"))
            return;

        ClickLib.ClickAddonButton(&addon->AtkUnitBase, addon->YesButton, 0);
    }
}