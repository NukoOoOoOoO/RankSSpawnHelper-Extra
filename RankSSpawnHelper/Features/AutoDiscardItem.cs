using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Dalamud.Game.Addon.Lifecycle;
using Dalamud.Game.Addon.Lifecycle.AddonArgTypes;
using Dalamud.Hooking;
using Dalamud.Logging;
using FFXIVClientStructs.FFXIV.Client.Game;
using FFXIVClientStructs.FFXIV.Client.UI;
using FFXIVClientStructs.FFXIV.Client.UI.Agent;
using FFXIVClientStructs.FFXIV.Component.GUI;
using Lumina.Excel.GeneratedSheets;
using ValueType = FFXIVClientStructs.FFXIV.Component.GUI.ValueType;

namespace RankSSpawnHelper.Features;

public class ItemInfo
{
    public uint id;
    public string name;

    public ItemInfo(uint id, string name)
    {
        this.id = id;
        this.name = name;
    }
}

public class AutoDiscardItem : IDisposable
{
    // Credit goes to SimpleTweaks https://github.com/Caraxi/SimpleTweaksPlugin/blob/0ddc29b5ceb4942a0234de7f3e0f34eb981f8e3b/Tweaks/QuickSellItems.cs
    private static Hook<OpenInventoryContext> _openInventoryContextHook;
    private bool _discarded;
    public List<ItemInfo> ItemInfos = new();

    public unsafe AutoDiscardItem()
    {
        DalamudApi.AddonLifecycle.RegisterListener(AddonEvent.PostSetup, "InputNumeric", AddonInputNumericHandler);
        DalamudApi.AddonLifecycle.RegisterListener(AddonEvent.PostSetup, "SelectYesno", AddonSelectYesnoHandler);

        _openInventoryContextHook = DalamudApi.GameInteropProvider.HookFromSignature<OpenInventoryContext>("83 B9 ?? ?? ?? ?? ?? 7E ?? 39 91", hk_OpenInventoryContext);
        _openInventoryContextHook.Enable();


        ItemInfos.AddRange(DalamudApi.DataManager.GetExcelSheet<Item>().Where(
                                                                              i => !string.IsNullOrEmpty(i.Name) &&
                                                                                   (
                                                                                       (i.FilterGroup == 4 && i.LevelItem.Value.RowId == 1 &&
                                                                                        !i.IsUnique) || // 普通装备且装等为1的物品 比如草布马裤，超级米饭的斗笠
                                                                                       (i.FilterGroup == 12 && i.RowId != 36256 &&
                                                                                        i.RowId != 27850 && i.RowId != 7767) || // 材料比如矮人棉，庵摩罗果等 但因为秧鸡胸脯肉，厄尔庇斯鸟蛋和鱼粉是特定地图扔的，所以不会加进列表里
                                                                                       i.FilterGroup == 17 // 鱼饵，比如沙蚕
                                                                                   )
                                                                             ).Select(i => new ItemInfo(i.RowId, i.Name)));

        DalamudApi.PluginLog.Debug($"Finished loading {ItemInfos.Count} items");
    }

    private unsafe void AddonInputNumericHandler(AddonEvent type, AddonArgs args)
    {
        if (DalamudApi.ClientState.TerritoryType != 961 && DalamudApi.ClientState.TerritoryType != 813 && DalamudApi.ClientState.TerritoryType != 621 && DalamudApi.ClientState.TerritoryType != 1189)
            return;

        if (!DalamudApi.Configuration._autoDiscardItem)
            return;

        var addon = (AtkUnitBase*)args.Addon;
        if (addon->ContextMenuParentId == 0)
            return;

        // 检查附属的是哪个addon
        var parentAddon = RaptureAtkUnitManager.Instance()->GetAddonById(addon->ContextMenuParentId);
        var parentAddonName = Encoding.UTF8.GetString(parentAddon->Name);

        // PluginLog.Debug($"{parentAddonName}");

        // 傻逼SE我操你妈
        if (!parentAddonName.StartsWith("Inventory"))
            return;

        try
        {
            var inventoryContext = AgentInventoryContext.Instance();

            var itemId = inventoryContext->TargetDummyItem.ItemId;
            if (!IsAllowedToDiscard(itemId))
                return;

            var isEgg = itemId == 36256;
            var isFishMeal = itemId == 7767;
            
            if (isEgg && inventoryContext->TargetDummyItem.Quantity < 5)
                return;
            if (isFishMeal && inventoryContext->TargetDummyItem.Quantity < 50)
                return;

            var numericAddon = (AtkComponentNumericInput*)addon->UldManager.NodeList[4]->GetComponent();

            var numVal = isEgg ? 5 : isFishMeal ? 50: 1;
            // clamp
            if (isEgg)
                numVal = Math.Min(5, numericAddon->Data.Max);
            else if (isFishMeal)
                numVal = Math.Min(50, numericAddon->Data.Max);

            numericAddon->SetValue(numVal);
            var confirmButton = (AtkComponentButton*)addon->UldManager.NodeList[3]->GetComponent();
            ClickLib.ClickAddonButton(addon, confirmButton, 0);
        }
        catch (Exception e)
        {
            DalamudApi.PluginLog.Error(e.ToString());
        }
    }

    public void Dispose()
    {
        DalamudApi.AddonLifecycle.UnregisterListener(AddonEvent.PostSetup, "InputNumeric", AddonInputNumericHandler);
        DalamudApi.AddonLifecycle.UnregisterListener(AddonEvent.PostSetup, "SelectYesno", AddonSelectYesnoHandler);

        _openInventoryContextHook.Dispose();
    }

    private unsafe void AddonSelectYesnoHandler(AddonEvent type, AddonArgs args)
    {
        if (!_discarded)
            return;
        _discarded = false;
        var addon = (AddonSelectYesno*)args.Addon;
        DalamudApi.PluginLog.Debug($"addon->AtkUnitBase.ParentID: {addon->AtkUnitBase.ParentId}");
        ClickLib.ClickAddonButton(&addon->AtkUnitBase, addon->YesButton, 0);
    }

    private static bool IsAllowedToDiscard(uint id)
    {
        var item = DalamudApi.DataManager.Excel.GetSheet<Item>()?.GetRow(id);
        if (item == null)
            return false;

        return DalamudApi.Configuration._itemsToDiscard.Contains(id) || 
               (DalamudApi.ClientState.TerritoryType == 961 && id == 36256) ||
               (DalamudApi.ClientState.TerritoryType == 813 && id == 27850) || 
               (DalamudApi.ClientState.TerritoryType == 1189 && id == 7767) ;
    }

    private unsafe void* hk_OpenInventoryContext(AgentInventoryContext* agent, InventoryType inventoryType, ushort slot, int a4, ushort a5, byte a6)
    {
        var original = _openInventoryContextHook.Original(agent, inventoryType, slot, a4, a5, a6);

        if (!DalamudApi.Configuration._autoDiscardItem)
            return original;

        // Elips 雷克兰德 湖区 树海
        if (DalamudApi.ClientState.TerritoryType != 961 && DalamudApi.ClientState.TerritoryType != 813 && DalamudApi.ClientState.TerritoryType != 621 && DalamudApi.ClientState.TerritoryType != 1189)
            return original;

        if (DalamudApi.GameGui.GetAddonByName("InventoryBuddy") != IntPtr.Zero || DalamudApi.GameGui.GetAddonByName("InventoryEvent") != IntPtr.Zero ||
            DalamudApi.GameGui.GetAddonByName("ArmouryBoard") != IntPtr.Zero)
            return original;

        var inventory = InventoryManager.Instance()->GetInventoryContainer(inventoryType);
        if (inventory == null)
            return original;

        var itemSlot = inventory->GetInventorySlot(slot);
        if (itemSlot == null)
            return original;

        var itemId = itemSlot->ItemId;
        var itemQuantity = itemSlot->Quantity;

        if (!IsAllowedToDiscard(itemId))
            return original;

        switch (DalamudApi.ClientState.TerritoryType)
        {
            case 961 when itemId != 36256:
            case 813 when itemId != 27850:
            case 1189 when itemId != 7767:
                return original;
        }

        var addonId = agent->AgentInterface.GetAddonId();
        if (addonId == 0)
            return original;

        var addon = RaptureAtkUnitManager.Instance()->GetAddonById((ushort)addonId);
        if (addon == null)
            return original;

        for (var i = 0; i < agent->ContextItemCount; i++)
        {
            var contextItemParam = agent->EventParams[agent->ContexItemStartIndex + i];
            if (contextItemParam.Type != ValueType.String)
                continue;

            var contextItemName = Marshal.PtrToStringUTF8(new(contextItemParam.String));
            if (contextItemName == null)
                continue;

            // PluginLog.Debug($"index:{i} | contextItemName: {contextItemName}");

            switch (contextItemName)
            {
                case "拆分":
                    switch (DalamudApi.ClientState.TerritoryType)
                    {
                        case 961 when itemQuantity <= 5:
                        case 813 when itemQuantity == 1:
                        case 621 when itemQuantity == 1:
                        case 1189 when itemQuantity <= 50:
                            continue;
                        default:
                            ClickLib.FireCallback(addon, 0, i, 0, 0);
                            return original;
                    }

                case "舍弃":
                    switch (DalamudApi.ClientState.TerritoryType)
                    {
                        // 湖区只能扔一件
                        case 961 when itemQuantity != 5:
                        case 813 when itemQuantity != 1:
                        case 621 when itemQuantity != 1:
                        case 1189 when itemQuantity != 50:
                            continue;
                    }

                    ClickLib.FireCallback(addon, 0, i, 0, 0);
                    _discarded = true;
                    return original;
            }
        }

        return original;
    }

    private unsafe delegate void* OpenInventoryContext(AgentInventoryContext* agent, InventoryType inventory, ushort slot, int a4, ushort a5, byte a6);
}