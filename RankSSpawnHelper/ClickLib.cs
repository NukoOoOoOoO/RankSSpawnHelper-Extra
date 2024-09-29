using FFXIVClientStructs.FFXIV.Client.UI;
using FFXIVClientStructs.FFXIV.Component.GUI;
using Reloaded.Memory.Sources;
using Reloaded.Memory.Utilities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using ValueType = FFXIVClientStructs.FFXIV.Component.GUI.ValueType;

namespace RankSSpawnHelper;

public abstract class SharedBuffer
{
    static SharedBuffer()
    {
        Buffer = new(0x2048, Memory.Instance);
    }

    /// <summary>
    /// Gets the shared buffer.
    /// </summary>
    protected static CircularBuffer Buffer { get; }

    /// <summary>
    /// Dispose.
    /// </summary>
    public static void Dispose()
    {
        Buffer.Dispose();
    }
}

public sealed unsafe class EventData : SharedBuffer
{
    /// <summary>
    /// Initializes a new instance of the <see cref="EventData"/> class.
    /// </summary>
    private EventData()
    {
        this.Data = (void**)Buffer.Add(new byte[0x18]);
        if (this.Data == null)
            throw new ArgumentNullException("EventData could not be created, null");

        this.Data[0] = null;
        this.Data[1] = null;
        this.Data[2] = null;
    }

    /// <summary>
    /// Gets the data pointer.
    /// </summary>
    public void** Data { get; }

    /// <summary>
    /// Initializes a new instance of the <see cref="EventData"/> class.
    /// </summary>
    /// <param name="target">Target.</param>
    /// <param name="listener">Event listener.</param>
    /// <returns>Event data.</returns>
    public static EventData ForNormalTarget(void* target, void* listener)
    {
        var data = new EventData();
        data.Data[1] = target;
        data.Data[2] = listener;
        return data;
    }
}

public unsafe sealed class InputData : SharedBuffer
{
    /// <summary>
    /// Initializes a new instance of the <see cref="InputData"/> class.
    /// </summary>
    private InputData()
    {
        this.Data = (void**)Buffer.Add(new byte[0x40]);
        if (this.Data == null)
            throw new ArgumentNullException("InputData could not be created, null");

        this.Data[0] = null;
        this.Data[1] = null;
        this.Data[2] = null;
        this.Data[3] = null;
        this.Data[4] = null;
        this.Data[5] = null;
        this.Data[6] = null;
        this.Data[7] = null;
    }

    /// <summary>
    /// Gets the data pointer.
    /// </summary>
    public void** Data { get; }

    /// <summary>
    /// Initializes a new instance of the <see cref="InputData"/> class.
    /// </summary>
    /// <returns>Input data.</returns>
    public static InputData Empty()
    {
        return new InputData();
    }

    /// <summary>
    /// Initializes a new instance of the <see cref="InputData"/> class.
    /// </summary>
    /// <param name="popupMenu">List popup menu.</param>
    /// <param name="index">Selected index.</param>
    /// <returns>Input data.</returns>
    public static InputData ForPopupMenu(PopupMenu* popupMenu, ushort index)
    {
        var data = new InputData();
        data.Data[0] = popupMenu->List->ItemRendererList[index].AtkComponentListItemRenderer;
        data.Data[2] = (void*)(index | ((ulong)index << 48));
        return data;
    }
}

public enum EventType : ushort
{
#pragma warning disable SA1602
    NORMAL = 1,
    NORMAL_MAX = 2,
    MOUSE_DOWN = 3,
    MOUSE_UP = 4,
    MOUSE_MOVE = 5,
    MOUSE_ROLL_OVER = 6,
    MOUSE_ROLL_OUT = 7,
    MOUSE_WHEEL = 8,
    MOUSE_CLICK = 9,
    MOUSE_DOUBLE_CLICK = 10,
    MOUSE_MAX = 11,
    INPUT = 12,
    INPUT_KEY = 13,
    INPUT_MAX = 14,
    PAD = 15,
    PAD_MAX = 16,
    FOCUS_IN = 17,
    FOCUS_OUT = 18,
    FOCUS_MAX = 19,
    RESIZE = 20,
    RESIZE_MAX = 21,
    BUTTON_PRESS = 22,
    BUTTON_CLICK = 23,
    BUTTON_MAX = 24,
    CHANGE = 25,
    CHANGE_MAX = 26,
    SLIDER_CHANGE = 27,
    SLIDER_CHANGE_END = 28,
    LIST_ITEM_PRESS = 29,
    LIST_ITEM_UP = 30,
    LIST_ITEM_ROLL_OVER = 31,
    LIST_ITEM_ROLL_OUT = 32,
    LIST_ITEM_CLICK = 33,
    LIST_ITEM_DOUBLE_CLICK = 34,
    LIST_INDEX_CHANGE = 35,
    LIST_FOCUS_CHANGE = 36,
    LIST_ITEM_CANCEL = 37,
    LIST_ITEM_PICKUP_START = 38,
    LIST_ITEM_PICKUP_END = 39,
    LIST_ITEM_EXCHANGE = 40,
    LIST_TREE_EXPAND = 41,
    LIST_MAX = 42,
    DDL_LIST_OPEN = 43,
    DDL_LIST_CLOSE = 44,
    DD_DRAG_START = 45,
    DD_DRAG_END = 46,
    DD_DROP = 47,
    DD_DROP_EXCHANGE = 48,
    DD_DROP_NOTICE = 49,
    DD_ROLL_OVER = 50,
    DD_ROLL_OUT = 51,
    DD_DROP_STAGE = 52,
    DD_EXECUTE = 53,
    ICON_TEXT_ROLL_OVER = 54,
    ICON_TEXT_ROLL_OUT = 55,
    ICON_TEXT_CLICK = 56,
    DIALOGUE_CLOSE = 57,
    DIALOGUE_SUBMIT = 58,
    TIMER = 59,
    TIMER_COMPLETE = 60,
    SIMPLETWEEN_UPDATE = 61,
    SIMPLETWEEN_COMPLETE = 62,
    SETUP_ADDON = 63,
    UNIT_BASE_OVER = 64,
    UNIT_BASE_OUT = 65,
    UNIT_SCALE_CHANEGED = 66,
    UNIT_RESOLUTION_SCALE_CHANEGED = 67,
    TIMELINE_STATECHANGE = 68,
    WORDLINK_CLICK = 69,
    WORDLINK_ROLL_OVER = 70,
    WORDLINK_ROLL_OUT = 71,
    CHANGE_TEXT = 72,
    COMPONENT_IN = 73,
    COMPONENT_OUT = 74,
    COMPONENT_SCROLL = 75,
    COMPONENT_FOCUSED = 76, // Maybe
#pragma warning restore SA1602
}

internal static unsafe class ClickLib
{
    internal delegate IntPtr ReceiveEventDelegate(AtkEventListener* eventListener, EventType evt, uint which, void* eventData, void* inputData);

    public static void ClickAddonButton(AtkUnitBase* unitBase, AtkComponentButton* target, uint which, EventType type = EventType.CHANGE)
        => ClickAddonComponent(unitBase, target->AtkComponentBase.OwnerNode, which, type);

    public static void ClickAddonButtonIndex(AtkUnitBase* unitBase, int nodeIndex, uint which, EventType type = EventType.CHANGE)
    {
        var node = (AtkComponentButton*)unitBase->UldManager.NodeList[nodeIndex];
        ClickAddonButton(unitBase, node, which, type);
    }

    private static void ClickAddonComponent(AtkUnitBase* unitBase, AtkComponentNode* target, uint which, EventType type, EventData? eventData = null,
                                            InputData? inputData = null)
    {
        eventData ??= EventData.ForNormalTarget(target, unitBase);
        inputData ??= InputData.Empty();

        InvokeReceiveEvent(&unitBase->AtkEventListener, type, which, eventData, inputData);
    }

    private static void InvokeReceiveEvent(AtkEventListener* eventListener, EventType type, uint which, EventData eventData, InputData inputData)
    {
        var receiveEvent = GetReceiveEvent(eventListener);
        receiveEvent(eventListener, type, which, eventData.Data, inputData.Data);
    }

    private static ReceiveEventDelegate GetReceiveEvent(AtkEventListener* listener)
    {
        var receiveEventAddress = new IntPtr(((nint*)listener->VirtualTable)[2]);
        return Marshal.GetDelegateForFunctionPointer<ReceiveEventDelegate>(receiveEventAddress)!;
    }

    public static void FireCallback(AtkUnitBase* unitBase, params object[] values)
    {
        var atkValues = CreateAtkValueArray(values);
        if (atkValues == null) return;
        try
        {
            unitBase->FireCallback((uint)values.Length, atkValues);
        }
        finally
        {
            for (var i = 0; i < values.Length; i++)
            {
                if (atkValues[i].Type == ValueType.String)
                {
                    Marshal.FreeHGlobal(new(atkValues[i].String));
                }
            }

            Marshal.FreeHGlobal(new(atkValues));
        }
    }

    private static AtkValue* CreateAtkValueArray(params object[] values)
    {
        var atkValues = (AtkValue*)Marshal.AllocHGlobal(values.Length * sizeof(AtkValue));
        if (atkValues == null) return null;
        try
        {
            for (var i = 0; i < values.Length; i++)
            {
                var v = values[i];
                switch (v)
                {
                    case uint uintValue:
                        atkValues[i].Type = ValueType.UInt;
                        atkValues[i].UInt = uintValue;
                        break;
                    case int intValue:
                        atkValues[i].Type = ValueType.Int;
                        atkValues[i].Int = intValue;
                        break;
                    case float floatValue:
                        atkValues[i].Type = ValueType.Float;
                        atkValues[i].Float = floatValue;
                        break;
                    case bool boolValue:
                        atkValues[i].Type = ValueType.Bool;
                        atkValues[i].Byte = (byte)(boolValue ? 1 : 0);
                        break;
                    case string stringValue:
                    {
                        atkValues[i].Type = ValueType.String;
                        var stringBytes = Encoding.UTF8.GetBytes(stringValue);
                        var stringAlloc = Marshal.AllocHGlobal(stringBytes.Length + 1);
                        Marshal.Copy(stringBytes, 0, stringAlloc, stringBytes.Length);
                        Marshal.WriteByte(stringAlloc, stringBytes.Length, 0);
                        atkValues[i].String = (byte*)stringAlloc;
                        break;
                    }
                    default:
                        throw new ArgumentException($"Unable to convert type {v.GetType()} to AtkValue");
                }
            }
        }
        catch
        {
            return null;
        }

        return atkValues;
    }
}