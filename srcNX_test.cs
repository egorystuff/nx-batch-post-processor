using System;
using System.Windows.Forms;
using NXOpen;
using NXOpen.CAM;
using NXOpen.UF;

public class CamProgram_PrintGroupStructure
{
    public static Session theSession;
    public static UI theUI;
    public static Part workPart;
    public static UFSession ufSession;

    public static void Main(string[] args)
    {
        try
        {
            theSession = Session.GetSession();
            theUI = UI.GetUI();
            workPart = theSession.Parts.Work;
            ufSession = UFSession.GetUFSession();

            if (workPart == null || workPart.CAMSetup == null)
            {
                theUI.NXMessageBox.Show("Ошибка", NXMessageBox.DialogType.Error, "Нет активной детали или CAM Setup.");
                return;
            }

            if (theUI.SelectionManager.GetNumSelectedObjects() == 0)
            {
                theUI.NXMessageBox.Show("Выделение", NXMessageBox.DialogType.Warning,
                    "Выберите NCGroup (программу) или операцию внутри неё.");
                return;
            }

            TaggedObject selected = theUI.SelectionManager.GetSelectedTaggedObject(0);
            if (selected == null)
            {
                theUI.NXMessageBox.Show("Ошибка", NXMessageBox.DialogType.Error,
                    "Не удалось получить выбранный объект.");
                return;
            }

            NCGroup startGroup = selected as NCGroup;
            if (startGroup == null)
            {
                NXOpen.CAM.Operation selOp = selected as NXOpen.CAM.Operation;
                if (selOp != null)
                {
                    NCGroup root = workPart.CAMSetup.GetRoot(CAMSetup.View.ProgramOrder);
                    startGroup = FindOwningGroupForOperation(root, selOp);
                }
            }
            if (startGroup == null)
            {
                theUI.NXMessageBox.Show("Выделение", NXMessageBox.DialogType.Warning,
                    "Выбранный объект не является NCGroup или Operation.");
                return;
            }

            theSession.ListingWindow.Open();
            theSession.ListingWindow.WriteLine("=== Структура выбранной CAM группы ===");
            theSession.ListingWindow.WriteLine("Группа: " + SafeName(startGroup));
            theSession.ListingWindow.WriteLine("");

            PrintGroupStructure(startGroup, 0);
        }
        catch (Exception ex)
        {
            theUI.NXMessageBox.Show("Ошибка", NXMessageBox.DialogType.Error, ex.ToString());
        }
    }

    private static void PrintGroupStructure(NCGroup group, int indentLevel)
    {
        if (group == null) return;

        string indent = new string(' ', indentLevel * 2);
        theSession.ListingWindow.WriteLine(indent + "Группа: " + SafeName(group));

        CAMObject[] members = group.GetMembers();
        if (members != null)
        {
            foreach (CAMObject member in members)
            {
                NCGroup childGroup = member as NCGroup;
                if (childGroup != null)
                {
                    PrintGroupStructure(childGroup, indentLevel + 1);
                }
                else
                {
                    NXOpen.CAM.Operation operation = member as NXOpen.CAM.Operation;
                    if (operation != null)
                    {
                        theSession.ListingWindow.WriteLine(indent + "▶ Операция: " + SafeName(operation));
                        theSession.ListingWindow.WriteLine(indent + "  ───────────────────────────────────────────────────────────");

                        // Вывод информации об инструменте
                        NXObject toolObj = operation.ParentMachineTool;
                        if (toolObj != null)
                        {
                            DumpToolInfo(toolObj, indent + "  ");
                        }
                        else
                        {
                            theSession.ListingWindow.WriteLine(indent + "  Инструмент: <нет>");
                        }

                        theSession.ListingWindow.WriteLine("");
                    }
                }
            }
        }
    }

    private static void DumpToolInfo(NXObject toolObj, string indent)
    {
        if (toolObj == null) return;

        string toolName = SafeName(toolObj as CAMObject);
        theSession.ListingWindow.WriteLine(indent + "mom_oper_tool".PadRight(30) + " : " + toolName);

        Tag toolTag = toolObj.Tag;

        // Получаем тип инструмента
        int toolType, toolSubtype;
        ufSession.Cutter.AskTypeAndSubtype(toolTag, out toolType, out toolSubtype);
        theSession.ListingWindow.WriteLine(indent + "Тип инструмента: " + toolType + ", подтип: " + toolSubtype);

        // Пробуем получить параметры инструмента
        DumpAllToolParameters(toolTag, indent);
    }

    private static void DumpAllToolParameters(Tag toolTag, string indent)
    {
        // Пробуем разные диапазоны индексов
        int[] paramIndices = new int[] 
        {
            1038, // UF_PARAM_TL_NUMBER
            1039, // UF_PARAM_TL_DIAMETER
            1040, // UF_PARAM_TL_COR1_RAD
            1041, // UF_PARAM_TL_LENGTH
            1042, // UF_PARAM_TL_FLUTE_LENGTH
            1064, // UF_PARAM_TL_DESCR
            1105, // UF_PARAM_TL_EXT_LENGTH
            1070, // UF_PARAM_TL_POINT_ANG
            1071, 1072, 1073, 1074, 1075,
            1080, 1081, 1082, 1083, 1084, 1085,
            1090, 1091, 1092, 1093, 1094, 1095,
            1100, 1101, 1102, 1103, 1104, 1106, 1107, 1108, 1109, 1110
        };

        string[] paramNames = new string[] 
        {
            "mom_tool_number",
            "mom_tool_diameter",
            "mom_tool_corner_radius",
            "mom_tool_length",
            "mom_tool_flute_length",
            "mom_tool_description",
            "mom_tool_extension_length",
            "mom_tool_tip_angle",
            "param_1071", "param_1072", "param_1073", "param_1074", "param_1075",
            "param_1080", "param_1081", "param_1082", "param_1083", "param_1084", "param_1085",
            "param_1090", "param_1091", "param_1092", "param_1093", "param_1094", "param_1095",
            "param_1100", "param_1101", "param_1102", "param_1103", "param_1104", "param_1106", "param_1107", "param_1108", "param_1109", "param_1110"
        };

        for (int i = 0; i < paramIndices.Length; i++)
        {
            // Пробуем как double
            try
            {
                double dValue;
                ufSession.Param.AskDoubleValue(toolTag, paramIndices[i], out dValue);
                if (Math.Abs(dValue) > 0.0001)
                {
                    theSession.ListingWindow.WriteLine(indent + paramNames[i].PadRight(30) + " : " + dValue.ToString("F3"));
                    continue;
                }
            }
            catch { }

            // Пробуем как int
            try
            {
                int iValue;
                ufSession.Param.AskIntValue(toolTag, paramIndices[i], out iValue);
                if (iValue != 0)
                {
                    theSession.ListingWindow.WriteLine(indent + paramNames[i].PadRight(30) + " : " + iValue.ToString());
                }
            }
            catch { }
        }
    }

    private static string SafeName(CAMObject obj)
    {
        if (obj == null) return "<null>";
        try { return string.IsNullOrEmpty(obj.Name) ? "<unnamed>" : obj.Name; }
        catch { return "<no-name>"; }
    }

    private static NCGroup FindOwningGroupForOperation(NCGroup group, NXOpen.CAM.Operation targetOp)
    {
        if (group == null || targetOp == null) return null;

        CAMObject[] members = group.GetMembers();
        if (members == null) return null;

        foreach (CAMObject m in members)
        {
            NXOpen.CAM.Operation op = m as NXOpen.CAM.Operation;
            if (op != null && op == targetOp)
                return group;

            NCGroup childGroup = m as NCGroup;
            if (childGroup != null)
            {
                NCGroup found = FindOwningGroupForOperation(childGroup, targetOp);
                if (found != null) return found;
            }
        }
        return null;
    }
}