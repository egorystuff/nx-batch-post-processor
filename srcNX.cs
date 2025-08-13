// CamProgramStructureViewer_FromSelection_FixedOperation.cs
using System;
using System.Text;
using NXOpen;
using NXOpen.CAM;

public class CamProgramStructureViewer_FromSelection_FixedOperation
{
    public static Session theSession;
    public static UI theUI;
    public static Part workPart;

    public static void Main(string[] args)
    {
        try
        {
            theSession = Session.GetSession();
            theUI = UI.GetUI();
            workPart = theSession.Parts.Work;

            theSession.ListingWindow.Open();

            if (workPart == null)
            {
                theSession.ListingWindow.WriteLine("Нет активной детали (work part).");
                return;
            }

            if (workPart.CAMSetup == null)
            {
                theSession.ListingWindow.WriteLine("CAM-сессия не загружена в текущем файле.");
                return;
            }

            CAMSetup setup = workPart.CAMSetup;

            // 1) Берём первый выделенный объект
            int selCount = 0;
            try { selCount = theUI.SelectionManager.GetNumSelectedObjects(); } catch { selCount = 0; }

            if (selCount == 0)
            {
                theUI.NXMessageBox.Show("Выделение",
                    NXMessageBox.DialogType.Warning,
                    "Ничего не выделено в CAM Navigator. Выберите программу (NCGroup) или операцию внутри неё.");
                return;
            }

            TaggedObject selected = theUI.SelectionManager.GetSelectedTaggedObject(0);

            // 2) Определяем стартовую группу:
            NCGroup startGroup = selected as NCGroup;
            if (startGroup == null)
            {
                NXOpen.CAM.Operation selOp = selected as NXOpen.CAM.Operation;
                if (selOp == null)
                {
                    theUI.NXMessageBox.Show("CAM Program Structure",
                        NXMessageBox.DialogType.Warning,
                        "Выберите NCGroup (программу) или операцию внутри программы.");
                    return;
                }

                NCGroup root = null;
                try { root = setup.GetRoot(CAMSetup.View.ProgramOrder); } catch { root = null; }

                if (root == null)
                {
                    theUI.NXMessageBox.Show("CAM Program Structure",
                        NXMessageBox.DialogType.Error,
                        "Не удалось получить корень ProgramOrder.");
                    return;
                }

                startGroup = FindOwningGroupForOperation(root, selOp);
                if (startGroup == null)
                {
                    theUI.NXMessageBox.Show("CAM Program Structure",
                        NXMessageBox.DialogType.Warning,
                        "Не удалось найти группу, содержащую выбранную операцию.");
                    return;
                }
            }

            // 3) Печатаем ТОЛЬКО структуру групп, начиная с найденной
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("=== STRUCTURE (FROM SELECTED GROUP) ===");
            sb.AppendLine();
            sb.AppendLine("- " + SafeName(startGroup));
            PrintChildGroups(startGroup, sb, 1);

            theSession.ListingWindow.WriteLine(sb.ToString());

            // theUI.NXMessageBox.Show("CAM Program Structure",
            //     NXMessageBox.DialogType.Information,
            //     "Структура выбранной программы/группы выведена в Listing Window.");



            
        }
        catch (Exception ex)
        {
            UI.GetUI().NXMessageBox.Show("CamProgramStructureViewer - Error",
                NXMessageBox.DialogType.Error, ex.ToString());
        }
    }

    /// <summary>
    /// Рекурсивный поиск ближайшей группы, которая содержит данную операцию.
    /// </summary>
    private static NCGroup FindOwningGroupForOperation(NCGroup group, NXOpen.CAM.Operation targetOp)
    {
        if (group == null || targetOp == null) return null;

        CAMObject[] members = null;
        try { members = group.GetMembers(); } catch { members = null; }

        if (members == null || members.Length == 0) return null;

        // Сначала проверим прямых детей
        for (int i = 0; i < members.Length; i++)
        {
            CAMObject m = members[i];
            if (m == null) continue;

            NXOpen.CAM.Operation op = m as NXOpen.CAM.Operation;
            if (op != null)
            {
                if (object.ReferenceEquals(op, targetOp) || op == targetOp)
                    return group;
            }
        }

        // Затем углубляемся в подгруппы
        for (int i = 0; i < members.Length; i++)
        {
            CAMObject m = members[i];
            if (m == null) continue;

            NCGroup childGroup = m as NCGroup;
            if (childGroup != null)
            {
                NCGroup found = FindOwningGroupForOperation(childGroup, targetOp);
                if (found != null) return found;
            }
        }

        return null;
    }

    /// <summary>
    /// Печатает только дочерние NCGroup’ы рекурсивно (операции и прочее игнорируются).
    /// </summary>
    private static void PrintChildGroups(NCGroup group, StringBuilder sb, int indent)
    {
        if (group == null) return;

        CAMObject[] members = null;
        try { members = group.GetMembers(); } catch { members = null; }

        if (members == null || members.Length == 0) return;

        string prefix = new string(' ', indent * 2);

        for (int i = 0; i < members.Length; i++)
        {
            CAMObject camObj = members[i];
            if (camObj == null) continue;

            NCGroup childGroup = camObj as NCGroup;
            if (childGroup != null)
            {
                sb.AppendLine(prefix + "- " + SafeName(childGroup));
                PrintChildGroups(childGroup, sb, indent + 1);
            }
        }
    }

    private static string SafeName(CAMObject obj)
    {
        try
        {
            return string.IsNullOrEmpty(obj.Name) ? "<unnamed>" : obj.Name;
        }
        catch
        {
            return "<no-name>";
        }
    }

    private static string SafeName(NCGroup grp)
    {
        try
        {
            return string.IsNullOrEmpty(grp.Name) ? "<unnamed-group>" : grp.Name;
        }
        catch
        {
            return "<no-name-group>";
        }
    }
}
