// CamStructureViewer.cs
using System;
using System.Text;
using NXOpen;
using NXOpen.CAM;

public class CamStructureViewer
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
            StringBuilder sb = new StringBuilder();

            sb.AppendLine("=== CAM STRUCTURE ===");
            sb.AppendLine();

            // Перечислим view, которые хотим обойти
            CAMSetup.View[] views = new CAMSetup.View[] {
                CAMSetup.View.ProgramOrder,
                // CAMSetup.View.MachineMethod,
                // CAMSetup.View.Geometry,
                // CAMSetup.View.MachineTool
            };

            foreach (CAMSetup.View view in views)
            {
                sb.AppendLine("---- VIEW: " + view.ToString() + " ----");
                NCGroup root = null;
                try
                {
                    root = setup.GetRoot(view);
                }
                catch
                {
                    root = null;
                }

                if (root == null)
                {
                    sb.AppendLine("  <root is empty or inaccessible>");
                }
                else
                {
                    PrintNCGroup(root, sb, 0);
                }

                sb.AppendLine();
            }

            // Пишем в Listing Window (удобно для копирования/длинных списков)
            theSession.ListingWindow.WriteLine(sb.ToString());

            // Короткое подтверждение в модальном окне
            theUI.NXMessageBox.Show("CAM Structure", NXMessageBox.DialogType.Information,
                "Структура CAM выведена в Listing Window.");
        }
        catch (Exception ex)
        {
            UI.GetUI().NXMessageBox.Show("CamStructureViewer - Error", NXMessageBox.DialogType.Error, ex.ToString());
        }
    }

    // Рекурсивный обход NCGroup
    private static void PrintNCGroup(NCGroup group, StringBuilder sb, int indent)
    {
        if (group == null) return;

        string prefix = new string(' ', indent * 2);
        string groupName = SafeName(group);
        sb.AppendLine(prefix + "- " + groupName );

        CAMObject[] members = null;
        try { members = group.GetMembers(); }
        catch { members = null; }

        if (members == null || members.Length == 0) return;

        foreach (CAMObject camObj in members)
        {
            if (camObj == null) continue;

            string objName = SafeName(camObj);
            string typeName = camObj.GetType().Name;

            // Если вложенная группа — рекурсивно углубляемся
            if (camObj is NCGroup)
            {
                PrintNCGroup((NCGroup)camObj, sb, indent + 1);
            }
     
        
        }
    }

    // Безопасное чтение имени объекта CAM (Name у CAMObject обычно доступно)
    private static string SafeName(CAMObject obj)
    {
        try
        {
            // CAMObject наследует NXObject, у которого есть Name
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
