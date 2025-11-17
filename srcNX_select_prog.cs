// CamProgram_BatchPostprocess_Single_3Axis.cs
//
// Постпроцессинг только одной выделенной программы (NCGroup) или операции.
// Обрабатывается ТОЛЬКО один объект, без дочерних элементов.
// Используются только 3-х осевые постпроцессоры.
//


using System;
using System.Windows.Forms;
using SysIO = System.IO;
using NXOpen;
using NXOpen.CAM;

public class CamProgram_BatchPostprocess_Single_3Axis
{
    public static Session theSession;
    public static UI theUI;
    public static Part workPart;

    private class PostConfig
    {
        public string PostName;
        public string Extension;
        public PostConfig(string postName, string extension)
        {
            PostName = postName;
            Extension = extension;
        }
    }

    // --- Набор постпроцессоров для 3-х осевой обработки ---
    private static readonly PostConfig[] Posts3Axis = new PostConfig[]
    {
        new PostConfig("DMU-60T", ".i"),
        new PostConfig("OKUMA_MB-46VAE_3X", ".min"),
        new PostConfig("HAAS-VF2", ".txt")
    };

    // Опция перезаписи
    private static int overwriteChoice = -1;

    public static void Main(string[] args)
    {
        try
        {
            theSession = Session.GetSession();
            theUI = UI.GetUI();
            workPart = theSession.Parts.Work;

            if (workPart == null || workPart.CAMSetup == null)
            {
                theUI.NXMessageBox.Show("Ошибка", NXMessageBox.DialogType.Error,
                    "Нет активной детали или CAM Setup.");
                return;
            }

            if (theUI.SelectionManager.GetNumSelectedObjects() == 0)
            {
                theUI.NXMessageBox.Show("Выделение", NXMessageBox.DialogType.Warning,
                    "Выберите NCGroup (программу) или операцию.");
                return;
            }

            TaggedObject selected = theUI.SelectionManager.GetSelectedTaggedObject(0);
            if (selected == null)
            {
                theUI.NXMessageBox.Show("Ошибка", NXMessageBox.DialogType.Error,
                    "Не удалось получить выбранный объект.");
                return;
            }

            // Определяем объект для постпроцессинга
            NCGroup targetGroup = selected as NCGroup;
            if (targetGroup == null)
            {
                NXOpen.CAM.Operation selOp = selected as NXOpen.CAM.Operation;
                if (selOp != null)
                {
                    NCGroup root = workPart.CAMSetup.GetRoot(CAMSetup.View.ProgramOrder);
                    targetGroup = FindOwningGroupForOperation(root, selOp);
                }
            }

            if (targetGroup == null)
            {
                theUI.NXMessageBox.Show("Ошибка", NXMessageBox.DialogType.Error,
                    "Не удалось определить программу для постпроцессинга.");
                return;
            }

            // Диалог выбора директории для сохранения
            string defaultDir;
            if (!string.IsNullOrEmpty(workPart.FullPath))
            {
                try { defaultDir = SysIO.Path.GetDirectoryName(workPart.FullPath); }
                catch { defaultDir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); }
            }
            else
                defaultDir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            string outputDir = SelectOutputFolder(defaultDir);
            if (string.IsNullOrEmpty(outputDir)) return;

            string fullName = SafeName(targetGroup);
            string shortName = fullName.Split('_')[0];

            theSession.ListingWindow.Open();
            theSession.ListingWindow.WriteLine("=== Single Postprocess (3-axis) ===");
            theSession.ListingWindow.WriteLine("Программа: " + fullName);
            theSession.ListingWindow.WriteLine("Папка вывода: " + outputDir);
            theSession.ListingWindow.WriteLine("");

            CAMSetup setup = workPart.CAMSetup;

            foreach (PostConfig cfg in Posts3Axis)
            {
                string outFile = SysIO.Path.Combine(outputDir, shortName + cfg.Extension);

                if (SysIO.File.Exists(outFile))
                {
                    if (!HandleOverwrite(outFile)) continue;
                }

                try
                {
                    CAMObject[] toPost = new CAMObject[] { targetGroup };
                    setup.PostprocessWithSetting(
                        toPost,
                        cfg.PostName,
                        outFile,
                        CAMSetup.OutputUnits.Metric,
                        CAMSetup.PostprocessSettingsOutputWarning.PostDefined,
                        CAMSetup.PostprocessSettingsReviewTool.PostDefined
                    );
                    theSession.ListingWindow.WriteLine("   ✔ " + cfg.PostName + " -> " + SysIO.Path.GetFileName(outFile));
                }
                catch (Exception exPost)
                {
                    theSession.ListingWindow.WriteLine("   ✘ Ошибка (" + cfg.PostName + "): " + exPost.Message);
                }
            }

            theUI.NXMessageBox.Show("Готово", NXMessageBox.DialogType.Information,
                "Постпроцессинг завершён.\nФайлы сохранены в:\n" + outputDir);
        }
        catch (Exception ex)
        {
            UI.GetUI().NXMessageBox.Show("Ошибка", NXMessageBox.DialogType.Error, ex.ToString());
        }
    }

    private static string SelectOutputFolder(string defaultDir)
    {
        using (SaveFileDialog sfd = new SaveFileDialog())
        {
            sfd.Title = "Выберите папку для сохранения NC-программ";
            sfd.InitialDirectory = defaultDir;
            sfd.FileName = "Укажите путь";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try { return SysIO.Path.GetDirectoryName(sfd.FileName); }
                catch { return defaultDir; }
            }
        }
        return null;
    }

    private static bool HandleOverwrite(string filePath)
    {
        if (overwriteChoice == -1)
        {
            DialogResult res = MessageBox.Show(
                "Файл:\n" + filePath + "\n\nуже существует.\nПерезаписать?\n\n" +
                "Да = перезаписывать все\nНет = не перезаписывать\nОтмена = спрашивать каждый раз",
                "Перезапись файлов",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question);

            if (res == DialogResult.Yes) overwriteChoice = 1;
            else if (res == DialogResult.No) overwriteChoice = 0;
            else overwriteChoice = -2;
        }

        if (overwriteChoice == 0) return false;

        if (overwriteChoice == -2)
        {
            DialogResult res2 = MessageBox.Show(
                "Файл:\n" + filePath + "\n\nуже существует. Перезаписать?",
                "Перезапись файла",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (res2 != DialogResult.Yes) return false;
        }

        return true;
    }

    private static NCGroup FindOwningGroupForOperation(NCGroup group, NXOpen.CAM.Operation targetOp)
    {
        if (group == null || targetOp == null) return null;

        foreach (CAMObject m in group.GetMembers())
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

    private static string SafeName(CAMObject obj)
    {
        if (obj == null) return "<null>";
        try { return string.IsNullOrEmpty(obj.Name) ? "<unnamed>" : obj.Name; }
        catch { return "<no-name>"; }
    }
}
