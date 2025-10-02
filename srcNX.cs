// CamProgram_BatchPostprocess_3Posts_SelectFolder_Modern.cs
// Для каждой подпрограммы выбранной группы выполняет постпроцессинг тремя постпроцессорами.
// Перед запуском выводится современное диалоговое окно выбора папки для сохранения.
// Файлы сохраняются с расширениями .i, .min, .txt
// Опция перезаписи: Да / Нет / Спрашивать каждый раз.

using System;
using System.Windows.Forms;
using SysIO = System.IO;
using NXOpen;
using NXOpen.CAM;

public class CamProgram_BatchPostprocess_3Posts_SelectFolder_Modern
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

    // Список постпроцессоров
    private static readonly PostConfig[] PostConfigs = new PostConfig[]
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

            // Определяем стартовую группу
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

            // Диалог выбора папки
            string defaultDir = "";
            if (!string.IsNullOrEmpty(workPart.FullPath))
            {
                try { defaultDir = SysIO.Path.GetDirectoryName(workPart.FullPath); }
                catch { defaultDir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop); }
            }
            else
            {
                defaultDir = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            }

            string outputDir = SelectOutputFolder(defaultDir);
            if (string.IsNullOrEmpty(outputDir)) return;

            // Получаем членов группы
            CAMObject[] members = startGroup.GetMembers();
            if (members == null || members.Length == 0)
            {
                theUI.NXMessageBox.Show("Инфо", NXMessageBox.DialogType.Warning,
                    "Выбранная группа не содержит дочерних элементов.");
                return;
            }

            theSession.ListingWindow.Open();
            theSession.ListingWindow.WriteLine("=== Batch postprocess (3 posts) ===");
            theSession.ListingWindow.WriteLine("Группа: " + SafeName(startGroup));
            theSession.ListingWindow.WriteLine("Папка вывода: " + outputDir);
            theSession.ListingWindow.WriteLine("");

            CAMSetup setup = workPart.CAMSetup;

            foreach (CAMObject camObj in members)
            {
                NCGroup childGroup = camObj as NCGroup;
                if (childGroup == null) continue;

                string fullName = SafeName(childGroup);
                string shortName = fullName.Split('_')[0]; 

                foreach (PostConfig cfg in PostConfigs)
                {
                    string outFile = SysIO.Path.Combine(outputDir, shortName + cfg.Extension);

                    if (SysIO.File.Exists(outFile))
                    {
                        if (!HandleOverwrite(outFile)) continue;
                    }

                    try
                    {
                        CAMObject[] toPost = new CAMObject[] { childGroup };
                        setup.PostprocessWithSetting(
                            toPost,
                            cfg.PostName,
                            outFile,
                            CAMSetup.OutputUnits.Metric,
                            CAMSetup.PostprocessSettingsOutputWarning.PostDefined,
                            CAMSetup.PostprocessSettingsReviewTool.PostDefined
                        );
                        theSession.ListingWindow.WriteLine("✔ " + fullName + " (" + cfg.PostName + ") -> " + outFile);
                    }
                    catch (Exception exPost)
                    {
                        theSession.ListingWindow.WriteLine("✘ Ошибка (" + cfg.PostName + ") для " + fullName + ": " + exPost.Message);
                    }
                }
            }

            theUI.NXMessageBox.Show("Готово", NXMessageBox.DialogType.Information,
                "Постпроцессинг завершён. Файлы сохранены в папке:\n" + outputDir);
        }
        catch (Exception ex)
        {
            UI.GetUI().NXMessageBox.Show("Ошибка", NXMessageBox.DialogType.Error, ex.ToString());
        }
    }

    // Метод выбора папки через SaveFileDialog (удобное стандартное окно Windows)
    private static string SelectOutputFolder(string defaultDir)
    {
        using (SaveFileDialog sfd = new SaveFileDialog())
        {
            sfd.Title = "Выберите папку для сохранения NC-программ";
            sfd.InitialDirectory = defaultDir;
            sfd.FileName = "Укажите путь"; // фиктивное имя

            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    return SysIO.Path.GetDirectoryName(sfd.FileName);
                }
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

    private static string SafeName(CAMObject obj)
    {
        if (obj == null) return "<null>";
        try { return string.IsNullOrEmpty(obj.Name) ? "<unnamed>" : obj.Name; }
        catch { return "<no-name>"; }
    }
}
