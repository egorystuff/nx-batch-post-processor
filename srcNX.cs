// CamProgram_BatchPostprocess_3Posts_Fixed.cs
// Для каждой подпрограммы выбранной группы выполняет постпроцессинг тремя постпроцессорами.
// Файлы сохраняются на рабочем столе в папку "Prog output" с расширениями .i, .min, .txt.
// Опция перезаписи: Да / Нет / Спрашивать каждый раз.

using System;
using System.Windows.Forms;
using SysIO = System.IO;
using NXOpen;
using NXOpen.CAM;

public class CamProgram_BatchPostprocess_3Posts_Fixed
{
    public static Session theSession;
    public static UI theUI;
    public static Part workPart;

    // Небольший класс вместо кортежа — чтобы гарантированно компилировалось в старых средах
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

    // Жёстко заданные постпроцессоры и расширения
    private static readonly PostConfig[] PostConfigs = new PostConfig[]
    {
        new PostConfig("DMU-60T", ".i"),
        new PostConfig("OKUMA_MB-46VAE_3X", ".min"),
        new PostConfig("HAAS-VF2", ".txt")
    };

    // Опция перезаписи: -1 = ещё не спрашивали, 0 = никогда, 1 = всегда, -2 = спрашивать каждый раз
    private static int overwriteChoice = -1;

    public static void Main(string[] args)
    {
        try
        {
            theSession = Session.GetSession();
            theUI = UI.GetUI();
            workPart = theSession.Parts.Work;

            if (workPart == null)
            {
                theUI.NXMessageBox.Show("Ошибка", NXMessageBox.DialogType.Error, "Нет активной детали (work part).");
                return;
            }

            if (workPart.CAMSetup == null)
            {
                theUI.NXMessageBox.Show("Ошибка", NXMessageBox.DialogType.Error, "CAM-сессия не загружена в текущем файле.");
                return;
            }

            // 1) Получаем выбранный объект (первый)
            int selCount = 0;
            try { selCount = theUI.SelectionManager.GetNumSelectedObjects(); } catch { selCount = 0; }

            if (selCount == 0)
            {
                theUI.NXMessageBox.Show("Выделение", NXMessageBox.DialogType.Warning,
                    "Ничего не выделено. Выберите NCGroup (программу) или операцию внутри неё.");
                return;
            }

            TaggedObject selected = null;
            try { selected = theUI.SelectionManager.GetSelectedTaggedObject(0); } catch { selected = null; }
            if (selected == null)
            {
                theUI.NXMessageBox.Show("Ошибка", NXMessageBox.DialogType.Error, "Не удалось получить выбранный объект.");
                return;
            }

            // 2) Определяем стартовую NCGroup: если выделена операция — ищем родительскую группу
            NCGroup startGroup = selected as NCGroup;
            if (startGroup == null)
            {
                NXOpen.CAM.Operation selOp = selected as NXOpen.CAM.Operation;
                if (selOp != null)
                {
                    NCGroup root = null;
                    try { root = workPart.CAMSetup.GetRoot(CAMSetup.View.ProgramOrder); } catch { root = null; }

                    if (root == null)
                    {
                        theUI.NXMessageBox.Show("Ошибка", NXMessageBox.DialogType.Error, "Не удалось получить корень ProgramOrder.");
                        return;
                    }

                    startGroup = FindOwningGroupForOperation(root, selOp);
                    if (startGroup == null)
                    {
                        theUI.NXMessageBox.Show("Ошибка", NXMessageBox.DialogType.Warning, "Не удалось найти группу, содержащую выбранную операцию.");
                        return;
                    }
                }
                else
                {
                    theUI.NXMessageBox.Show("Выделение", NXMessageBox.DialogType.Warning, "Выбранный объект не является NCGroup или Operation.");
                    return;
                }
            }

            // 3) Создаём папку вывода на рабочем столе
            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string outputDir = SysIO.Path.Combine(desktop, "Prog output");

            try
            {
                SysIO.Directory.CreateDirectory(outputDir);
            }
            catch (Exception exDir)
            {
                theUI.NXMessageBox.Show("Ошибка папки", NXMessageBox.DialogType.Error, "Не удалось создать папку вывода: " + exDir.Message);
                return;
            }

            // 4) Получаем прямых детей выбранной группы и запускаем для каждого 3 поста
            CAMObject[] members = null;
            try { members = startGroup.GetMembers(); } catch { members = null; }

            if (members == null || members.Length == 0)
            {
                theUI.NXMessageBox.Show("Инфо", NXMessageBox.DialogType.Warning, "Выбранная группа не содержит дочерних элементов.");
                return;
            }

            theSession.ListingWindow.Open();
            theSession.ListingWindow.WriteLine("=== Batch postprocess (3 posts) для группы: " + SafeName(startGroup) + " ===");
            theSession.ListingWindow.WriteLine("Папка вывода: " + outputDir);
            theSession.ListingWindow.WriteLine("");

            CAMSetup setup = workPart.CAMSetup;

            foreach (CAMObject camObj in members)
            {
                if (camObj == null) continue;

                NCGroup childGroup = camObj as NCGroup;
                if (childGroup == null) continue;

                string fullName = SafeName(childGroup);

                // Имя до первого '_' или вся строка, если '_' нет
                string shortName;
                int idx = fullName.IndexOf('_');
                if (idx > 0)
                    shortName = fullName.Substring(0, idx);
                else
                    shortName = fullName;

                foreach (PostConfig cfg in PostConfigs)
                {
                    string outFile = SysIO.Path.Combine(outputDir, shortName + cfg.Extension);

                    // Проверка существования и логика перезаписи
                    if (SysIO.File.Exists(outFile))
                    {
                        if (overwriteChoice == -1) // ещё не спрашивали
                        {
                            DialogResult res = System.Windows.Forms.MessageBox.Show(
                                "Файл:\n" + outFile + "\n\nуже существует.\nПерезаписать?\n\n" +
                                "Да = перезаписывать все\nНет = не перезаписывать\nОтмена = спрашивать каждый раз",
                                "Перезапись файлов",
                                MessageBoxButtons.YesNoCancel,
                                MessageBoxIcon.Question);

                            if (res == DialogResult.Yes) overwriteChoice = 1;
                            else if (res == DialogResult.No) overwriteChoice = 0;
                            else overwriteChoice = -2; // спрашивать каждый раз
                        }

                        if (overwriteChoice == 0)
                        {
                            theSession.ListingWindow.WriteLine("✘ Пропущен (файл существует): " + outFile);
                            continue;
                        }

                        if (overwriteChoice == -2)
                        {
                            DialogResult res2 = System.Windows.Forms.MessageBox.Show(
                                "Файл:\n" + outFile + "\n\nуже существует. Перезаписать?",
                                "Перезапись файла",
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Question);

                            if (res2 != DialogResult.Yes)
                            {
                                theSession.ListingWindow.WriteLine("✘ Пропущен (файл существует): " + outFile);
                                continue;
                            }
                        }
                    }

                    try
                    {
                        CAMObject[] toPost = new CAMObject[] { childGroup };

                        // Запуск постпроцессинга. Если в вашей версии API сигнатура отличается — поправьте соответствующие enum'ы
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

            theUI.NXMessageBox.Show("Готово", NXMessageBox.DialogType.Information, "Постпроцессинг завершён. Файлы в папке:\n" + outputDir);
        }
        catch (Exception ex)
        {
            UI.GetUI().NXMessageBox.Show("Ошибка", NXMessageBox.DialogType.Error, ex.ToString());
        }
    }

    /// <summary>
    /// Находит ближайшую NCGroup, содержащую данную операцию (рекурсивно).
    /// </summary>
    private static NCGroup FindOwningGroupForOperation(NCGroup group, NXOpen.CAM.Operation targetOp)
    {
        if (group == null || targetOp == null) return null;

        CAMObject[] members = null;
        try { members = group.GetMembers(); } catch { members = null; }

        if (members == null || members.Length == 0) return null;

        // Сначала прямые дети (операции)
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

        // Затем рекурсивно в подгруппы
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

    private static string SafeName(CAMObject obj)
    {
        if (obj == null) return "<null>";
        try
        {
            return string.IsNullOrEmpty(obj.Name) ? "<unnamed>" : obj.Name;
        }
        catch
        {
            return "<no-name>";
        }
    }
}
