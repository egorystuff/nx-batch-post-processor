// CamProgram_BatchPostprocess_3Posts_SelectFolder_Modern.cs
//
// Для каждой подпрограммы выбранной группы выполняет постпроцессинг тремя постпроцессорами.
// Перед запуском выводится современное диалоговое окно выбора папки для сохранения.
// Файлы сохраняются с расширениями .i, .min, .txt
// Опция перезаписи: Да / Нет / Спрашивать каждый раз
//
// ✅ Добавлен 4-й режим: "Авто"
// ✅ В режиме "Авто" тип обработки определяется по имени (_4X_, _5X_)
// ✅ Если флагов нет — используется 3-х осевая обработка (по умолчанию)
// ✅ Добавлен вывод в лог о выбранном типе обработки для каждой программы

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

    // --- Разные группы постпроцессоров ---
    private static readonly PostConfig[] Posts3Axis = new PostConfig[]
    {
        new PostConfig("DMU-60T", ".i"),
        new PostConfig("OKUMA_MB-46VAE_3X", ".min"),
        new PostConfig("HAAS-VF2", ".txt")
    };

    private static readonly PostConfig[] Posts4Axis = new PostConfig[]
    {
        new PostConfig("OKUMA_MB-46VAE_4X", ".min"),
    };

    private static readonly PostConfig[] Posts5Axis = new PostConfig[]
    {
        new PostConfig("DMU-5axis", ".i"),
    };

    // Текущий набор постпроцессоров
    private static PostConfig[] PostConfigs;

    // Опция перезаписи
    private static int overwriteChoice = -1;

    public static void Main(string[] args)
    {
        try
        {
            theSession = Session.GetSession();
            theUI = UI.GetUI();
            workPart = theSession.Parts.Work;

            // === Выбор типа обработки ===
            string[] options = {
                "3-х осевая обработка",
                "4-х осевая обработка",
                "5 осевая обработка",
                "Авто"
            };
            string choice = ShowProcessingTypeDialog(options);
            if (choice == null) return;

            switch (choice)
            {
                case "3-х осевая обработка": PostConfigs = Posts3Axis; break;
                case "4-х осевая обработка": PostConfigs = Posts4Axis; break;
                case "5 осевая обработка": PostConfigs = Posts5Axis; break;
                case "Авто": PostConfigs = null; break;
                default: return;
            }

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
            theSession.ListingWindow.WriteLine("=== Batch postprocess (" + choice + ") ===");
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

                // --- Определяем тип обработки в режиме "Авто" ---
                PostConfig[] currentPosts = PostConfigs;
                string autoDetectedType = "";

                if (choice == "Авто")
                {
                    string upperName = fullName.ToUpper();

                    if (upperName.Contains("_5X_"))
                    {
                        currentPosts = Posts5Axis;
                        autoDetectedType = "5-осевая обработка";
                    }
                    else if (upperName.Contains("_4X_"))
                    {
                        currentPosts = Posts4Axis;
                        autoDetectedType = "4-осевая обработка";
                    }
                    else
                    {
                        currentPosts = Posts3Axis;
                        autoDetectedType = "3-осевая обработка";
                    }

                    // Добавляем строку в лог
                    theSession.ListingWindow.WriteLine(fullName + " → " + autoDetectedType);
                }
                else
                {
                    theSession.ListingWindow.WriteLine(fullName);
                }

                foreach (PostConfig cfg in currentPosts)
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
                        theSession.ListingWindow.WriteLine("   ✔ " + cfg.PostName + " -> " + SysIO.Path.GetFileName(outFile));
                    }
                    catch (Exception exPost)
                    {
                        theSession.ListingWindow.WriteLine("   ✘ Ошибка (" + cfg.PostName + "): " + exPost.Message);
                    }
                }

                theSession.ListingWindow.WriteLine("");
            }

            theUI.NXMessageBox.Show("Готово", NXMessageBox.DialogType.Information,
                "Постпроцессинг (" + choice + ") завершён. Файлы сохранены в:\n" + outputDir);
        }
        catch (Exception ex)
        {
            UI.GetUI().NXMessageBox.Show("Ошибка", NXMessageBox.DialogType.Error, ex.ToString());
        }
    }

    // === Диалог выбора типа обработки ===
    private static string ShowProcessingTypeDialog(string[] options)
    {
        using (Form form = new Form())
        {
            form.Text = "Выбор типа обработки";
            form.Width = 500;
            form.Height = 300;
            form.StartPosition = FormStartPosition.CenterScreen;
            form.FormBorderStyle = FormBorderStyle.FixedDialog;
            form.MaximizeBox = false;
            form.MinimizeBox = false;

            Label label = new Label() { Left = 20, Top = 20, Width = 450, Text = "Выберите тип обработки:" };
            ComboBox combo = new ComboBox() { Left = 20, Top = 50, Width = 440, DropDownStyle = ComboBoxStyle.DropDownList };
            combo.Items.AddRange(options);
            combo.SelectedIndex = 0;

            Button ok = new Button() { Text = "OK", Left = 160, Width = 70, Top = 100, DialogResult = DialogResult.OK };
            Button cancel = new Button() { Text = "Отмена", Left = 250, Width = 80, Top = 100, DialogResult = DialogResult.Cancel };

            form.Controls.Add(label);
            form.Controls.Add(combo);
            form.Controls.Add(ok);
            form.Controls.Add(cancel);

            form.AcceptButton = ok;
            form.CancelButton = cancel;

            DialogResult result = form.ShowDialog();
            if (result == DialogResult.OK)
            {
                return combo.SelectedItem.ToString();
            }
        }
        return null;
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
