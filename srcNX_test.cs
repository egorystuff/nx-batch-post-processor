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

    public static void Main(string[] args)
    {
        try
        {
            theSession = Session.GetSession();
            theUI = UI.GetUI();
            workPart = theSession.Parts.Work;

            // Проверка на наличие активной детали и CAM Setup
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

            // Получаем выбранный объект
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

            // Открываем окно вывода и начинаем логирование
            theSession.ListingWindow.Open();
            theSession.ListingWindow.WriteLine("=== Структура выбранной CAM группы ===");
            theSession.ListingWindow.WriteLine("Группа: " + SafeName(startGroup));
            theSession.ListingWindow.WriteLine("");

            // Рекурсивный обход группы и вывод всех дочерних объектов
            PrintGroupStructure(startGroup, 0);
        }
        catch (Exception ex)
        {
            theUI.NXMessageBox.Show("Ошибка", NXMessageBox.DialogType.Error, ex.ToString());
        }
    }

    // Рекурсивный метод для обхода и вывода всех объектов в группе
    private static void PrintGroupStructure(NCGroup group, int indentLevel)
    {
        if (group == null) return;

        // Отступы для правильного отображения вложенности
        string indent = new string(' ', indentLevel * 2);

        // Логируем имя группы
        theSession.ListingWindow.WriteLine(indent + "Группа: " + SafeName(group));

        // Получаем и обрабатываем все члены группы
        CAMObject[] members = group.GetMembers();
        if (members != null)
        {
            foreach (CAMObject member in members)
            {
                // Если объект - это еще одна группа, рекурсивно обрабатываем её
                NCGroup childGroup = member as NCGroup;
                if (childGroup != null)
                {
                    PrintGroupStructure(childGroup, indentLevel + 1);
                }
                else
                {
                    // Если это операция, выводим её
                    NXOpen.CAM.Operation operation = member as NXOpen.CAM.Operation;
                    if (operation != null)
                  {
                      theSession.ListingWindow.WriteLine(indent + "  Операция: " + SafeName(operation));

                      NXOpen.CAM.Tool tool = operation.ParentMachineTool as NXOpen.CAM.Tool;

                      if (tool != null)
                      {
                          theSession.ListingWindow.WriteLine(indent + "    Инструмент: " + SafeName(tool));

                          // DumpAllParams(operation, indent + "      ");
                          // DumpTool(tool, indent + "      ");
                      }
                      else
                      {
                          theSession.ListingWindow.WriteLine(indent + "    Инструмент: <нет>");
                      }
                  }
                }
            }
        }
    }

    // Метод для безопасного получения имени объекта
    private static string SafeName(CAMObject obj)
    {
        if (obj == null) return "<null>";
        try { return string.IsNullOrEmpty(obj.Name) ? "<unnamed>" : obj.Name; }
        catch { return "<no-name>"; }
    }

    // Метод для поиска группы, содержащей операцию
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




//     private static void DumpTool(NXOpen.CAM.Tool tool, string indent)
// {
//     if (tool == null) return;

//     var type = tool.GetType();
//     var props = type.GetProperties();

//     theSession.ListingWindow.WriteLine(indent + "Тип инструмента: " + type.FullName);

//     foreach (var p in props)
//     {
//         try
//         {
//             object val = p.GetValue(tool, null);
//             string valueStr = val == null ? "null" : val.ToString();

//             theSession.ListingWindow.WriteLine(indent + p.Name + " = " + valueStr);
//         }
//         catch
//         {
//             theSession.ListingWindow.WriteLine(indent + p.Name + " = <ошибка>");
//         }
//     }
// }



// private static void DumpAllParams(NXOpen.CAM.Operation op, string indent)
// {
//     theSession.ListingWindow.WriteLine(indent + "=== PARAM DUMP ===");

//     try
//     {
//         string[] names = op.GetParameterNames();

//         foreach (string n in names)
//         {
//             try
//             {
//                 int type = op.GetParameterType(n);

//                 switch (type)
//                 {
//                     case 1: // int
//                         theSession.ListingWindow.WriteLine(indent + n + " = " + op.GetIntegerValue(n));
//                         break;

//                     case 2: // double
//                         theSession.ListingWindow.WriteLine(indent + n + " = " + op.GetRealValue(n));
//                         break;

//                     case 3: // string
//                         theSession.ListingWindow.WriteLine(indent + n + " = " + op.GetStringValue(n));
//                         break;

//                     default:
//                         theSession.ListingWindow.WriteLine(indent + n + " (type " + type + ")");
//                         break;
//                 }
//             }
//             catch { }
//         }
//     }
//     catch (Exception ex)
//     {
//         theSession.ListingWindow.WriteLine(indent + "Dump error: " + ex.Message);
//     }
// }







}
