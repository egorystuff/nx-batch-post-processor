import NXOpen
import NXOpen.CAM

def is_setup_group(group_name):
    # Определяем, является ли группа группой установки
    return (group_name.startswith("SETUP-") or 
            group_name.startswith("SETUP_") or
            group_name.startswith("УСТАНОВ-") or
            group_name.startswith("УСТАНОВ_") or
            group_name.startswith("SETTING-") or
            group_name.startswith("SETTING_") or
            group_name.startswith("SET-") or
            group_name.startswith("SET_") or
            group_name.startswith("УСТ_") or
            group_name.startswith("УСТ-"))

def print_group_structure(group, level, listing_window, processed_groups):
    if group in processed_groups:
        return
    processed_groups.add(group)
    
    indent = "    " * level
    listing_window.WriteLine(f"{indent}- {group.Name}")

    try:
        for obj in group.GetMembers():
            if isinstance(obj, NXOpen.CAM.NCGroup):
                # Для групп установок печатаем все вложенные группы полностью
                if level == 0 or is_setup_group(group.Name) or level > 0:
                    print_group_structure(obj, level + 1, listing_window, processed_groups)
            else:
                # Печатаем операции для всех уровней вложенности
                # listing_window.WriteLine(f"{indent}    • {obj.Name}")
                return
    except Exception as e:
        listing_window.WriteLine(f"{indent}    [Ошибка: {str(e)}]")

def main():
    the_session = NXOpen.Session.GetSession()
    listing_window = the_session.ListingWindow
    listing_window.Open()

    try:
        work_part = the_session.Parts.Work

        if not work_part:
            listing_window.WriteLine("ОШИБКА: Нет открытой рабочей детали!")
            return

        cam_setup = work_part.CAMSetup
        if not cam_setup:
            listing_window.WriteLine("ОШИБКА: CAM не активирован!")
            return

        listing_window.WriteLine(f"Имя детали: {work_part.Name}")
        listing_window.WriteLine("Структура CAM-групп:")

        processed_groups = set()
        
        # Выводим только группы установок и их полную иерархию
        for group in cam_setup.CAMGroupCollection:
            if is_setup_group(group.Name):
                print_group_structure(group, 0, listing_window, processed_groups)

        listing_window.WriteLine("=== Конец вывода ===")

    except Exception as e:
        listing_window.WriteLine(f"КРИТИЧЕСКАЯ ОШИБКА: {str(e)}")

if __name__ == "__main__":
    main()