import NXOpen
import NXOpen.CAM
import os

# Константы
NX_ROOT_DEFAULT = r"C:\Program Files\Siemens\NX2306"
POSTPROCESSOR_CACHE = None
SETUP_PREFIXES = ["SETUP-", "SETUP_", "УСТАНОВ-", "УСТАНОВ_"]
EXCLUDE_SUBGROUPS = ["WORKPIECE-SET-", "ROTARY_GEOM", "WORKPIECE_"]

def create_output_folder():
    """Создает папку Program output на рабочем столе, если ее нет"""
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    output_folder = os.path.join(desktop_path, "Program output")
    
    if not os.path.exists(output_folder):
        try:
            os.makedirs(output_folder)
        except Exception as e:
            pass

def list_postprocessors_from_template():
    """Возвращает список доступных постпроцессоров"""
    global POSTPROCESSOR_CACHE
    if POSTPROCESSOR_CACHE is not None:
        return POSTPROCESSOR_CACHE

    posts = []
    nx_root = os.environ.get("UGII_ROOT_DIR", NX_ROOT_DEFAULT)
    template_path = os.path.join(nx_root, "MACH", "resource", "postprocessor", "template_post.dat")

    if not os.path.exists(template_path):
        return ["[Файл template_post.dat не найден]"]

    try:
        with open(template_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith("#"):
                    parts = line.split(",")
                    if len(parts) >= 2:
                        post_name = parts[0].strip()
                        tcl_file = os.path.basename(parts[1].strip())
                        posts.append(f"{post_name} -> {tcl_file}")
        POSTPROCESSOR_CACHE = posts
    except Exception as e:
        POSTPROCESSOR_CACHE = [f"[Ошибка чтения: {str(e)}"]
    return POSTPROCESSOR_CACHE

def is_setup_group(group_name):
    """Проверяет, является ли группа SETUP-группой"""
    group_name_upper = group_name.upper()
    return any(group_name_upper.startswith(prefix.upper()) for prefix in SETUP_PREFIXES)

def should_exclude_group(group_name):
    """Проверяет, нужно ли исключить группу из вывода"""
    group_name_upper = group_name.upper()
    return any(group_name_upper.startswith(prefix.upper()) for prefix in EXCLUDE_SUBGROUPS)

def print_operation(op, listing_window, level):
    """Выводит информацию об операции"""
    indent = "    " * level
    op_type = getattr(op, "TemplateOperationType", "UNKNOWN")
    listing_window.WriteLine(f"{indent}- {op.Name} ({op_type})")

def traverse_setup_group(group, listing_window, level=0, is_main_setup=True):
    """Рекурсивный обход SETUP-группы с фильтрацией"""
    indent = "    " * level
    
    # Выводим только основные SETUP-группы и их непосредственные подгруппы
    if is_main_setup or not should_exclude_group(group.Name):
        listing_window.WriteLine(f"{indent}+ {group.Name} (NCGroup)")
    
    try:
        for member in group.GetMembers():
            if isinstance(member, NXOpen.CAM.Operation):
                if is_main_setup or not should_exclude_group(group.Name):
                    print_operation(member, listing_window, level + 1)
            elif isinstance(member, NXOpen.CAM.NCGroup):
                # Пропускаем группы с геометрией и другими служебными элементами
                if not should_exclude_group(member.Name):
                    traverse_setup_group(member, listing_window, level + 1, False)
    except Exception as e:
        listing_window.WriteLine(f"{indent}  [Ошибка: {str(e)}]")

def main():
    """Основная функция"""
    create_output_folder()
    
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

        listing_window.WriteLine(f"\nДеталь: {work_part.Name}")
        listing_window.WriteLine("="*50)
        listing_window.WriteLine("Структура CAM (основные SETUP группы и операции):")
        
        # Собираем и сортируем SETUP-группы верхнего уровня
        setup_groups = []
        for obj in cam_setup.CAMGroupCollection:
            if isinstance(obj, NXOpen.CAM.NCGroup) and is_setup_group(obj.Name):
                setup_groups.append(obj)
        
        # Сортируем группы по имени для последовательного вывода
        setup_groups.sort(key=lambda x: x.Name)
        
        # Выводим только основные SETUP-группы и их операции
        processed_groups = set()
        for group in setup_groups:
            if group.Name not in processed_groups:
                traverse_setup_group(group, listing_window)
                processed_groups.add(group.Name)
        
        # Вывод постпроцессоров
        listing_window.WriteLine("\n" + "="*50)
        listing_window.WriteLine("Доступные постпроцессоры:")
        for post in list_postprocessors_from_template():
            listing_window.WriteLine(f"  {post}")
        
        listing_window.WriteLine("\n" + "="*50)
        listing_window.WriteLine("Анализ завершен")

    except Exception as e:
        listing_window.WriteLine(f"\nКРИТИЧЕСКАЯ ОШИБКА: {str(e)}")

if __name__ == "__main__":
    main()