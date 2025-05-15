import NXOpen
import NXOpen.CAM
import os

# Константы
SETUP_PREFIXES = [
    "SETUP-", "SETUP_",
    "УСТАНОВ-", "УСТАНОВ_",
    "SETTING-", "SETTING_",
    "SET-", "SET_",
    "УСТ_", "УСТ-"
]
NX_ROOT_DEFAULT = r"C:\Program Files\Siemens\NX2306"
POSTPROCESSOR_CACHE = None

def is_setup_group(group_name):
    """Определяет, является ли группа SETUP-группой"""
    return any(group_name.startswith(prefix) for prefix in SETUP_PREFIXES)

def print_setup_structure(group, listing_window, processed_setups):
    """Выводит структуру SETUP-группы с операциями (только первый уровень)"""
    if group.Name in processed_setups:
        return
    processed_setups.add(group.Name)

    listing_window.WriteLine(f"- {group.Name}")

    try:
        for obj in group.GetMembers():
            if isinstance(obj, NXOpen.CAM.Operation):
                # Выводим только операции первого уровня
                listing_window.WriteLine(f"    - {obj.Name}")
            elif isinstance(obj, NXOpen.CAM.NCGroup):
                # Для вложенных групп выводим только их название (без содержимого)
                if not obj.Name.startswith("WORKPIECE-SET-"):
                    listing_window.WriteLine(f"    - {obj.Name}")
    except Exception as e:
        listing_window.WriteLine(f"    [Ошибка: {str(e)}]")

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

        processed_setups = set()  # Для отслеживания уже обработанных SETUP-групп
        
        # Обрабатываем только SETUP-группы верхнего уровня
        for group in cam_setup.CAMGroupCollection:
            if is_setup_group(group.Name):
                print_setup_structure(group, listing_window, processed_setups)

        listing_window.WriteLine("\nДоступные постпроцессоры:")
        for post in list_postprocessors_from_template():
            listing_window.WriteLine(f"  {post}")

        listing_window.WriteLine("\n=== Конец вывода ===")

    except Exception as e:
        listing_window.WriteLine(f"КРИТИЧЕСКАЯ ОШИБКА: {str(e)}")

if __name__ == "__main__":
    main()