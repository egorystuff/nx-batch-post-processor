import NXOpen
import NXOpen.CAM
import os

def is_setup_group(group_name):
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
                if level == 0 or is_setup_group(group.Name) or level > 0:
                    print_group_structure(obj, level + 1, listing_window, processed_groups)
    except Exception as e:
        listing_window.WriteLine(f"{indent}    [Ошибка: {str(e)}]")

def list_postprocessors_from_template():
    posts = []
    # Пример пути, можно подкорректировать под твою версию NX
    nx_root = os.environ.get("UGII_ROOT_DIR", r"C:\Program Files\Siemens\NX2406")
    template_path = os.path.join(nx_root, "MACH", "resource", "postprocessor", "template_post.dat")

    if not os.path.exists(template_path):
        return ["[Файл template_post.dat не найден]"]

    try:
        with open(template_path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if not line or line.startswith("#"):
                    continue
                parts = line.split(",")
                if len(parts) >= 2:
                    post_name = parts[0].strip()
                    tcl_path = parts[1].strip()
                    tcl_file = os.path.basename(tcl_path.replace("${UGII_CAM_POST_DIR}", ""))
                    posts.append(f"{post_name} -> {tcl_file}")
    except Exception as e:
        posts.append(f"[Ошибка чтения template_post.dat: {str(e)}]")

    return posts

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
        for group in cam_setup.CAMGroupCollection:
            if is_setup_group(group.Name):
                print_group_structure(group, 0, listing_window, processed_groups)

        listing_window.WriteLine("\nДоступные постпроцессоры:")
        posts = list_postprocessors_from_template()
        for post in posts:
            listing_window.WriteLine(f"  {post}")

        listing_window.WriteLine("\n=== Конец вывода ===")

    except Exception as e:
        listing_window.WriteLine(f"КРИТИЧЕСКАЯ ОШИБКА: {str(e)}")

if __name__ == "__main__":
    main()
