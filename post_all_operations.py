import os
import shutil
import NXOpen
import NXOpen.CAM

def get_cam_structure(group, indent=0):
    result = ""
    prefix = "    " * indent + "- "
    result += f"{prefix}{group.Name}\n"
    children = group.GetMembers()
    for child in children:
        if isinstance(child, NXOpen.CAM.NCGroup):
            result += get_cam_structure(child, indent + 1)
    return result

def get_postprocessors_from_template(template_path):
    post_names = []
    try:
        with open(template_path, 'r', encoding='utf-8') as f:
            for line in f:
                line = line.strip()
                if line and not line.startswith('#'):
                    parts = line.split(',')
                    if len(parts) >= 3:
                        post_names.append(parts[0])
    except Exception as e:
        post_names.append(f"[Ошибка при чтении template_post.dat: {str(e)}]")
    return post_names

def create_output_folder_on_desktop(folder_name="Output_Programs"):
    desktop_path = os.path.join(os.path.join(os.environ["USERPROFILE"]), "Desktop")
    output_path = os.path.join(desktop_path, folder_name)
    if os.path.exists(output_path):
        shutil.rmtree(output_path)
    os.makedirs(output_path)
    return output_path

def save_log_file(output_path, content):
    log_path = os.path.join(output_path, "log.txt")
    with open(log_path, "w", encoding="utf-8") as f:
        f.write(content)

def postprocess_setup1(work_part, cam_setup, output_path, listing_window):
    try:
        program_group = cam_setup.CAMGroupCollection.FindObject("SETUP-1")
        if not program_group:
            listing_window.WriteLine("CAM группа SETUP-1 не найдена.")
            return

        builder = NXOpen.CAM.Postprocess.CreatePostprocessBuilder(work_part)
        builder.SetProgramGroup(program_group)
        builder.Postprocessor = "HAAS-VF2"
        builder.OutputFile = os.path.join(output_path, "SETUP-1.nc")
        builder.Commit()
        builder.Destroy()
        listing_window.WriteLine("Постпроцессинг SETUP-1 завершен.")
    except Exception as e:
        listing_window.WriteLine(f"[Ошибка постпроцессинга SETUP-1: {str(e)}]")

def main():
    the_session = NXOpen.Session.GetSession()
    work_part = the_session.Parts.Work
    cam_setup = work_part.CAMSetup
    listing_window = the_session.ListingWindow
    listing_window.Open()

    output_folder = create_output_folder_on_desktop()

    result = f"Имя детали: {work_part.Leaf}\n"
    try:
        root_group = cam_setup.CAMGroupCollection.RootGroup
        result += "Структура CAM-групп:\n"
        result += get_cam_structure(root_group)
    except Exception as e:
        result += f"[Ошибка при получении структуры CAM-групп: {str(e)}]\n"

    try:
        template_path = r"C:\\Program Files\\Siemens\\NX2406\\MACH\\resource\\postprocessor\\template_post.dat"
        post_list = get_postprocessors_from_template(template_path)
        result += "\nДоступные постпроцессоры:\n"
        for post in post_list:
            result += f"  {post}\n"
    except Exception as e:
        result += f"[Ошибка при получении постпроцессоров: {str(e)}]\n"

    save_log_file(output_folder, result)
    listing_window.WriteLine(result)

    postprocess_setup1(work_part, cam_setup, output_folder, listing_window)

if __name__ == "__main__":
    main()
