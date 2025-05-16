import NXOpen
import NXOpen.CAM
import os

def postprocess_operation(operation_name, postprocessor_name, output_folder):
    the_session = NXOpen.Session.GetSession()
    listing_window = the_session.ListingWindow
    listing_window.Open()
    
    try:
        listing_window.WriteLine("=== Начало автоматического постпроцессинга ===")
        
        # Получаем текущую деталь и CAMSetup
        work_part = the_session.Parts.Work
        if not work_part:
            listing_window.WriteLine("ОШИБКА: Нет открытой рабочей детали!")
            return False

        cam_setup = work_part.CAMSetup
        if not cam_setup:
            listing_window.WriteLine("ОШИБКА: CAM не активирован!")
            return False

        # Находим операцию по имени
        operation = None
        for group in cam_setup.CAMGroupCollection:
            for member in group.GetMembers():
                if isinstance(member, NXOpen.CAM.Operation) and member.Name == operation_name:
                    operation = member
                    break

        if not operation:
            listing_window.WriteLine(f"ОШИБКА: Операция {operation_name} не найдена!")
            return False

        # Создаем билдер постпроцессинга
        post_builder = cam_setup.CreatePostprocessBuilder()
        post_builder.SetOperation(operation)
        post_builder.Postprocessor = postprocessor_name
        
        # Формируем путь для сохранения
        output_filename = f"{operation_name}.nc"
        output_path = os.path.join(output_folder, output_filename)
        post_builder.OutputFile = output_path
        
        # Настройки из журнала
        post_builder.ListingOutput = True  # Вывод листинга
        post_builder.OutputBallCenter = False  # Не выводить центр шара

        # Выполняем постпроцессинг
        listing_window.WriteLine(f"Начинаю постпроцессинг операции {operation_name}...")
        listing_window.WriteLine(f"Постпроцессор: {postprocessor_name}")
        listing_window.WriteLine(f"Выходной файл: {output_path}")
        
        result = post_builder.Commit()
        post_builder.Destroy()
        
        if result == 0:  # 0 = успешно
            listing_window.WriteLine("Постпроцессинг завершен успешно!")
            return True
        else:
            listing_window.WriteLine(f"ОШИБКА: Код возврата {result}")
            return False

    except Exception as e:
        listing_window.WriteLine(f"КРИТИЧЕСКАЯ ОШИБКА: {str(e)}")
        return False
    finally:
        listing_window.WriteLine("=== Завершение работы скрипта ===")

# Параметры для запуска
if __name__ == "__main__":
    operation_to_process = "O454010104_СВЕРЛЕНИЕ_ПОД_ВОДУ_ГЛУБИНА-70ММ_И_ГЛУХОЕ_ОТВ"
    selected_postprocessor = "DMU-60T"  # Из вашего журнала
    output_dir = os.path.join(os.environ["USERPROFILE"], "Desktop", "Program output")
    
    # Создаем папку для вывода
    os.makedirs(output_dir, exist_ok=True)
    
    # Запускаем процесс
    success = postprocess_operation(operation_to_process, selected_postprocessor, output_dir)
    
    if success:
        print(f"Файл успешно сохранен в: {output_dir}")
    else:
        print("Произошла ошибка при постпроцессинге. Проверьте окно листинга NX.")