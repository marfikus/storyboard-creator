
import pandas as pd
import os
import uuid
import json
import cv2


def extract_frames(filepath, num_frames):
    """
    Извлекает num_frames кадров из filepath видео
    """

    if not os.path.exists(filepath):
        print("Video file is not found!")
        return

    if num_frames <= 0:
        print("Incorrect value of num_frames:", num_frames)
        return

    cap = cv2.VideoCapture(filepath)

    # получаем общее количество кадров в видео
    total_frame_count = cap.get(cv2.CAP_PROP_FRAME_COUNT)

    # считаем отступ в начале и в конце видео (титры, например..)
    offset = 5 # отступ в %
    offset_frames = total_frame_count // 100 * 5

    # получаем количество кадров, с которых будем делать раскадровку
    frame_count = total_frame_count - offset_frames * 2

    if num_frames > frame_count:
        print("num_frames > frame_count")
        return

    # получаем шаг извлечения кадров
    # (-1 для более равномерного распределения)
    frame_step = frame_count // (num_frames - 1)

    # позиция извлекаемого кадра
    frame_pos = offset_frames

    frames = []
    for i in range(num_frames):
        # устанавливаем текущую позицию
        cap.set(cv2.CAP_PROP_POS_FRAMES, frame_pos)

        # извлекаем кадр
        ret, frame = cap.read()
        # cv2.imshow("frame", frame)
        # cv2.waitKey()

        # немного уменьшаем размер
        scaled_frame = cv2.resize(
            frame, 
            None, 
            fx=0.7, 
            fy=0.7, 
            interpolation=cv2.INTER_AREA
        )

        frames.append(scaled_frame)
        frame_pos += frame_step

    cap.release()
    cv2.destroyAllWindows()

    return frames


def create_frame_grid(frames, frames_in_row):
    """
    Создаёт сетку из массива кадров frames
    с количеством кадров в строке = frames_in_row
    """

    if frames_in_row <= 0:
        print("Incorrect value of frames_in_row:", frames_in_row)
        return

    if frames_in_row > len(frames):
        print("frames_in_row > len(frames)")
        return

    # если сетка получается несимметричная
    # (число кадров в строке != числу строк), то:
        # либо просто выходим, 
        # либо строим несимметричную сетку (пока не сделал)
        # либо ищем ближайшее значение frames_in_row,
        # чтобы сетка была симметрична (текущий вариант)
    if len(frames) % frames_in_row != 0:
        print("Unbalanced grid!")
        # return

        # определяем границы для циклов поиска ближайшего значения
        max_frames_in_row = frames_in_row + 10
        min_frames_in_row = 1
        # определяем значения на случай, если в цикле они не изменятся
        found_value_more = max_frames_in_row
        found_value_less = min_frames_in_row
        # бежим вправо от текущего значения
        for value in range(frames_in_row, max_frames_in_row):
            if len(frames) % value == 0:
                found_value_more = value
                break
        # а теперь влево
        for value in range(frames_in_row, min_frames_in_row - 1, -1):
            if len(frames) % value == 0:
                found_value_less = value
                break
        # считаем разницы и выбираем наименьшую
        positive_diff = found_value_more - frames_in_row
        negative_diff = frames_in_row - found_value_less
        if positive_diff <= negative_diff:
            frames_in_row = found_value_more
        else:
            frames_in_row = found_value_less

    frame_rows_count = int(len(frames) / frames_in_row)
    start_row_pos = 0
    end_row_pos = frames_in_row
    frame_rows = []

    # разбиваем массив фреймов на строки
    for row in range(frame_rows_count):
        frame_row = cv2.hconcat(frames[start_row_pos:end_row_pos])
        frame_rows.append(frame_row)
        start_row_pos += frames_in_row
        end_row_pos += frames_in_row

    # объединяем строки в сетку
    frame_grid = cv2.vconcat(frame_rows)
    return frame_grid


def data_processing(input_file_path, output_file_path):
    """
    Парсит эксель
    В цикле перебирает видео:
        Извлекает кадры
        Создаёт раскадровку
        Заполняет словарь данными
    Сохраняет данные в файл json
    """

    if not os.path.exists(input_file_path):
        print("Input file is not found!")
        return

    df = pd.read_excel(input_file_path, sheet_name=0)
    filepath_list = df["filepath"].tolist()

    output_data = []
    for filepath in filepath_list:
        print(filepath)
        frames = extract_frames(filepath, num_frames=16)
        # если извлечь кадры не получилось, то пропускаем этот файл
        # (не заносим в результат)
        if frames is None:
            continue

        path, video_full_name = os.path.split(filepath)
        video_name, ext = os.path.splitext(video_full_name)

        image_path = os.path.join(path, f"{video_name}_preview.jpg") 
        image_path = os.path.normpath(image_path)

        # создание сетки фреймов 4х4
        frame_grid = create_frame_grid(frames, frames_in_row=4)

        # если сетка создалась, то сохраняем, иначе пропускаем итерацию
        if not frame_grid is None:
            cv2.imwrite(image_path, frame_grid)
        else:
            continue

        # ид создаётся пока просто случайный
        file_id = str(uuid.uuid4())

        data_dict = {
            "id": file_id,
            "name": video_name,
            "image": image_path 
        }
        output_data.append(data_dict)

    output_data_json = json.dumps(output_data, indent=4)
    with open(output_file_path, "w") as f:
        json.dump(output_data_json, f)


if __name__ == "__main__":
    input_file = os.path.abspath("input_data.xlsx")
    output_file = os.path.abspath("output_data.json")

    data_processing(input_file, output_file)

    with open(output_file, "r") as f:
        d = json.load(f)
        print(d)
