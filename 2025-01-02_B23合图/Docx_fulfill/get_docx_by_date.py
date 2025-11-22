import datetime
import os


def get_file_by_date_type(
    file_key_word: str, file_type: str, date_str: str, date_split_str: list
):
    """根据关键词 文件类型 指定日期从上层目录里获取对应docx文件的路径

    Args:
        file_key_word (str): file文件名的关键词
        file_type (str): file扩展名的关键词
        date_str (str): 自定的日期
        date_split_str (str): 日期区间

    Returns:
        [str]: 对应的日期的file文件的路径信息
    """

    # 遍历指定文件夹，获取目标文件的列表
    file_with_dir = ""
    file_list = []
    files = os.listdir("./")
    for file in files:
        if file_key_word in file and file_type in file and "$" not in file:
            file_with_dir = os.path.join("./", file)
            print(file_with_dir)
            file_list.append(file_with_dir)
    file_list.sort()

    # 处理日期字符串转换为date类
    date = datetime.date(*map(int, date_str.split(".")))
    date_split = [
        datetime.date(*map(int, date_str.split("-"))) for date_str in date_split_str
    ]

    # 获取目前日期在日期列表中的位置
    date_pos = 0
    while date_pos < len(date_split) and date >= date_split[date_pos]:
        date_pos += 1
    print(date, date_pos)

    return file_list[date_pos]


if __name__ == "__main__":
    date_split_str = ["2022-01-28", "2023-08-01"]
    B25B26 = get_file_by_date_type("B25B26", "docx", "2021.11.01", date_split_str)
    print(B25B26)
