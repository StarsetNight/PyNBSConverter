# coding: utf-8
import pynbs
import xlwings as xw
import json

from write_settings import *

note_list: dict[int, list[pynbs.Note]] = {}  # 该字典是整首曲子的编码表，第一层是列，第二层是列里面的编码
latest_notes: dict[int, list[int]] = {}  # 该字典是最近4个tick的音符
tick_length: int = 0
out_ticks: list[int] = []  # 超轨道的Tick
fast_ticks: list[int] = []  # 小于4gt响应的Tick


def writeCommand(*items) -> str:
    """
    将要写入的物品制作为give指令。
    :param items: 潜影盒内的物品列表，从左上至右下，横向排列。
    :return: /give指令
    """
    nbt_data: dict[str, dict[str, list]] = {
        "BlockEntityTag":
            {
                "Items": []
            }
    }
    for slot, item in enumerate(items):
        nbt_data["BlockEntityTag"]["Items"].append(
            {"Slot": slot, "id": item, "Count": 1, "tag": {"display": {"Name": note_name_table[item]}}}
            if item in list(note_name_table) else  # 如果item存在于命名表中，则命名，否则默认
            {"Slot": slot, "id": item, "Count": 1}
        )
    return f"/give @s shulker_box{json.dumps(nbt_data)}"


def getRange(excel: xw.Book, x: int, y: int) -> xw.Range:
    return excel.sheets['Process'].range(f"{intToLetter(x)}{y}")


def intToLetter(number: int) -> str:
    # 将输出字符串初始化为空
    result = ''

    while number > 0:
        # 查找下一个字母的索引并连接该字母
        # 到解决方案

        # 这里索引0对应`A`，25对应`Z`
        index = (number - 1) % 26
        result += chr(index + ord('A'))
        number = (number - 1) // 26

    return result[::-1]


def detectLayer(length: int, layer: int) -> bool:
    """
    检测是否为可用层
    :param length: 音符长度
    :param layer: 层数
    :return: 是否够空间
    """
    ticks = [tick for tick in note_list][-length-1:]  # 预截取
    for tick in ticks:
        note_row = note_list[tick]  # 代表此tick对应的列
        for note in note_row:
            if note.layer == layer:  # 如果layer对应上，则代表音符冲突，这层不能用
                return False
    return True  # 如果一直没有对应，那就能用


def parse(excel: xw.Book, tick: int, row: list[pynbs.Note], max_layer: int):
    """
    处理列并转译到新文件
    :param excel: Excel对象
    :param tick: 旧文件某个特定音符的tick，要与新文件的对应执行编码tick对齐
    :param row: 列
    :param max_layer: 预期的最大使用编码器层数
    :return:
    """
    global tick_length

    linec: int = len(row) + 1  # 加一是因为还有执行编码，实际占用的tick为8倍linec
    keyv: list[int] = [note.key for note in row]  # 本列的纯音符音高列表
    keyv: list[int] = sorted(set(keyv), key=keyv.index)
    print(f"\033[1;32;40m[SESSION]\033[0m ({tick}/{tick_length})本列共需用{linec}个编码，编码后占用{linec*8}gt。")

    latest_notes[tick] = keyv
    latest_tick: int = list(latest_notes)[-1]
    latest_keys: list[int] = []
    for note in list(latest_notes):  # tick筛选
        if latest_tick - note > 4:
            latest_notes.pop(note)  # 如果tick相差超过4，则剔除
    for note in latest_notes.values():  # 音高检测
        latest_keys.extend([key for key in note])
    if len(set(latest_keys)) < len(latest_keys):  # 说明4个gt内有重复的音高
        print(f"\033[1;31;40m[WARNING]\033[0m Tick为{latest_tick}的编码播放了间隔小于4gt的同一音符，音符盒将无法响应此播放。")
        fast_ticks.append(latest_tick)

    for layer in range(max_layer * 8):
        if layer >= max_layer:  # 最多max_layer层编码器，不能多了
            print(f"\033[1;31;40m[WARNING]\033[0m Tick为{tick}的编码超出了{max_layer}层，正在使用第{layer}层。")
            if tick not in out_ticks:
                out_ticks.append(tick)
        if detectLayer(linec, layer):  # 如果这层放得下编码
            delay = tick % 8  # 由于每个编码必须间隔8gt，所以应当有延迟
            tick -= delay  # 先减去延迟
            try:  # 如果列不存在，则创建新列
                note_list[tick]
            except KeyError:
                note_list[tick] = []
            if not delay:  # 填充黄色的无延迟执行编码
                note_list[tick].append(pynbs.Note(tick=tick, layer=layer, instrument=3, key=23))
                getRange(excel, tick + 1, layer + 1).value = "@"
                getRange(excel, tick + 1, layer + 1).color = (255, 255, 64)

            else:  # 填充绿色的有延迟执行编码
                for note_delay in range(delay):  # 逐渐往后填充1个gt的绿色编码
                    note_list[tick].append(pynbs.Note(tick=tick + note_delay, layer=layer, instrument=1, key=23))
                    getRange(excel, tick + note_delay + 1, layer + 1).value = "#" * delay
                    getRange(excel, tick + note_delay + 1, layer + 1).color = (64, 255, 64)
            for key in reversed(keyv):
                tick -= 8
                try:  # 如果列不存在，则创建新列
                    note_list[tick]
                except KeyError:
                    note_list[tick] = []
                # 填充声明的音符编码
                note_list[tick].append(pynbs.Note(tick=tick, layer=layer, instrument=0, key=key))
                getRange(excel, tick + 1, layer + 1).value = f"!{key - 33}"  # 得到的key实质上是音高+33
                getRange(excel, tick + 1, layer + 1).color = (128, 128, 255)
            break


def process(in_file: str, out_file: str, max_layer: int):
    """
    如果计划没问题的话……tick指的是旧文件某个特定音符的tick，然后它也要与新文件的对应执行编码tick对齐；
    chord指的是tick对应的列音符……
    也许我们是在将传统Minecraft红石音乐转换成第四代编码格式……
    :param max_layer: 预期的最大使用编码器层数
    :param in_file: 输入的nbs文件
    :param out_file: 输出的Excel文件
    :return:
    """
    global tick_length

    core_song = pynbs.read(in_file + ".nbs")
    tick_length = core_song.header.song_length
    print(f"\033[1;33;40m[SESSION]\033[0m 开始处理{in_file}.nbs，总时长为{tick_length}tick。")
    app = xw.App(visible=False, add_book=False)
    wb = app.books.add()
    wb.sheets.add("Process")
    for tick, chord in core_song:
        # note_list[tick] = chord
        parse(wb, tick, chord, max_layer)
    wb.save(out_file + "-after.xlsx")
    print(f"\033[1;32;40m[SUCCESS]\033[0m 编码文件已成功写入{out_file}-after.xlsx。")
    print(f"\033[1;33;40m[MESSAGE]\033[0m 最大使用层数为{wb.sheets['Process'].used_range.last_cell.row}，"
          f"超出24轨的Tick如下：{out_ticks}，小于4gt的同音符Tick如下：{fast_ticks}。")
    choice: str = input("是否要将潜影盒装填写入至/give命令列表[y/N]：")
    if choice == "y":
        with open(out_file + "-after-command.txt", "w+", encoding="utf-8") as command_file:
            item_count: int = 0  # 用于潜影盒27格计数
            item_list: list[str] = []  # 用于存储潜影盒27格物品
            for y in range(1, wb.sheets['Process'].used_range.last_cell.row + 1):  # y的值在1到最大使用层数之间
                print(f"\033[1;33;40m[SESSION]\033[0m 正在处理第{y}层的潜影盒装填……")
                for x in range(1, tick_length + 1, 8):  # x的值在1到歌曲长度（tick）之间
                    if getRange(wb, x, y).value:
                        match getRange(wb, x, y).value[0]:
                            case "!":  # 音符声明
                                item_list.append(
                                    list(note_name_table)  # 将note_name_table转换为键列表
                                    [int(getRange(wb, x, y).value[1:])]  # 找到对应键
                                )
                            case "@":  # 无延迟执行
                                item_list.append("purple_carpet")
                            case "#":
                                item_list.append(executer_name_table[len(getRange(wb, x, y).value) - 1])
                    else:
                        item_list.append("oak_fence")
                    item_count += 1
                    if item_count == 27:
                        command_file.write(writeCommand(item_list) + "\n\n")
                        item_list.clear()
                        item_count = 0
    print(f"\033[1;32;40m[SUCCESS]\033[0m 命令文件已成功写入{out_file}-after-command.txt。")
    wb.close()
    app.quit()


if __name__ == "__main__":
    file_name = input("请输入处理好的NBS文件名（不含扩展名）：")
    max_layer = int(input("请输入预期的最大编码器层数："))
    process(file_name, file_name, max_layer)
