import pynbs
import xlwings as xw

note_list: dict[int, list[pynbs.Note]] = {}  # 该字典是整首曲子的音符表，第一层是列，第二层是列里面的音符
tick_length: int = 0
fail_ticks: list[int] = []  # 超轨道的Tick


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

    linec = len(row) + 1  # 加一是因为还有执行编码，实际占用的tick为8倍linec
    keyv = [note.key for note in row]  # 本列的纯音符音高列表
    print(f"\033[1;32;40m[SESSION]\033[0m ({tick}/{tick_length})本列共需用{linec}个编码，编码后占用{linec*8}gt。")
    for layer in range(max_layer * 8):
        if layer >= max_layer:  # 最多max_layer层编码器，不能多了
            print(f"\033[1;31;40m[WARNING]\033[0m Tick为{tick}的编码超出了{max_layer}层，正在使用第{layer}层。")
            if tick not in fail_ticks:
                fail_ticks.append(tick)
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
                    getRange(excel, tick + note_delay + 1, layer + 1).value = "#"
                    getRange(excel, tick + note_delay + 1, layer + 1).color = (64, 255, 64)
            for key in reversed(keyv):
                tick -= 8
                try:  # 如果列不存在，则创建新列
                    note_list[tick]
                except KeyError:
                    note_list[tick] = []
                # 填充声明的音符编码
                note_list[tick].append(pynbs.Note(tick=tick, layer=layer, instrument=0, key=key))
                getRange(excel, tick + 1, layer + 1).value = f"!{key - 24}"
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
    wb.save(out_file)
    print(f"\033[1;32;40m[SUCCESS]\033[0m 编码文件已成功写入{out_file}。")
    print(f"\033[1;33;40m[MESSAGE]\033[0m 最大使用层数为{wb.sheets['Process'].used_range.last_cell.row}，超出24轨的Tick如下：{fail_ticks}")
    wb.close()
    app.quit()


if __name__ == "__main__":
    file_name = input("请输入处理好的NBS文件名（不含扩展名）：")
    max_layer = int(input("请输入预期的最大编码器层数："))
    process(file_name, file_name + "-after.xlsx", max_layer)
