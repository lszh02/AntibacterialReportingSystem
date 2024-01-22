import os
import json


def update_modifying_words():
    # 读取原有修饰词
    with open(file=rf"{os.path.dirname(__file__)}\modifying_words.json", mode='r', encoding='utf-8') as f:
        old_modifying_words = json.load(f)
    #  获取新增修饰词
    newly_added_modifying_words = find_modifying_words()

    # 无更新时，返回原有修饰词
    if len(newly_added_modifying_words) == 0:
        return old_modifying_words

    # 更新修饰词
    new_modifying_words = old_modifying_words + list(newly_added_modifying_words)
    # 重新写入json
    with open(file=rf"{os.path.dirname(__file__)}\modifying_words.json", mode='w', encoding='utf-8') as f:
        json.dump(new_modifying_words, f, ensure_ascii=False, indent=2)
    return new_modifying_words


def find_modifying_words():
    modifying_words_set = set()
    # 读取无法输入的诊断文本
    with open(file=rf"{os.path.dirname(__file__)}\diagnosis_cant_input.txt", mode='r', encoding='utf-8') as f:
        diagnosis_cant_input_text = f.read()
        # 追加备份
        with open(file=rf"{os.path.dirname(__file__)}\diagnosis_cant_input_backups.txt", mode='a',
                  encoding='utf-8') as f:
            f.write(diagnosis_cant_input_text)
        # 提取修饰词
        diagnosis_cant_input_list = diagnosis_cant_input_text.split('\n')
        for i in diagnosis_cant_input_list[0:-1]:  # 以\n分割后，最后一个是空字符串，舍弃
            original_text, the_remaining_text = i.split('>>>')
            if the_remaining_text in original_text:
                possible_modifiers = original_text.split(the_remaining_text)
                if possible_modifiers[0] == '':
                    modifying_words_set.add(possible_modifiers[1])
                elif possible_modifiers[1] == '':
                    modifying_words_set.add(possible_modifiers[0])
            else:
                # 将非修饰词写入另一个文件
                with open(file=rf"{os.path.dirname(__file__)}\diagnosis_cant_input_others.txt", mode='a',
                          encoding='utf-8') as f:
                    f.write(f"{original_text}>>>{the_remaining_text}\n")
    # 清空内容
    with open(file=rf"{os.path.dirname(__file__)}\diagnosis_cant_input.txt", mode='w', encoding='utf-8') as f:
        f.write('')

    return modifying_words_set
