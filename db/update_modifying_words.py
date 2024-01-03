import json


def update_modifying_words():
    # 读取原有修饰词
    with open(file="modifying_words.json", mode='r', encoding='utf-8') as f:
        old_modifying_words = json.load(f)

    newly_added_modifying_words = find_modifying_words()
    # 连接原有修饰词和新增修饰词
    new_modifying_words = old_modifying_words + list(newly_added_modifying_words)

    with open(file=f'modifying_words.json', mode='w', encoding='utf-8') as f:
        json.dump(new_modifying_words, f, ensure_ascii=False, indent=2)
    print(f'诊断修饰词已更新!增加{len(newly_added_modifying_words)}条')
    return new_modifying_words


def find_modifying_words():
    modifying_words_set = set()
    # 读取无法输入的诊断文本
    with open(file="diagnosis_cant_input.txt", mode='r', encoding='utf-8') as f:
        diagnosis_cant_input_text = f.read()

    diagnosis_cant_input_list = diagnosis_cant_input_text.split()
    for i in diagnosis_cant_input_list:
        original_text, the_remaining_text = i.split('>>>')
        if the_remaining_text in original_text:
            possible_modifiers = original_text.split(the_remaining_text)
            if possible_modifiers[0] == '':
                modifying_words_set.add(possible_modifiers[1])
            elif possible_modifiers[1] == '':
                modifying_words_set.add(possible_modifiers[0])
        else:
            print(original_text, "变成了", the_remaining_text)
    return modifying_words_set

