import json

modifying_words = []

right_modifying_words = set()
left_modifying_words = set()
other_modifying_words = []


def get_modifying_words():
    # 读取诊断前后缀文件
    with open(file="modifying_words.json", mode='r', encoding='utf-8') as f:
        content = json.load(f)
    return content


def get_diagnosis_cant_input():
    # 读取无法输入的诊断信息
    with open(file="diagnosis_cant_input.txt", mode='r', encoding='utf-8') as f:
        content = f.read()
    return content


def find_modifying_words(text):
    list_text = text.split()
    print('*' * 20)
    for i in list_text:
        l, r = i.split('>>>')
        if r in l:
            modifying_words.append(l.split(r))
        else:
            print(l, "变成了", r)
            other_modifying_words.append(l)

    print('*' * 20)
    for i in modifying_words:
        if i[0] == '':
            right_modifying_words.add(i[1])
        elif i[1] == '':
            left_modifying_words.add(i[0])
        else:
            left_modifying_words.add(i[0])
            right_modifying_words.add(i[1])
            other_modifying_words.append(i)

    print(f'右侧修饰词{len(right_modifying_words)}个：', right_modifying_words)
    print(f'左侧修饰词{len(left_modifying_words)}个：', left_modifying_words)
    print(f'其他修饰词{len(other_modifying_words)}个：', other_modifying_words)

    return left_modifying_words | right_modifying_words


old_modifying_words = get_modifying_words()
newly_added_modifying_words = find_modifying_words(get_diagnosis_cant_input())
# print(f'原有的诊断修饰词有{len(old_modifying_words)}个：{old_modifying_words}')
print(f'新增的诊断修饰词有{len(newly_added_modifying_words)}个：{list(newly_added_modifying_words)}')
print('*' * 20)
new_modifying_words = old_modifying_words + list(newly_added_modifying_words)

with open(file=f'modifying_words.json', mode='w', encoding='utf-8') as f:
    json.dump(new_modifying_words, f, ensure_ascii=False, indent=2)
print(f'诊断修饰词已更新!增加{len(newly_added_modifying_words)}条')
