import time
import win32api
import win32con

voiced_map = {
    # 濁音和半濁音的映射
    'が': 'か゛',
    'ぎ': 'き゛',
    'ぐ': 'く゛',
    'げ': 'け゛',
    'ご': 'こ゛',
    'ざ': 'さ゛',
    'じ': 'し゛',
    'ず': 'す゛',
    'ぜ': 'せ゛',
    'ぞ': 'そ゛',
    'だ': 'た゛',
    'ぢ': 'ち゛',
    'づ': 'つ゛',
    'で': 'て゛',
    'ど': 'と゛',
    'ば': 'は゛',
    'び': 'ひ゛',
    'ぶ': 'ふ゛',
    'べ': 'へ゛',
    'ぼ': 'ほ゛',
    'ぱ': 'は゜',
    'ぴ': 'ひ゜',
    'ぷ': 'ふ゜',
    'ぺ': 'へ゜',
    'ぽ': 'ほ゜',
    ##########
    'ヴ': 'ウ゛',
    'ガ': 'カ゛',
    'ギ': 'キ゛',
    'グ': 'ク゛',
    'ゲ': 'ケ゛',
    'ゴ': 'コ゛',
    'ザ': 'サ゛',
    'ジ': 'シ゛',
    'ズ': 'ス゛',
    'ゼ': 'セ゛',
    'ゾ': 'ソ゛',
    'ダ': 'タ゛',
    'ヂ': 'チ゛',
    'ヅ': 'ツ゛',
    'デ': 'テ゛',
    'ド': 'ト゛',
    'バ': 'ハ゛',
    'ビ': 'ヒ゛',
    'ブ': 'フ゛',
    'ベ': 'ヘ゛',
    'ボ': 'ホ゛',
    'パ': 'ハ゜',
    'ピ': 'ヒ゜',
    'プ': 'フ゜',
    'ペ': 'ヘ゜',
    'ポ': 'ホ゜'
}

keyboard = {
    'English': {
        'Lowercase': [
            ['1', '2', '3', '4', '5', '6', '7', '8', '9', '0', '-', '+'],
            ['q', 'w', 'e', 'r', 't', 'y', 'u', 'i', 'o', 'p', '[', ']'],
            ['a', 's', 'd', 'f', 'g', 'h', 'j', 'k', 'l', ';', '@', '~'],
            ['z', 'x', 'c', 'v', 'b', 'n', 'm', ',', '.', '/', '?', ' ', ' '],
        ],
        'Uppercase': [
            ['！', '@', '#', '$', '%', '^', '&', '*', '(', ')', '_', '+'],
            ['Q', 'W', 'E', 'R', 'T', 'Y', 'U', 'I', 'O', 'P', '{', '}'],
            ['A', 'S', 'D', 'F', 'G', 'H', 'J', 'K', 'L', ':', '\\\\', '"', '`'],
            ['Z', 'X', 'C', 'V', 'B', 'N', 'M', '<', '>', '/', '?', '|'],
        ],
    },
    'Japanese': {
        'Hiragana': [
        ['あ', 'か', 'さ', 'た', 'な', 'は', 'ま', 'や', 'ら', 'わ', 'ぁ', 'ゃ', 'っ'],
        ['い', 'き', 'し', 'ち', 'に', 'ひ', 'み', '　', 'り', '　', 'ぃ', '　', '　'],
        ['う', 'く', 'す', 'つ', 'ぬ', 'ふ', 'む', 'ゆ', 'る', 'を', 'ぅ', 'ゅ', '゛'],
        ['え', 'け', 'せ', 'て', 'ね', 'へ', 'め', '　', 'れ', '　', 'ぇ', '　', '゜'],
        ['お', 'こ', 'そ', 'と', 'の', 'ほ', 'も', 'よ', 'ろ', 'ん', 'ぉ', 'ょ', 'ー'],
        ['・', '。', '、', '「', '」', '『', '』', '〜', ' ', ' ', ' ', ' ', ' '],
    ],
        'Katakana': [
        ['ア', 'カ', 'サ', 'タ', 'ナ', 'ハ', 'マ', 'ヤ', 'ラ', 'ワ', 'ァ', 'ャ', 'ッ'],
        ['イ', 'キ', 'シ', 'チ', 'ニ', 'ヒ', 'ミ', '　', 'リ', '　', 'ィ', '　', '　'],
        ['ウ', 'ク', 'ス', 'ツ', 'ヌ', 'フ', 'ム', 'ユ', 'ル', 'ヲ', 'ゥ', 'ュ', '゛'],
        ['エ', 'ケ', 'セ', 'テ', 'ネ', 'ヘ', 'メ', '　', 'レ', '　', 'ェ', '　', '゜'],
        ['オ', 'コ', 'ソ', 'ト', 'ノ', 'ホ', 'モ', 'ヨ', 'ロ', 'ン', 'ォ', 'ョ', 'ー'],
        ['・', '。', '、', '「', '」', '『', '』', '〜', ' ', ' ', ' ', ' ', ' '],
    ]
    }

}


def find_position(char, keyboard_page):
    for i, row in enumerate(keyboard_page):
        if char in row:
            return (i, row.index(char))
    return None


def input_text(text):
    global current_page  # 使用全域變數
    current_position = (0, 0)  # 初始化當前位置
    current_language_index = 0  # 初始化當前語言索引
    current_case_index = 0  # 初始化當前大小寫或假名類型索引
    for char in text:
        # 先嘗試在當前語言和頁面中找到字符
        current_language = list(keyboard.keys())[current_language_index]
        current_case = list(keyboard[current_language].keys())[current_case_index]
        target_position = find_position(char, keyboard[current_language][current_case])
        if target_position is None:
            # 如果在當前頁面找不到，嘗試翻頁
            current_case_index = (current_case_index + 1) % len(keyboard[current_language].keys())
            target_position = find_position(char, keyboard[current_language][list(keyboard[current_language].keys())[current_case_index]])
            if target_position is None:
                # 如果翻頁後仍然找不到，嘗試換語言
                current_language_index = (current_language_index + 1) % len(keyboard)
                current_case_index = 0
                target_position = find_position(char, keyboard[list(keyboard.keys())[current_language_index]][list(keyboard[list(keyboard.keys())[current_language_index]].keys())[current_case_index]])
        # 然後按照找到的字符位置進行操作
        dx = target_position[1] - current_position[1]
        dy = target_position[0] - current_position[0]
        for _ in range(abs(dx)):
            key = "Right" if dx > 0 else "Left"
            print(f"Move {key} {abs(dx)} steps")
        for _ in range(abs(dy)):
            key = "Down" if dy > 0 else "Up"
            print(f"Move {key} {abs(dy)} steps")
        print(f"Press key for character {char}")
        current_position = target_position  # 更新當前位置



def replace_voiced(text):
    for voiced, unvoiced in voiced_map.items():
        text = text.replace(voiced, unvoiced)
    return text


def press_key(key):
    win32api.keybd_event(key, 0, 0, 0)
    time.sleep(0.05)  # pause for a bit
    win32api.keybd_event(key, 0, win32con.KEYEVENTF_KEYUP, 0)


# 初始化當前語言頁面
current_page = 0

def input_text(text):
    current_position = (0, 0)  # 初始化當前位置
    current_language_index = 0  # 初始化當前語言索引
    current_case_index = 0  # 初始化當前大小寫或假名類型索引
    for char in text:
        target_position = None
        target_language_index = current_language_index
        target_case_index = current_case_index
        durning_language_count = 0
        durning_case_count = 0
        while target_position is None and durning_language_count < len(keyboard):
            current_language = list(keyboard.keys())[target_language_index]
            current_case = list(keyboard[current_language].keys())[target_case_index]
            target_position = find_position(char, keyboard[current_language][current_case])
            if target_position is None:
                target_case_index = (target_case_index + 1) % len(keyboard[current_language].keys())
                durning_case_count += 1
                if durning_case_count == len(keyboard[current_language].keys()):  # 如果已經遍歷完當前語言的所有頁面
                    target_language_index = (target_language_index + 1) % len(keyboard)
                    target_case_index = 0
                    durning_language_count += 1
                    durning_case_count = 0
                    current_case_index = 0 #換語言要重置頁數
        if target_position is None:
            print(f"Character {char} not found in any language or page.")
            continue
        # 然後按照找到的字符位置進行操作
        for _ in range(abs(target_language_index - current_language_index)):
            print("換語言")
            press_key(0x51)
            current_language_index = target_language_index
        for _ in range(abs(target_case_index - current_case_index)):
            print("翻頁")
            press_key(0x56)
            current_case_index = target_case_index
        dx = target_position[1] - current_position[1]
        dy = target_position[0] - current_position[0]
        for _ in range(abs(dx)):
            key = win32con.VK_RIGHT if dx > 0 else win32con.VK_LEFT
            press_key(key)
            time.sleep(0.05)
        for _ in range(abs(dy)):
            key = win32con.VK_DOWN if dy > 0 else win32con.VK_UP
            press_key(key)
            time.sleep(0.05)
        press_key(0x5A)  # 'Z' 的虛擬鍵碼是 0x5A
        current_position = target_position  # 更新當前位置















time.sleep(5)

print("開始打字")
text = replace_voiced("おんらいんぷれいやーEX")  # 替換濁音和半濁音
print(text)
input_text(text)
