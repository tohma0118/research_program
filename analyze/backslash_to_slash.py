import pyperclip


def convert_backslash_to_slash(path):
    return path.replace("\\", "/")


# コンソールからアドレスパスを入力
address = input("アドレスパスを入力してください: ")

# バックスラッシュをスラッシュに変換
converted_address = convert_backslash_to_slash(address)

# 変換されたアドレスパスを表示
print("変換後のアドレスパス:", converted_address)

# 変換されたアドレスパスをクリップボードに保存
pyperclip.copy(converted_address)

# クリップボードに保存されたアドレスパスを表示
print("変換後のアドレスパスがクリップボードに保存されました。")
