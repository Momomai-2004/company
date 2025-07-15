def col_letter_to_index(letter: str) -> int:
    """
    将 Excel 列字母转换为 0 基索引。

    支持多字符列名，例如:
        "A"  -> 0
        "Z"  -> 25
        "AA" -> 26

    Args:
        letter: 列字母字符串

    Returns:
        int: 0 基的列索引
    """
    index = 0
    for char in letter:
        index = index * 26 + (ord(char.upper()) - ord("A") + 1)
    return index - 1

def parse_result_column(result: str) -> str | None:
    """
    从形如 `PN($E*)` 的 Result 字符串中解析出列字母。

    Args:
        result: 规则表 Result 字段字符串

    Returns:
        列字母 (如 'E')；若格式不符合预期则返回 None。
    """
    if not result or "$" not in result or "*" not in result:
        return None
    start = result.find("$") + 1
    end = result.find("*")
    if start >= end:
        return None
    col_letters = result[start:end]
    if not col_letters.isalpha():
        return None
    return col_letters
