def col_letter_to_index(letter: str) -> int:
    index = 0
    for char in letter:
        index = index * 26 + (ord(char.upper()) - ord("A") + 1)
    return index - 1

def parse_result_column(result: str) -> str | None:
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
