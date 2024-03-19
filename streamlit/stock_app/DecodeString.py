def decodeString(s: str) -> str:
    stack = []
    current_num = 0
    current_str = ""

    for char in s:
        if char.isdigit():
            if int(char) < 0:
                pass
            else:
                current_num = current_num * 10 + int(char)
        elif char == "[":
            stack.append((current_str, current_num))
            current_str = ""
            current_num = 0
        elif char == "]":
            prev_str, prev_num = stack.pop()
            current_str = prev_str + current_str * prev_num
        else:
            current_str += char

    return current_str


s1 = "3[a]2[bc]"
print(decodeString(s1))
s2 = "3[a2[c]]"
print(decodeString(s2))
s3 = "2[abc]3[cd]ef"
print(decodeString(s3))
