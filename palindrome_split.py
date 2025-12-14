# maximum substrings needed to create a palindrome

# For a palindrome to exit, it must have repeating characters

def sub_string(string):
    char_count = {} # re-occuring char

    # sub_str re-occuring char, store in char_count
    for char in string:
        if char not in char_count:
            char_count[char] = 0
        char_count[char]+=1

    # repeated_chars_count // 2 gives the number of substrings needed
    return sum(freq//2 for freq in char_count.values() if freq >= 2)


