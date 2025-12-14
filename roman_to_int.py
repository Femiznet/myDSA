#My First trial on converting roman numerals to integers

#     CONVERTING ROMAN NUMERALS TO INTEGERS

def roman_to_int(roman_num:str):
    roman_num = roman_num.upper() 
    roman_dict = {"I":1,"V":5,"X":10,"L":50,"C":100,"D":500,"M":1000}
    
    length_roman = len(roman_num)
    indx = 0 # index of each numerals in the roman_num
    result = 0 

    # converting each roman_num index by index
    while indx < length_roman:
        current_num = roman_dict[roman_num[indx]] 
        next_indx = indx+1 
        nxt_num = 0 

        # if there is a next index
        if next_indx <= length_roman-1:
            nxt_num = roman_dict[roman_num[next_indx]] 

        # roman numeral calculations
        if current_num < nxt_num:
            result += (nxt_num - current_num)
            indx+=2 # skip the two index worked with
        elif current_num == nxt_num:
            result += (current_num + nxt_num)
            indx+=2
        else:
            result+=current_num
            indx+=1 # Skip only one index if current_num > next_num

    return result


if __name__ == "__main__":
    roman_numerals = ["I", "II", "III", "IV", "V"
                      "X", "L", "C", "D", "M"]
    
    for num in roman_numerals:
        rom_int = roman_to_int(num)
        try:
            assert rom_int in [1, 2, 3, 4, 5, 10, 50, 100, 500, 1000]
            print (rom_int, end=" ")
        except AssertionError:
            print ("Wrong numeral for ", num)
            break
