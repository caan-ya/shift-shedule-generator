import string

def getLettersForRef(sheet_names):
            sheet_name_len = len(sheet_names)
            letters = ""
            out_of_range_count = -1
            while sheet_name_len >= len(string.ascii_uppercase):
               out_of_range_count+=1
               sheet_name_len-=len(string.ascii_uppercase)
            if out_of_range_count > -1:
                letters += string.ascii_uppercase[out_of_range_count]
            letters += string.ascii_uppercase[sheet_name_len]
            return letters