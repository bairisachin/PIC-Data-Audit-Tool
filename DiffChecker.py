import re
from difflib import ndiff

# pattern = r'[^.\n]'
# class DifferenceFinder:
def check_diff(original_text, modified_text):

    if modified_text is not None and modified_text != '':
        original_text = re.sub(r'[^a-zA-Z0-9\s.\n]', '', original_text)
        original_text = RemoveDotBetweenWords(original_text)
        original_text = original_text.replace('\n', '')


        modified_text = re.sub(r'[^a-zA-Z0-9\s.\n]', '', modified_text)
        modified_text = RemoveDotBetweenWords(modified_text)
        modified_text = modified_text.replace('\n', '')

        diff = list(ndiff(original_text.lower().split(), modified_text.lower().split()))

        for item in diff:
            if item.startswith('-'):
                return False    # Strings are not equal
    
        return True # String are equal
    
    return False

def get_diff(original_text, modified_text):
    if modified_text is not None and modified_text != '':
        original_words = original_text.split()  # Split the original text
        
        original_text = re.sub(r'[^a-zA-Z0-9\s.\n]', '', original_text)
        # ''' INFO: Find all matches of the pattern in the text
        #     regular expression pattern to match numbers with a dot in between
        #     Replace '.' with ' ' only if it's not part of a number with a dot in between'''
        # original_text = re.sub(r'\d+\.\d+', lambda match: match.group().replace('.', ' '), original_text)
        original_text = RemoveDotBetweenWords(original_text)
        original_text = original_text.replace('\n', '')


        modified_text = re.sub(r'[^a-zA-Z0-9\s.\n]', '', modified_text)
        modified_text = RemoveDotBetweenWords(modified_text)
        modified_text = modified_text.replace('\n', '')

        diff = list(ndiff(original_text.lower().split(), modified_text.lower().split()))
        diff_list = []

        for item in diff:
            if item.startswith(' '):
                diff_list.append(item[2:])
            elif item.startswith('-'):
                diff_list.append(item)
            else:
                continue

        output = []
        flag = False

        for i, item in enumerate(diff_list):
            if item.startswith('-'):
                if i == len(diff_list) - 1:
                    if flag:
                        output.append(f"{original_words[i]}]")  # Use original word from original_words list
                    else:
                        output.append(f"[{original_words[i]}]")
                elif diff_list[i + 1].startswith('-'):
                    if flag:
                        output.append(original_words[i])
                    else:
                        output.append(f"[{original_words[i]}")  # Use original word from original_words list
                        flag = True
                elif not diff_list[i + 1].startswith('-'):
                    if not flag:
                        output.append(f"[{original_words[i]}]")  # Use original word from original_words list
                    else:
                        output.append(f"{original_words[i]}]")  # Use original word from original_words list
                        flag = False
                else:
                    output.append(f"{original_words[i]}]")  # Use original word from original_words list
                    flag = False
            else:
                output.append(original_words[i])
        # print(diff_list)
        # print(' '.join(output))
        return ' '.join(output)

    return ""

# Below Func remove the '.' between words and leaves the '.' between numbers
def RemoveDotBetweenWords(text):
    # text = "Animal.Pate is a delicious dish 65.5"
    result = ''

    ignore_dot = False
    for char in text:
        if char == '.' and not ignore_dot:
            result += ' '
        else:
            result += char
        ignore_dot = False
        if char.isdigit():
            ignore_dot = True

    return result

# def check_difference(self, string1, string2):
#     # Split both strings into list items
#     string1 = self.normalize_sentence(string1).split() # Remove Special Chars + Lowercase + strip extra space
#     string2 = self.normalize_sentence(string2).split()

#     A = set(string1) # Store all string1 list items in set A
#     B = set(string2) # Store all string2 list items in set B
    
#     str_diff = A.symmetric_difference(B)
#     return str_diff

# def get_difference(self, string1, string2):
#     # Split both strings into list items
#     string1 = self.normalize_sentence(string1).split() # Remove Special Chars + Lowercase + strip extra space
#     string2 = self.normalize_sentence(string2).split()

#     A = set(string1) # Store all string1 list items in set A
#     B = set(string2) # Store all string2 list items in set B
    
#     uniqueA = A - B
#     uniqueB = B - A
#     uncommon = A.symmetric_difference(B)