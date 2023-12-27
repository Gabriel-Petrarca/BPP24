from fuzzywuzzy import fuzz

word_to_check = "pharmacist"
string_to_compare = "pharmacuetical tester"

similarity_ratio = fuzz.token_set_ratio(word_to_check, string_to_compare.lower())
print(similarity_ratio)
