""" EXAMPLE USAGE
rules = {
    "Numbers" : [
    lambda df: df['field'] > 0,
    lambda df: df['field'] >= 0,
    lambda df: df['field'] < 0,
    lambda df: df['field'] <= 0,
    lambda df: df['field'] != 0,
    lambda df: df['field'] == 0,
],
    "Strings" : [
    lambda df: df['field'] == "string",
    lambda df: df['field'].str.contains("string")
],
    "Combinations" : [
    lambda df: df['field'] != 0 & df['field'] != 1, # AND 
    lambda df: df['field'] != 0 | df['field'] != 1, # OR
    lambda df: df['field'] != 0 ^ df['field'] != 1, # XOR
]}
"""