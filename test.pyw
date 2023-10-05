data_words = ["human", "i", "am", "a", "human", "i", "human"]
text = ["human", "being"]

foundWords = False

for i in range(len(data_words) - len(text) + 1):
    wordsSet = data_words[i:i+len(text)]
    for i2 in range(len(wordsSet)):
        if text == wordsSet:
            start_index = i
            end_index = i+len(text)-1
            foundWords = True
            break
    if foundWords:
        break

print(data_words[start_index], data_words[end_index])