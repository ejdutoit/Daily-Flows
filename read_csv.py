import csv

inputFile = open('output.csv')
read = csv.reader(inputFile)
text = ''

for row in read:
    if read.line_num == 243:
        text = row


inputFile.close()

print(text)
print(len(text))

num = ''
for i in range(0,len(text)):
    num = num + text[i]
    print(i, text[i])

print(num)
