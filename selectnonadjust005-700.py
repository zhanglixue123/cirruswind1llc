r = int(input("Iput the date: "))

print('-------------005-700-------------')
for i in range((r-1)*288+2, (r-1)*288+85+1):
    print("=J%d" % (i))

input('Press any key to continue')
print("\n"*100)

for i in range((r-1)*288+2, (r-1)*288+85+1):
    print("=K%d" % (i))


# # summer time in March：
# r = int(input("输入日子："))
#
# print('-------------205-700-------------')
# for i in range((r-1)*288+2-12, (r-1)*288+85+1-12):
#     print("=J%d" % (i))
#
# input('Press any key to continue')
# print("\n"*100)
#
# for i in range((r-1)*288+2-12, (r-1)*288+85+1-12):
#     print("=K%d" % (i))
