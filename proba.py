s = ['Константин', 'Семён', 'Фёдор', 'Антон', 'Вячеслав']

man=[]
for i in range(len(s)):
    man.append(f"{s[i]} {s[i-1]}ович, возраст - {str(18+i)}")
    print(man[i])
print()
