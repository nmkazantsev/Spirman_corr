import pandas as pd

df = pd.read_excel("results.xlsx")

# сортировка по времени или нет
time_cmp = False


class row:
    def __init__(self, time, stress):
        self.time = time
        self.stress = stress

    # компаратор для сравнения при сортировке
    def __lt__(self, other):
        if not time_cmp:
            return self.stress > other.stress
        return self.time > other.time


stress = []
# читаем значения стресса
for i in df.get("Столбец 3"):
    # при чтении заголовка не получится его конвертировать в число
    # так что просто пустой обработчик ошибки
    try:
        stress.append(int(i[0:2]))
    except:
        pass

times = []
# читаем значения времени
for i in df.get("Столбец 4"):
    # при чтении заголовка не получится его конвертировать в число
    # так что просто пустой обработчик ошибки
    try:
        h, m = i.split("ч")
        m = m[:-3]  # убрать все буквы (ч и м)
        if m != "":
            times.append(int(h) + int(m) / 60)
        else:
            times.append(int(h))
    except:
        pass

res = []
for i in range(len(times)):
    res.append(row(times[i], stress[i]))

# сортировка резов по уровню стресса по убыванию
res = sorted(res)

# проставляем ранги по стрессу

cur_stress = res[0].stress
start_cur_stress = 0
for i in range(len(res)):

    if cur_stress != res[i].stress:
        num = 0
        summ = 0
        # считаем средний ранг
        # можно без цикла, но мне лень)
        for g in range(start_cur_stress, i):
            num += 1
            summ += (g + 1)

        avg = summ / num

        for g in range(start_cur_stress, i):
            res[g].rang_stress = avg

        cur_stress = res[i].stress
        start_cur_stress = i

# цикл выше не учитывает последний элемент (или несколько, если у них одинаковый стресс), это делаем тут
num = 0
summ = 0
for g in range(start_cur_stress, len(res)):
    num += 1
    summ += (g + 1)

avg = summ / num

for g in range(start_cur_stress, len(res)):
    res[g].rang_stress = avg

# сортировка по времени
time_cmp = True
res = sorted(res)
# проставляем ранги по времени

cur_time = res[0].time
start_cur_time = 0
for i in range(len(res)):

    if cur_time != res[i].time:
        num = 0
        summ = 0
        # считаем средний ранг
        # можно без цикла, но мне лень)
        for g in range(start_cur_time, i):
            num += 1
            summ += (g + 1)

        avg = summ / num

        for g in range(start_cur_time, i):
            res[g].rang_time = avg

        cur_time = res[i].time
        start_cur_time = i

# цикл выше не учитывает последний элемент (или несколько, если у них одинаковый стресс), это делаем тут
num = 0
summ = 0
for g in range(start_cur_time, len(res)):
    num += 1
    summ += (g + 1)

avg = summ / num

for g in range(start_cur_time, len(res)):
    res[g].rang_time = avg

summ = 0
for i in range(len(res)):
    print(res[i].rang_stress, res[i].stress, res[i].rang_time, res[i].time)
    summ += (res[i].rang_stress - res[i].rang_time) ** 2

n = len(res)
print(1 - 6 * summ / n / (n ** 2 - 1))
