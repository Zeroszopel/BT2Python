f = open("C:/Users/Hp/Downloads/input", "r")
global list1


def EvenOdd():
    global odd
    global even
    odd = 0
    even = 0
    for i in list1:
        if i % 2 == 0:
            even += 1
        else:
            odd += 1


def checkElement(e):
    if e < 2:
        return False
    else:
        t = True
        for i in range(2, e):
            if e % i == 0:
                t = True
                break
        return t


def Element():
    global element
    element = 0
    for i in list1:
        if checkElement(i) is True:
            element += 1


def maxcount():
    global count
    count = []
    for i in list1:
        try:
            if count[i] is not None:
                count[i] += 1

        except IndexError:
            while len(count) < i:
                count.append(0)
            count.append(1)
    return count


data = f.read().replace("\n", " ")
list1 = data.split(" ")
for i in range(len(list1)):
    list1[i] = int(list1[i])
EvenOdd()
Element()
maxcount()
print("Số số chẵn là: ", even)
print("Số số lẻ là: ", odd)
print("Số các số nguyên tố là: ", element)
print("Số xuất hiện nhiều nhất là {} với {} lần".format(count.index(max(count)),max(count)))
