import os
from threading import Thread
from main import pdf_ex

put_in = 'чертежи\\Чертежи для Алрино'
fil = []
ex = []
ex_out = []
for root, dirs, fiels in os.walk(put_in):
    for i in fiels:
        if '.pdf' in i:
            fil.append(root + '\\' + i)
            put = root + '\\' + i
            put1 = root.split('\\')
            put1[1] = 'excel1'
            pyt_ex = '\\'.join(put1)
            ex.append(pyt_ex)
            ex_out.append(pyt_ex + '\\' + i[:-4])
exz = list(set(ex))
for j in exz:
    os.makedirs(j, exist_ok=True)

print(len(fil))
def many_y(put_in,put_out):
    u = 0
    for in_a, out_a in zip(put_in,put_out):
        u+=1
        print(in_a)
        print(f'{u} из {len(put_in)}')
        pdf_ex(in_a,out_a)
#


many_y(fil,ex_out)

# для запуска в 4 потока
# l = int(len(fil)/4)
# print(l)
# in1 = fil[0:l]
# ex_out1 = ex_out[0:l]
#
# in2 = fil[l:l*2]
# ex_out2 = ex_out[l:l*2]
#
# in3 = fil[l*2:l*3]
# ex_out3 = ex_out[l*2:l*3]
#
# in4 = fil[l*3:len(fil)]
# ex_out4 = ex_out[l*3:len(fil)]
#
#
#
# thread1 = Thread(target=many_y, args=(in1,ex_out1))
# thread2 = Thread(target=many_y, args=(in2,ex_out2))
# thread3 = Thread(target=many_y, args=(in3,ex_out3))
# thread4 = Thread(target=many_y, args=(in4,ex_out4))
#
# thread1.start()
# thread2.start()
# thread3.start()
# thread4.start()
# thread1.join()
# thread2.join()
# thread3.join()
# thread4.join()
