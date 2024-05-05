import time
#Pegando o tempo de execução do código
tempo_inicial = time.time()
seconds = 0
hour = 0
minutes = 0

print("--- %s segundos ---" % (time.time() - tempo_inicial))
seconds = (time.time() - tempo_inicial)
hour = seconds // 3600
minutes = seconds // 60
seconds %= 60

for i in range(0,1000000):
    i = i+1
    print (i)

print('Fim')

print("--- %s segundos ---" % (time.time() - tempo_inicial))
print("--- %s segundos ---" % ("%d:%02d:%02d" % (hour, minutes, seconds) ))
