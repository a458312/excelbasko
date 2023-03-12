import продажи_итог as prod
import продажи_по_дням_итог as dni
import графики as graph
import time
start_time = time.time()
day = '03'
month = '03'
date = day + month
date1 = day + '.' + month
date2 = '03.03'
name = date1 + ' - ' + date2
prod.prod(date)
print("--- %s seconds ---" % (time.time() - start_time))
prod.run_excel('C:/Users/a4583/OneDrive/Desktop/work/продажи.xlsx')
print("--- %s seconds ---" % (time.time() - start_time))
prod.perc(name)
print("--- %s seconds ---" % (time.time() - start_time))
dni.dni(date)
print("--- %s seconds ---" % (time.time() - start_time))
dni.sort(date)
print("--- %s seconds ---" % (time.time() - start_time))
dni.copy(date, date1, date2)
print("--- %s seconds ---" % (time.time() - start_time))
dni.run_excel('C:/Users/a4583/OneDrive/Desktop/work/по дням с ' + date1 + '.xlsx')
print("--- %s seconds ---" % (time.time() - start_time))
graph.graph(date1)
print("--- %s seconds ---" % (time.time() - start_time))
