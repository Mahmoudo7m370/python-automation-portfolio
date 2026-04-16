with open('clients.txt','w') as f:
    f.write("mahmoud\n")
    f.write("ahmed\n")
    f.write("mohamed\n")
    f.write("baraa\n")
    f.write("omar\n")
with open('clients.txt','r') as f:
    for line in f:
        print(line)
import csv
with open('clients.csv','w',newline='') as f:
    writer=csv.writer(f)
    writer.writerow(["Name","Email","Amount"])
    writer.writerow(["Ahmed","Ahmed@gmail.com",150])
    writer.writerow(["Mahmoud","Mahmoud@gmail.com",80])
    writer.writerow(["Omar","Omar@gmail.com",100])
with open('clients.csv','r',newline='') as f:
    reader=csv.reader(f)
    next(reader)
    for row in reader:
        if(int(row[2])>100):
            print(row)