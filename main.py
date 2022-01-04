import csv
import numpy as np

with open('C:/Jmeter/apache-jmeter-5.4.1/apache-jmeter-5.4.1/bin/demo1.csv', 'r') as f:
    csv_reader = csv.reader(f, delimiter=",")
    line_count = 0
    for line in csv_reader:
        print(line)
        line_count += 1

print("Total number of line(rows) in the file are = %d"%line_count)

df = pd.read_csv(f)