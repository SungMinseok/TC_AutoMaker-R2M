import pandas as pd
a = "2023-12-28\n05 : 00 :00"
b = '2023-12-25'

a = a.replace('\n', '').replace(' ', '')
print(pd.to_datetime(b, format='%Y-%m-%d%H:%M:%S'))
#print(pd.to_datetime(a))#a = a.replace('\n', '').replace(' ', '')
print(pd.to_datetime(a, format='%Y-%m-%d%H:%M:%S'))

# import pandas as pd

# a = "2023-12-28\n05 : 00 :00"
# b = '2023-12-25'

# # Correct the format in 'a'
# a = a.replace('\n', '').replace(' ', '').replace(':', '')

# print(pd.to_datetime(b))
