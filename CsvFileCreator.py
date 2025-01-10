import pandas as pd

names = ['name0', 'name1']
email = ['name0@gmail.com', 'name21@gmail.com']
ids = ['10001', '10002']
df = pd.DataFrame()
df['name'] = names
df['email'] = email
df['id'] = ids
df.to_csv('recipients.csv', index=False)