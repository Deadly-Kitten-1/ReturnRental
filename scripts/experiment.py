import pandas as pd

df = pd.DataFrame(columns=['big', 'small'],
                  data=[[45, 30], [200, 100], [1.5, 1], [30, 20],
                        [250, 150], [1.5, 0.8], [320, 250],
                        [1, 0.8], [0.3, 0.2]])
print(df['big'].tolist())
if 45.0 in df['big'].tolist():
    print('yes')
else:
    print('no')