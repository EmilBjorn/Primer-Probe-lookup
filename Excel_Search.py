# %%
import pandas as pd;
import numpy as np;
import os;
import datetime;

begin_time = datetime.now()


print('Searching...')

files = []
# Mappe som filerne ligger i. Kan godt læse undermapper.  Filerne findes i C:\Results på Hamilton PC'erne
os.chdir(r'S:\TestCenter\MASTERMIX\Udfyldte mastermixskemaer') 

# Genererer en liste over alle filer i mappen
for root, dir, file in os.walk('.'):
    for name in file:
      files.append(os.path.join(root, name))
      
files = files[:1]   # Skærer files-listen af til en enkelt fil mens der testes. 
                    #  Fjern for at løbe gennem alle filer



# %% Loop, som indlæser filerne
for filename in files:
    print(f'Opening {filename}')   
    df = pd.read_excel(filename, sheet_name='Mastermix Luna')
    npdf = df.to_numpy()
    
    print(f'Batch nr RC_{npdf[20,9]} er fremstillet med flg. prober:\n',
          '---------------\n',
    npdf[11,1],'\t\t: ',npdf[11,6],'\n',  # E-Sarbeco F1
    npdf[12,1],'\t\t: ',npdf[12,6],'\n',  # E-Sarbeco R2
    npdf[13,1],'\t: ',npdf[13,6])  # E-Sarbeco P1
    

print('done')
print(f'Search over after {datetime.datetime.now() - begin_time}')
# %%
