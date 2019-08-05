import os
import soundfile as sf
import xlwt 
from xlwt import Workbook 

wb = Workbook() 

sheet1 = wb.add_sheet('Calls') 
  
sheet1.write(0, 0, 'Date') 
sheet1.write(0, 1, 'Origin') 
sheet1.write(0, 2, 'Destination') 
sheet1.write(0, 3, 'Duration') 

linha = 1

for dia in os.listdir('audios'):
     for filename in os.listdir('audios/'+dia):
          if(filename != '.DS_Store'):     
               linha += 1
               sheet1.write(linha, 0, filename.split('-')[3][0:4]+'-'+filename.split('-')[3][4:6]+'-'+filename.split('-')[3][6:8]+' '+filename.split('-')[4][0:2]+':'+filename.split('-')[4][2:4]+':'+filename.split('-')[4][4:6])
               sheet1.write(linha, 1, filename.split('-')[2])
               sheet1.write(linha, 2, filename.split('-')[1])
               try:
                    f = sf.SoundFile('audios/'+dia+'/'+filename)
                    sheet1.write(linha, 3, len(f) / f.samplerate)
               except:
                    sheet1.write(linha, 3, 0)
                    print (filename+'\n')

wb.save('calls.xls') 