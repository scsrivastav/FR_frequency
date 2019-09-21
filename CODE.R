# XLConnect library required

F_sheet = readWorksheetFromFile("E:/Ashish Doc/Farhein/r_import.xlsx", sheet = 1)
fretable = readWorksheetFromFile("E:/Ashish Doc/Farhein/r_import.xlsx", sheet = 1, startRow = 1, endRow = 13)
for(i in 1:12)
 for(j in 2:110)
  fretable[i,j]=0
for (i in 2:length(F_sheet))
{
  for( j in 1:214)
  {
    F_sheet[j,i] = as.character(F_sheet[j,i])
  }
  
}

aflag = c(0,0,0,0,0,0,0,0)

for (i in 2:length(F_sheet))
{ 
  testarray = c("1","2","3","4","5","6","7","8","9","10","11","12")
  ct = c(0,0,0,0,0,0,0,0,0,0,0,0)
  for(j in 1:214)
  { 
    temp = F_sheet[j,i] 
    temp = strsplit(temp,"")[[1]]
    for(k in 1:length(temp))
    {
      for(l in 1:12)
        {
        if(temp[k]==testarray[l])
          ct[l]=ct[l]+1
        }
    }
  }
  for(l in 1:12)
    fretable[l,i]=ct[l]
}
# 
wb = loadWorkbook("E:/Ashish Doc/Farhein/freq_data.xlsx", create = T)
createSheet(wb,name= "freqdata2")
createName(wb, name = "freqdata2", formula = "freqdata2!$C$5")
writeNamedRegion(wb, fretable, name = "freqdata2")
saveWorkbook(wb)
