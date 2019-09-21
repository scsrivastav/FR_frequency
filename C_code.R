F_sheet = readWorksheetFromFile("E:/Ashish Doc/Farhein/r_import.xlsx", sheet = 1)
pt = readWorksheetFromFile("E:/Ashish Doc/Farhein/precise_import.xlsx", sheet = 1)
for (i in 2:length(F_sheet))
{
  for( j in 1:214)
  {
    F_sheet[j,i] = as.character(F_sheet[j,i])
  }
  
}
for (j in 2:101)
  hdr[j-1]=names(pt[j])

hdrf = hdr
hdrf[1]="temp"
for(i in 1:100)
  for(j in 1:12)
  {
    temp = j
    as.character(temp)
    no = 12*(i-1)+j+1
    temp = paste(hdr[i],temp, sep = ".")
    hdrf[no] = temp
  }
for(i in 1:1100)
{
  pt[i+101]=pt[i]
}
colnames(pt)=hdrf
ptc = pt[1:661]
hdrfc=NULL
hdrfc[1]=hdrf[1]
hdrfc[2:661]=hdrf[266:925]
colnames(ptc)=hdrfc

for(i in 1:9)
{
  for(j in 1:7)
  { 
    tempj = j
    as.character(tempj)
    temp = paste(names(F_sheet[i+1]),tempj,sep = ".")
    ptc[7*(i-1)+j,1]=temp
  }
}
ptc[2:661]=0
Aflag = c(1:63)
for (i in 33:87)
{ 
  testarray = c("1","2","3","4","5","6","7","8","9","10","11","12")
  for(j in 1:214)
  { 
    Aflag[1:63]=0
    for(p in 1:9)
    {
      tempA = F_sheet[j,p+1]
      tempA = strsplit(tempA,"")[[1]]
      for(k in 1:length(tempA))
        for(l in 1:7)
        {
          if(tempA[k]==testarray[l])
            Aflag[7*(p-1)+l]=1 
        }
    } 
    temp = F_sheet[j,i] 
    temp = strsplit(temp,"")[[1]]
    for(k in 1:length(temp))
    {  
      if(temp[k]!=",")
      {
        for(l in 1:12)
        { 
          if(temp[k]==testarray[l])
            loopvar = paste(names(F_sheet[i]),temp[k],sep = ".")
        }
        for(p in 1:63)
        {
          if(Aflag[p]==1)
            ptc[p,loopvar]=ptc[p,loopvar]+1
        }
      }  
    }
  }
}

