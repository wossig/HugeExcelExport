# HugeExcelExport
In some case, we need export huge data to excel, it will throw OutOfMemoryException if we use NPOI/EPPlus like these open component. so I did many invetigation from google, we can use Open XML SDK which microsoft provided for generate office file with XML protocal. EPPlus used Open XML SDK too, but it throws OutOfMemoryException.

these code only support to export excel, if you want to load the excel, please use NPOI/EPPlus if excel is not big, the did very well. 
all of these code collected from Internet, please google it. 

Why NPOI/EPPlus will throw memeory exception?
I think they will cache all data into memory, when huge data comes, like (500,000 rows 20+ columns), there is not enough memory to cache the data, so we don't cache them, we use file stream to store it to local disk.
