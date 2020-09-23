<div align="center">

## Access MDB


</div>

### Description

To open MS Access or Ms SQL server database using C/C++ or VC++
 
### More Info
 
Change the name of the database to any database and place it in the same directory as the .exe. You shoul also change the SQL to your particular database. You may change the connect string to connect to MS SQL Server.

Shows some recordsets

Great


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Mokarrabin A Rahman](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/mokarrabin-a-rahman.md)
**Level**          |Beginner
**User Rating**    |4.6 (51 globes from 11 users)
**Compatibility**  |C, C\+\+ \(general\), Microsoft Visual C\+\+
**Category**       |[Databases](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases__3-5.md)
**World**          |[C / C\+\+](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/c-c.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/mokarrabin-a-rahman-access-mdb__3-1371/archive/master.zip)





### Source Code

```
#import "c:\Program Files\Common Files\System\ADO\msado15.dll" \
no_namespace rename("EOF", "EndOfFile")
// This code comes from : www.geocities.com/mokarrabin
#include <stdio.h>
#include <iostream.h>
void main(void)
{
  CoInitialize(NULL);
  try
  {
  _RecordsetPtr pRst("ADODB.Recordset");
   // Connection String
  _bstr_t strCnn("DRIVER={Microsoft Access Driver (*.mdb)};UID=admin;DBQ=GBOOK.mdb");
	 // Open table
	pRst->Open("SELECT * FROM ProductService where ProductService like '%samir%';", strCnn, adOpenStatic, adLockReadOnly, adCmdText);
	 pRst->MoveFirst();
	 while (!pRst->EndOfFile) {
		 cout<<(char*) ((_bstr_t) pRst->GetFields()->GetItem("ProductService")->GetValue())<<endl;
		 pRst->MoveNext();
	 }
	 pRst->Close();
  }
  catch (_com_error &e)
  {
   cout<<(char*) e.Description();
  }
::CoUninitialize();
}
```

