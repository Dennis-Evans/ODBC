
!ABCIncludeFile

  OMIT('_EndOfInclude_',_ODBCColsPresent_)
_ODBCColsPresent_   EQUATE(1)

  include('odbcTypes.inc'),once
  include('dynstr.inc'),once
  
! queue to store the columns used in a query
ColumnsQueue queue,type
ColId          long
ColName        string(eColNameLength)
ColType        SQLSMALLINT
ColValue       SQLPOINTER
ColSize        SQLLEN
             end 

eDemoArraySize equate(5)

! array for the bind array and page loaded example
demoArrayG group,type,dim(eDemoArraySize)
fldOne       cstring(10),name('lrs.fldOne')
sizeOne      long   ! used by the driver
fldTwo       long,name('lrs.fldTwo')
sizeTwo      long
fldThree     real,name('lrs.FldThree')
sizeThree    long
fldFour      like(TIMESTAMP_STRUCT),name('lrs.fldFour') 
sizeFour     long
fldFive      long,name('lrs.fldFive')
sizeFive     long
fldSix       cstring(100),name('lrs.fldSix')
sizeSix      long
           end

columnsClass class,type,module('odbcColsCl.clw'),link('odbcColsCl.clw',_ABCLinkMode_),dll(_ABCDllMode_)
setupFailed    bool,private 
colQ           &ColumnsQueue
fieldList      &IDynStr
construct      procedure()
init           procedure(),byte,proc
kill           procedure()
destruct       procedure()
clearQ         procedure()

bindArray      procedure(SQLHSTMT hStmt, demoArrayG grp),sqlReturn
bindColumns    procedure(SQLHSTMT hStmt),sqlReturn
AddColumn      procedure(*long colPtr),sqlReturn,proc
AddColumn      procedure(string colLabel, *long colPtr),sqlReturn,proc
AddColumn      procedure(*string colPtr),sqlReturn,proc     
AddColumn      procedure(string colLabel, *string colPtr),sqlReturn,proc     
AddColumn      procedure(*real colPtr),sqlReturn,proc
AddColumn      procedure(string colLabel, *real colPtr),sqlReturn,proc
AddColumn      procedure(*TIMESTAMP_STRUCT colPtr),sqlReturn,proc
AddColumn      procedure(string colLabel, *TIMESTAMP_STRUCT colPtr),sqlReturn,proc
AddColumn      procedure(SQLSMALLINT TargetType, SQLPOINTER TargetValuePtr, SQLLEN BufferLength),sqlReturn,private,proc
addField       procedure(string colLabel),private
            end 

  _EndOfInclude_
