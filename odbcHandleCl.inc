
!ABCIncludeFile

  OMIT('_EndOfInclude_',_ODBCHandlePresent_)
_ODBCHandlePresent_   EQUATE(1)

  include('odbcTypes.inc'),once
  include('odbcErrorCl.inc'),once  
                       
OdbcHandleClType class,type,module('odbcHandleCl.clw'),link('odbcHandleCl.clw',_ABCLinkMode_),dll(_ABCDllMode_)
handle             SQLHANDLE,private
handleType         SQL_HANDLE_TYPE,private
kill               procedure(),virtual
destruct           procedure()
allocateHandle     procedure(SQL_HANDLE_TYPE hType, SQL_HANDLE_TYPE pType),sqlReturn,proc,virtual
freeHandle         procedure(),sqlReturn,proc,virtual
getHandle          procedure(),SQLHANDLE
                 end

OdbcStmtHandleClType class(OdbcHandleClType),type,module('odbcHandleCl.clw'),link('odbcHandleCl.clw',_ABCLinkMode_),dll(_ABCDllMode_)
freeHandle             procedure(),sqlReturn,proc,virtual
                     end                 
                
  _EndOfInclude_