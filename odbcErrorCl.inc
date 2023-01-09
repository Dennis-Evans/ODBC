
!ABCIncludeFile

  OMIT('_EndOfInclude_',_ODBCErrorClPresent_)
_ODBCErrorClPresent_   EQUATE(1)

  include('odbcTypes.inc'),once
  include('odbcConn.inc'),once
  include ('svcom.inc'),once

    
eLengthSqlState    equate(6)
eLengthErrText     equate(2000)
  
OdbcErrorGroup group,type
sqlState         cstring(eLengthSqlState)
NativeErrorPtr   long
MessageText      cstring(eLengthErrText)
textLengthPtr    SQLSMALLINT
               end 

OdbcErrorQueue queue(OdbcErrorGroup)
               end 
               
ODBCErrorClType class,type,module('odbcErrorCl.clw'),link('odbcErrorCl.clw',_ABCLinkMode_),dll(_ABCDllMode_)
errorMsgQ         &OdbcErrorQueue,private
displayError      bool,private    ! display an error using the message function when it happens
errorCount        long,private    ! number of error records retruned by the most recent error
!
init              procedure(),sqlReturn
kill              procedure(),virtual
destruct          procedure()
!
setDisplayError   procedure(bool onOff)
getConnError      procedure(ODBCConnectionClType conn),sqlReturn,proc 
getEnvError       procedure(ODBCConnectionClType conn),sqlReturn,proc
getDataBaseError  procedure(ODBCConnectionClType conn),sqlReturn,proc
getError          procedure(SQLSMALLINT HandleType, SQLHANDLE Handle),sqlReturn,proc
showError         procedure()
!
freeErrorMsgQ     procedure(),private
getDiagRecCount   procedure(SQLSMALLINT HandleType, SQLHANDLE Handle),long,private,proc
makeObjects       procedure(),sqlReturn,private
destroyObjects    procedure(),private
                end ! class 


  _EndOfInclude_
