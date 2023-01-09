
!ABCIncludeFile

  OMIT('_EndOfInclude_',_ODBCConnectionPresent_)
_ODBCConnectionPresent_   EQUATE(1)

  include('odbcTypes.inc'),once
  include('dynstr.inc'),once
  !include('odbcErrorCl.inc'),once  
  include('odbcConnStrCl.inc'),once
  include('OdbcHandleCl.inc'),once
                       
  ITEMIZE(0),PRE(Connection)
Opened equate
CallerOpened  equate
failed equate
  end
  
ODBCConnectionClType class,type,module('odbcConn.clw'),link('odbcConn.clw',_ABCLinkMode_),dll(_ABCDllMode_)
! ODBC handles 
hDbc                   &OdbcHandleClType,private ! db connection handle
hEnv                   &OdbcHandleClType,private ! enviorment handle
hStmt                  &OdbcStmtHandleClType,private! statement handle
!
connStr                &baseConnStrClType,private ! instance to hold the connection string
! some getters, no setters.  Sometimes code will need to query these values
! and use locally, but no code should be allowed to modify
gethDbc                procedure(),SQLHDBC
gethStmt               procedure(),SQLHStmt
gethEnv                procedure(),SQLHEnv
isConnected            procedure(),bool
! set up and clean up
init                   procedure(),sqlReturn
init                   procedure(baseConnStrClType  connString),sqlReturn,proc
kill                   procedure()
! public methods 
! sets the ODBC version is use based on the input
setOdbcVersion         procedure(long verId),sqlReturn,proc 
! connect and disconnect 
connect                procedure(bool statement = withStatement),sqlReturn,proc
Disconnect             procedure(bool statement = withStatement),sqlReturn,proc
! read the errors and show if wanted
getError               procedure(SQLSMALLINT HandleType, SQLHANDLE Handle),virtual,protected
getDatabaseError       procedure(),virtual,protected

                     end ! class 


  _EndOfInclude_
