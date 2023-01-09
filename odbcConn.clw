
  member()
  
  include('odbcConn.inc'),once 
  include'odbcTypes.inc'),once
  include('svcom.inc'),once

  map 
    module('odbc32')
      SQLConnect(SQLHDBC ConnectionHandle, *SQLCHAR ServerName, SQLSMALLINT NameLength1, long UserName, SQLSMALLINT NameLength2, long Authentication, SQLSMALLINT NameLength3),sqlReturn,pascal,raw
      SQLDriverConnect(SQLHDBC ConnectionHandle, SQLHWND WindowHandle, long InConnectionString, SQLSMALLINT StringLength1, long  OutConnectionString, SQLSMALLINT BufferLength, *SQLSMALLINT StringLength2Ptr, SQLUSMALLINT DriverCompletion),sqlReturn,pascal,raw,Name('SQLDriverConnectW')
      SQLDisconnect(SQLHDBC ConnectionHandle),sqlReturn,pascal   
      SQLGetInfo(SQLHDBC hDbc, long attrib, *cstring valuePtr, long buffLength, *long strLenPtr),long,pascal,raw
      SQLSetEnvAttr(SQLHENV EnvironmentHandle, SQLINTEGER Attribute,  SQLPOINTER Value, SQLINTEGER StringLength),sqlReturn,pascal
      SQLSetStmtAttr(SQLHSTMT hStmt, SQLINTEGER attribute, SQLPOINTER value, SQLINTEGER id),long,pascal,Name('SQLSetStmtAttrW')
      SQLSetConnectAttr(SQLHDBC Handle, SQLINTEGER Attribute, SQLPOINTER ValuePtr, SQLINTEGER StringLength),sqlReturn,pascal,name('SQLSetConnectAttrW')
      SQLGetConnectAttr(SQLHDBC Handle, SQLINTEGER Attribute, *SQLPOINTER ValuePtr, SQLINTEGER BufferLength, SQLINTEGER StringLengthPtr),sqlReturn,pascal,name('SQLGetConnectAttrW')
    end 
  end

eNoWindow  equate(0)

ODBCConnectionClType.init procedure() !,sqlReturn

retv    sqlReturn,auto

  code 

  self.hEnv &= new(OdbcHandleClType) 
  if (self.hEnv &= null) 
    return sql_Error
  end
      
  self.hDbc &= new(OdbcHandleClType)
  if (self.hDbc &= null) 
    return sql_Error
  end   

  self.hStmt &= new(OdbcStmtHandleClType)
  if (self.hStmt &= null) 
    return sql_Error
  end   
  
  retv = self.hEnv.allocateHandle(SQL_HANDLE_ENV, Sql_Null_Handle)
  
  return retv
! end init
! -------------------------------------------------------------------------

ODBCConnectionClType.init procedure(baseConnStrClType connString) !,sqlReturn

retv    sqlReturn,auto

  code 

  if (connString &= null) 
    return sql_Error
  end    

  self.connStr &= connString

  retv = self.Init()
    
  return retv
! end init
! -------------------------------------------------------------------------
    
ODBCConnectionClType.kill procedure()

  code 
  
  if (~self.hDbc &= null)
    dispose(self.hdbc)
    self.hDbc &= null
  end 
  
  if (~self.hEnv &= null)
    dispose(self.hEnv)
    self.hEnv &= null
  end 
  
  self.connStr &= null
  if (~self.hStmt &= null)
    dispose(self.hStmt)
    self.hStmt &= null
  end
  
  return 
! end kill 
! -------------------------------------------------------------------------  

ODBCConnectionClType.gethEnv procedure() !,SQLHEnv

  code 
  return self.hEnv.gethandle()
  
ODBCConnectionClType.gethDbc procedure() !,SQLHDBC

dbHandle    SQLHDBC(0)

  code 

  if (~self.hdbc &= null)
    dbHandle = self.hDbc.getHandle()
  end 
  
  return dbHandle
 
ODBCConnectionClType.gethStmt procedure() !,SQLHStmt

  code 
  return self.hStmt.gethandle()
! ---------------------------------------------------------------------------

! ---------------------------------------------------------------------------
! checks to see if the connection is dead or active.  the call to get the 
! connection attribute sets the status field to false if the connection is active
! and to true if the connection is dead, so the return value is set to false 
! at the start and if the connection is not dead then it is set to true
! ---------------------------------------------------------------------------
ODBCConnectionClType.isConnected procedure() !,bool

res      sqlReturn,auto
retv     bool(false)    ! assume the connection is dead or not active at the start
status   long,auto  
dbHandle SQLHDBC,auto

  code

  ! get the current connection handle, check for a value.  
  dbHandle = self.gethDbc()
  ! if > 0 then it has been connected
  if (dbHandle > 0) 
    ! get the staus, if the res is not good then just assume the connection is not active
    res = SQLGetConnectAttr(dbHandle, SQL_ATTR_CONNECTION_DEAD, status, SQL_IS_POINTER, 0)
    if (res = Sql_Success) 
      if (status = SQL_CD_FALSE)
        ! connection is active so return true
        retv = true
      end
    end  ! if (res = Sql_Success) 
  end ! if (dbHandle > 0) 

  return retv
! ---------------------------------------------------------------------------

ODBCConnectionClType.setOdbcVersion procedure(long verId) 

err             ODBCErrorClType
retv            sqlReturn

  code

  retv  = SQLSetEnvAttr(self.gethEnv(), SQL_ATTR_ODBC_VERSION, verId, SQL_IS_INTEGER);
  
  if (retv <> Sql_Success) 
    err.getError(SQL_HANDLE_ENV, self.gethEnv())
  end

  return retv
! --------------------------------------------------------------------------

! --------------------------------------------------------------------------
! connect to the database 
! the output connection is not used but is converted as demo of using wide strings 
! out form API calls.
! 
! if the statement parameter is true, the default, a statement handle is allocated 
! if false then the statement handle will need t obe allocated from the calling code
!
! typically better t ouse the default and just let this block allocate the statement handle
! one less step for the using code
! --------------------------------------------------------------------------
ODBCConnectionClType.connect procedure(bool statement = withStatement)

retv       sqlReturn,auto
outLength  sqlsmallint

! 501 is an arbitray value, the output connection string could be larger
! but we don't use it, so no one cares
outConnStr cstring(501)

wideStr    Cwidestr
outWideStr Cwidestr

outs cstr

  code 
  
  ! if there is already a handle (connect was called twice) just retunr success
  ! consider changing to throw na error, if called twice it is a programming error
  if (self.hdbc.getHandle() <= 0)
    retv = self.hDbc.allocateHandle(SQL_HANDLE_DBC, self.hEnv.getHandle())
  else
    retv = sql_Success
  end

  if (retv = sql_Success) or (retv = sql_success_with_info)
    ! make the connection string a wide string for the ODBC 
    stop(self.connStr.ConnectionString())
    if (wideStr.init(self.connStr.ConnectionString()) = false) 
      return sql_Error
    end
    ! make the output connection string a wide string 
    outConnStr = all(' ')
    if (outWideStr.Init(outConnStr) = false) 
      return sql_Error
    end

    retv = SQLDriverConnect(self.hDbc.getHandle(), eNoWindow, widestr.GetWideStr(), SQL_NTS, outWideStr.GetWideStr(), size(outConnStr), outLength, SQL_DRIVER_NOPROMPT)

    ! ---------------------------------------------------
    ! demo code, show the additional error messages that are avaliable from the driver
    ! ---------------------------------------------------
    !if (retv <> Sql_Success)
    !  self.getDatabaseError()
    !end  
    ! take the output string and convent it to a cla cstring 
    !if (outs.Init(outWideStr) = true)
      !stop(outs.getCStr())
    !end
  end   

  ! check for with info, always returns with info about the connection
  ! because the connection string is modified 
  if (retv <> sql_Success) and (retv <> Sql_Success_With_Info)
    self.getDatabaseError()
  else 
    ! allocate the statement handle if needed/wanted
    if (statement = true) 
      retv = self.hStmt.AllocateHandle(SQL_HANDLE_STMT, self.hDbc.getHandle())
      if (retv <> sql_Success) and (retv <> Sql_Success_With_Info)
        !self.getDatabaseError()
      end   
    end   
  end
 
  if (retv = Sql_Success_With_Info) 
   ! reset for the caller
    retv = sql_Success
  end   
  
  return retv 
! end connect 
! ----------------------------------------------------------------------

! ----------------------------------------------------------------------
! disconnect from the database, freeing all handle in use
! if the dbc handle is not valid then nothing is done
! ----------------------------------------------------------------------
ODBCConnectionClType.Disconnect procedure(bool statement = withStatement)

! assume success at the start
retv      sqlReturn(sql_Success)
h         SQLHDBC,auto

  code 

  if (statement = true) 
    self.hStmt.freeHandle()
  end

  h = self.hDbc.getHandle()
  if (h > 0)
    retv = SQLDisconnect(h)
    ! unlikely that an error will happen here or if it does that any cares 
    ! but grab the error, mainly for development purposes
    if (retv <> sql_Success) and (retv <> sql_Success_With_Info)
      self.getDatabaseError()
    end
    ! 'don't care about info messages here
    if (retv = sql_Success_With_Info)
      ! if with info reset for the caller
      retv = sql_Success  
    end
  end 

  return retv
! end Disconnect 
! ----------------------------------------------------------------------
  
! ----------------------------------------------------------------------
! two function to call the error class if needed
! ----------------------------------------------------------------------
ODBCConnectionClType.getDatabaseError procedure() !,virtual,protected  

err    ODBCErrorClType

  code 
  
  if (err.Init() = sql_success)
    err.getDatabaseError(self)
  end

  return 
! end getDatabaseError
! ----------------------------------------------------------------------

ODBCConnectionClType.getError procedure(SQLSMALLINT HandleType, SQLHANDLE Handle)  

err    ODBCErrorClType

  code 
  
  if (err.init() = sql_success)
    err.getError(handleType, handle)
  end
      
  return 
! end getError 
! ----------------------------------------------------------------------